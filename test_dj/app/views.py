"""
视图层：处理HTTP请求，实现用户认证、对话管理、知识库管理、消息收藏等核心业务。
"""
# 导入Python内置模块
import random
import time
import uuid
import os
import base64
import json
import logging
from datetime import timedelta
from sqlite3 import IntegrityError

# 导入第三方库
import jieba
import dashscope
import chardet
from docx import Document
import win32com.client
import pythoncom
from sentence_transformers import SentenceTransformer, util

# 导入Django核心模块
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect, reverse
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout, get_user_model
from django.core.mail import send_mail
from django.utils import timezone
from django.conf import settings
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.db.models import Q
from django.core.exceptions import ObjectDoesNotExist
from django.views.decorators.http import require_http_methods

# 导入自定义模块
from app.models import VerifyCode, KnowledgeBase, Chat, Message, MessageCollection, CloudModel, ModelCallLog, DoubaoChat, DeepSeekChat, Qwen3_4BChat
from app.services.qwen3_4b_service import local_chat
from app.services.doubao_service import call_doubao
from app.services.deepseek_service import call_deepseek, call_deepseek_v4

# 全局初始化配置
dashscope.base_http_api_url = 'https://dashscope.aliyuncs.com/api/v1'
# DASHSCOPE_API_KEY = os.getenv("DASHSCOPE_API_KEY", "")
DASHSCOPE_API_KEY = "sk-a99ed5ff2a1d461ea6ddf1a9d74f1d8c"
User = get_user_model()

# ========== 云端模型配置解析器 ==========
def parse_cloud_models_from_env():
    """
    从环境变量解析云端模型配置。
    
    配置格式：{MODEL_TYPE}_{MODEL_NAME}_{CONFIG_ITEM}
    
    Returns:
        dict: 按模型类型分组的配置列表
    """
    models = {'doubao': [], 'deepseek': []}
    env_vars = os.environ
    
    # 解析豆包模型配置
    for key, value in env_vars.items():
        if key.startswith('DOUBAO_') and key.endswith('_API_KEY'):
            prefix = key[:-8]  # 移除 '_API_KEY'
            model_info = {
                'name': env_vars.get(f'{prefix}_MODEL_NAME', prefix.replace('DOUBAO_', '').lower()),
                'api_key': value,
                'api_url': env_vars.get(f'{prefix}_API_URL', ''),
                'env_prefix': prefix
            }
            models['doubao'].append(model_info)
    
    # 解析DeepSeek模型配置
    for key, value in env_vars.items():
        if key.startswith('DEEPSEEK_') and key.endswith('_API_KEY'):
            prefix = key[:-8]  # 移除 '_API_KEY'
            model_info = {
                'name': env_vars.get(f'{prefix}_MODEL_NAME', prefix.replace('DEEPSEEK_', '').lower()),
                'api_key': value,
                'api_url': env_vars.get(f'{prefix}_API_URL', ''),
                'env_prefix': prefix
            }
            models['deepseek'].append(model_info)
    
    return models

# 预加载模型配置
CLOUD_MODELS = parse_cloud_models_from_env()


# 用户认证相关视图

@require_http_methods(['GET', 'POST'])
def register_view(request):
    """用户注册视图：处理注册请求，包含表单校验、验证码验证、用户创建"""
    if request.user.is_authenticated:
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': '您已登录，无需重复注册！'})
        return redirect('index')

    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        email = request.POST.get('email', '').strip()
        emailCode = request.POST.get('email_code', '').strip()
        password = request.POST.get('password', '').strip()
        password2 = request.POST.get('password2', '').strip()
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        res = {'status': 'error', 'message': ''}
        
        print(f"\n[注册请求] username={username}, email={email}, emailCode={emailCode}, password_len={len(password)}, is_ajax={is_ajax}")

        # 非空校验
        empty_fields = []
        if not username:
            empty_fields.append('用户名')
        if not email:
            empty_fields.append('邮箱')
        if not password:
            empty_fields.append('密码')
        if not password2:
            empty_fields.append('确认密码')

        if empty_fields:
            res['message'] = f'{empty_fields[0]}不能为空！' if len(empty_fields) == 1 else \
                f'以下字段不能为空：{"、".join(empty_fields)}'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        # 用户名长度校验
        if len(username) < 3 or len(username) > 20:
            res['message'] = '注册失败！用户名长度需在 3-20 位之间'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        # 密码校验
        if len(password) < 6 or len(password) > 14:
            res['message'] = '注册失败！密码长度需在 6-14 位之间'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        if not password.isalnum():
            res['message'] = '注册失败！密码仅支持数字、大（小）写英文字母'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        if password != password2:
            res['message'] = '注册失败！两次输入的密码不一致'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        # 验证码校验
        print(f"[验证码校验] 邮箱: {email}, 验证码: {emailCode}")
        try:
            verify_code = VerifyCode.objects.get(
                email=email,
                code=emailCode,
                is_used=False,
                expire_time__gte=timezone.now()
            )
            print(f"[验证码校验] 成功，标记为已使用")
            verify_code.is_used = True
            verify_code.save()
        except VerifyCode.DoesNotExist:
            code_exist = VerifyCode.objects.filter(email=email, code=emailCode).exists()
            res['message'] = '注册失败！验证码已使用' if code_exist else '注册失败！验证码已过期'
            print(f"[验证码校验] 失败: {res['message']}")
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)
        except Exception as e:
            res['message'] = f'注册失败！验证码校验异常：{str(e)}'
            print(f"[验证码校验] 异常: {str(e)}")
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        # 唯一性校验
        if User.objects.filter(username=username).exists():
            res['message'] = '注册失败！用户名已存在'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        if email != '570510032@qq.com' and User.objects.filter(email=email).exists():
            res['message'] = '注册失败！该邮箱已被注册'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

        # 创建用户
        try:
            user = User.objects.create_user(username=username, email=email, password=password)
            user.save()
            print(f"[注册成功] 用户: {username}, 邮箱: {email}, ID: {user.id}")
            res = {'status': 'success', 'message': '注册成功！即将跳转到登录页...'}
            if is_ajax:
                return JsonResponse(res)
            messages.success(request, res['message'])
            return redirect('login')
        except Exception as e:
            print(f"[注册失败] 错误: {str(e)}")
            res['message'] = '注册失败！服务器内部错误'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/register.html', res)

    return render(request, 'myapp/register.html')


@require_http_methods(['GET', 'POST'])
def login_view(request):
    """用户登录视图：支持账号/邮箱登录，提供记住我功能"""
    if request.user.is_authenticated:
        return redirect('index')

    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '').strip()
        remember = request.POST.get('remember')
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        res = {'status': 'error', 'message': '', 'redirect': ''}

        # 基础校验
        empty_fields = []
        if not username:
            empty_fields.append('账号/邮箱')
        if not password:
            empty_fields.append('密码')

        if empty_fields:
            res['message'] = f'{empty_fields[0]}不能为空！' if len(empty_fields) == 1 else \
                f'以下字段不能为空：{"、".join(empty_fields)}'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/login.html', res)

        # 查询用户
        try:
            user_obj = User.objects.filter(username=username).first() or User.objects.filter(email=username).first()
        except Exception as e:
            res['message'] = '登录失败！服务器内部错误'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/login.html', res)

        if not user_obj:
            res['message'] = '登录失败！账号/邮箱不存在'
            res['redirect'] = reverse('register')
            return JsonResponse(res) if is_ajax else redirect('register')

        # 用户认证
        user = authenticate(request, username=user_obj.username, password=password)
        if not user:
            res['message'] = '登录失败！密码错误'
            return JsonResponse(res) if is_ajax else render(request, 'myapp/login.html', res)

        # 登录成功
        login(request, user)
        user.last_login = timezone.now()
        user.save(update_fields=['last_login'])

        # 记住我功能
        request.session.set_expiry(24 * 60 * 60 if remember == 'on' else 0)

        res['status'] = 'success'
        res['message'] = '登录成功！即将跳转到首页...'
        res['redirect'] = reverse('admin:index') if user.is_staff else reverse('index')
        return JsonResponse(res) if is_ajax else redirect(res['redirect'])

    return render(request, 'myapp/login.html')


@require_http_methods(['POST'])
def verify_email_code(request):
    """校验邮箱验证码：验证验证码是否有效"""
    try:
        email = request.POST.get('email', '').strip()
        code = request.POST.get('code', '').strip()
        if not all([email, code]):
            return JsonResponse({'status': 'error', 'msg': '邮箱和验证码不能为空！'})

        verify_code = VerifyCode.objects.get(
            email=email,
            code=code,
            is_used=False,
            expire_time__gte=timezone.now()
        )
        return JsonResponse({'status': 'success', 'msg': '校验成功！'})

    except VerifyCode.DoesNotExist:
        code_exist = VerifyCode.objects.filter(email=email, code=code).exists()
        if code_exist:
            return JsonResponse({'status': 'error', 'msg': '修改失败！验证码已使用'})
        else:
            return JsonResponse({'status': 'error', 'msg': '修改失败！验证码已过期'})
    except Exception as e:
        return JsonResponse({'status': 'error', 'msg': f'校验失败：{str(e)}'})


@require_http_methods(['POST'])
def send_email_code(request):
    """发送邮箱验证码：支持注册、找回密码、修改密码场景"""
    email = request.POST.get('email', '').strip()
    scene = request.POST.get('scene', '').strip()

    if not email:
        return JsonResponse({'status': 'error', 'msg': '发送失败！邮箱不能为空'}, status=400)
    if not scene:
        return JsonResponse({'status': 'error', 'msg': '发送失败！参数异常'}, status=400)

    # 场景校验
    if scene == "forgot":
        if not User.objects.filter(email=email).exists():
            return JsonResponse({'status': 'error', 'msg': '发送失败！该邮箱未注册'}, status=400)
    elif scene == "change_pwd":
        if not request.user.is_authenticated:
            return JsonResponse({'status': 'error', 'msg': '发送失败！用户未登录'}, status=400)
        if email != request.user.email:
            return JsonResponse({'status': 'error', 'msg': '发送失败！验证码将发送至您绑定的邮箱'}, status=400)
    elif scene != "register":
        return JsonResponse({'status': 'error', 'msg': '发送失败！无效操作'}, status=400)

    # 生成验证码
    code = ''.join(random.choices('0123456789', k=6))
    expire_time = timezone.now() + timedelta(minutes=5)
    VerifyCode.objects.filter(email=email, is_used=False).delete()
    user_obj = User.objects.filter(email=email).first()
    VerifyCode.objects.create(email=email, code=code, expire_time=expire_time, user=user_obj)

    # 发送邮件
    try:
        send_mail(
            subject="【基于大语言模型的智能问答系统】验证码",
            message=f"您当前操作的验证码是：{code}，5分钟内有效，请及时验证！",
            from_email=settings.DEFAULT_FROM_EMAIL,
            recipient_list=[email],
            fail_silently=False,
        )
        return JsonResponse({'status': 'success', 'msg': '验证码已发送到您的邮箱！'})
    except Exception as e:
        return JsonResponse({'status': 'error', 'msg': f'发送失败！{str(e)}'}, status=500)


@require_http_methods(['GET', 'POST'])
def forgot_pwd_view(request):
    """找回密码视图：验证验证码并重置密码"""
    if request.method == 'POST':
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        email = request.POST.get('email', '').strip()
        code = request.POST.get('code', '').strip()
        new_password = request.POST.get('new_password', '').strip()
        new_password2 = request.POST.get('new_password2', '').strip()

        if not all([email, code, new_password, new_password2]):
            res = {'status': 'error', 'msg': '修改失败！所有字段不能为空'}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/forgot_pwd.html', res)

        if new_password != new_password2:
            res = {'status': 'error', 'msg': '修改失败！两次密码不一致'}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/forgot_pwd.html', res)

        # 验证码校验
        try:
            verify_code = VerifyCode.objects.get(
                email=email,
                code=code,
                is_used=False,
                expire_time__gte=timezone.now()
            )
        except VerifyCode.DoesNotExist:
            code_exist = VerifyCode.objects.filter(email=email, code=code).exists()
            msg = '修改失败！验证码已使用' if code_exist else '修改失败！验证码已过期'
            res = {'status': 'error', 'msg': msg}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/forgot_pwd.html', res)

        # 密码校验
        if len(new_password) < 6 or len(new_password) > 14:
            res = {'status': 'error', 'msg': '修改失败！密码长度需在 6-14 位之间'}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/forgot_pwd.html', res)

        if not new_password.isalnum():
            res = {'status': 'error', 'msg': '修改失败！密码仅支持数字、大（小）写英文字母'}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/forgot_pwd.html', res)

        # 更新密码状态
        verify_code.is_used = 1
        verify_code.save(update_fields=['is_used'])

        res = {'status': 'success', 'msg': '修改成功！请重新登录', 'redirect': reverse('login')}
        return JsonResponse(res) if is_ajax else redirect('login')

    return render(request, 'myapp/forgot_pwd.html')


@login_required(login_url='login')
def index_view(request):
    """首页视图：渲染模型选择页面"""
    context = {
        'user_id': request.user.id,
        'username': request.user.username,
        'email': request.user.email,
        'avatar_url': request.user.avatar_url,
        'message': "欢迎使用智能对话系统！"
    }
    return render(request, 'myapp/index.html', context)


@login_required(login_url='login')
@require_http_methods(['GET'])
def chat_page_view(request):
    """经典对话页面视图：渲染聊天界面"""
    context = {
        'user_id': request.user.id,
        'username': request.user.username,
        'email': request.user.email,
        'avatar_url': request.user.avatar_url,
        'message': "欢迎登录！"
    }
    return render(request, 'myapp/chat.html', context)


@login_required(login_url='login')
@require_http_methods(['GET'])
def qwen3_4b_page_view(request):
    """Qwen3-4B模型页面视图"""
    # 单一本地模型配置
    context = {
        'user_id': request.user.id,
        'username': request.user.username,
        'email': request.user.email,
        'avatar_url': request.user.avatar_url,
        'message': "欢迎使用Qwen3-4B本地模型！"
    }
    return render(request, 'myapp/qwen3_4b.html', context)


@login_required(login_url='login')
@require_http_methods(['GET'])
def doubao_page_view(request):
    """豆包模型专属页面视图"""
    # 获取所有豆包模型配置
    doubao_models = CLOUD_MODELS.get('doubao', [])
    # 如果没有配置，使用默认值
    if not doubao_models:
        doubao_models = [{
            'name': 'doubao-seed-2-0-pro-260215',
            'api_key': os.getenv('DOUBAO_API_KEY', ''),
            'api_url': os.getenv('DOUBAO_API_URL', ''),
            'env_prefix': 'DOUBAO'
        }]
    
    context = {
        'user_id': request.user.id,
        'username': request.user.username,
        'email': request.user.email,
        'avatar_url': request.user.avatar_url,
        'message': "欢迎使用豆包云端模型！",
        'models': doubao_models,
        'default_model': doubao_models[0]['name'] if doubao_models else ''
    }
    return render(request, 'myapp/doubao.html', context)


@login_required(login_url='login')
@require_http_methods(['GET'])
def deepseek_page_view(request):
    """DeepSeek模型专属页面视图"""
    # 获取所有DeepSeek模型配置
    deepseek_models = CLOUD_MODELS.get('deepseek', [])
    # 如果没有配置，使用默认值
    if not deepseek_models:
        deepseek_models = [{
            'name': 'deepseek-v4-flash',
            'api_key': os.getenv('DEEPSEEK_V4_API_KEY', ''),
            'api_url': os.getenv('DEEPSEEK_V4_API_URL', ''),
            'env_prefix': 'DEEPSEEK_V4'
        }]
    
    context = {
        'user_id': request.user.id,
        'username': request.user.username,
        'email': request.user.email,
        'avatar_url': request.user.avatar_url,
        'message': "欢迎使用DeepSeek云端模型！",
        'models': deepseek_models,
        'default_model': deepseek_models[0]['name'] if deepseek_models else ''
    }
    return render(request, 'myapp/deepseek.html', context)


@login_required(login_url='login')
@require_http_methods(['POST'])
def cloud_model_chat_api(request):
    """云端模型聊天API：支持豆包、DeepSeek调用，集成RAG检索"""
    try:
        data = json.loads(request.body)
        message = data.get('message')
        model_type = data.get('model_type', 'doubao')
        model_name = data.get('model_name')
        model_config = data.get('model_config', {})
        rag_config = data.get('rag_config', {})
        context_messages = data.get('context_messages', None)
        kb_id = data.get('kb_id', 'default')

        if not message:
            return JsonResponse({'code': 400, 'msg': '请输入问题！', 'data': {}})

        # ==================== RAG检索流程 ====================
        # 1. 接收用户查询后，先对查询进行预处理和向量化
        # 2. 使用向量化后的查询在知识库中进行相似度检索，获取相关文档片段
        # 3. 将检索到的文档片段与原始查询进行融合，作为上下文信息传递给模型
        
        query_with_context = message
        
        if kb_id != 'default':
            try:
                kb = KnowledgeBase.objects.get(kb_id=kb_id, user=request.user)
                # 对知识库内容进行分段处理
                chunks = split_text_by_chunk(kb.content, rag_config.chunk_size, rag_config.chunk_overlap)
                if chunks:
                    chunk_scores = []
                    # 对每个片段计算与查询的相似度得分
                    for chunk in chunks:
                        score = calculate_rag_score(message, chunk, rag_config)
                        if score >= rag_config.similarity_threshold:
                            chunk_scores.append((chunk, score))
                    # 按得分排序，取Top-N相关片段
                    chunk_scores.sort(key=lambda x: x[1], reverse=True)
                    top_chunks = [chunk for chunk, score in chunk_scores[:rag_config.top_n]]
                    if top_chunks:
                        # 将检索到的文档片段与原始查询融合
                        query_with_context = f"""以下是参考知识库的相关片段：
{chr(10).join([f"片段{idx + 1}：{chunk}" for idx, chunk in enumerate(top_chunks)])}

请基于上述片段回答用户问题，优先使用知识库信息，若片段无相关内容可结合自身知识回答，回答需简洁准确：
用户问题：{message}"""
                    else:
                        # 降级处理：知识库中无相关内容
                        query_with_context = f"知识库中无相关内容，用户问题：{message}"
            except KnowledgeBase.DoesNotExist:
                pass

        result_data = {}

        if model_type == 'doubao':
            # 获取豆包模型配置
            doubao_models = CLOUD_MODELS.get('doubao', [])
            # 根据model_name查找对应的配置
            model_config_found = None
            for m in doubao_models:
                if m['name'] == model_name:
                    model_config_found = m
                    break
            
            # 如果没找到，使用默认配置
            if not model_config_found:
                model_config_found = {
                    'name': os.getenv('DOUBAO_MODEL_NAME', 'doubao-seed-2-0-pro-260215'),
                    'api_key': os.getenv('DOUBAO_API_KEY', ''),
                    'api_url': os.getenv('DOUBAO_API_URL', 'https://ark.cn-beijing.volces.com/api/v3/chat/completions'),
                    'env_prefix': 'DOUBAO'
                }
            
            result, total_time_ms, ttft_ms, token_count = call_doubao(
                query_with_context,
                model_config,
                model_config_found['api_url'],
                model_config_found['name'],
                context_messages,
                model_config_found['api_key']
            )
            result_data = {
                'result': result,
                'total_time_ms': total_time_ms,
                'ttft_ms': ttft_ms,
                'token_count': token_count,
                'model_name': f'豆包 ({model_config_found["name"]})'
            }
        elif model_type == 'deepseek':
            # 获取DeepSeek模型配置
            deepseek_models = CLOUD_MODELS.get('deepseek', [])
            # 根据model_name查找对应的配置
            model_config_found = None
            for m in deepseek_models:
                if m['name'] == model_name:
                    model_config_found = m
                    break
            
            # 如果没找到，使用默认配置
            if not model_config_found:
                model_config_found = {
                    'name': os.getenv('DEEPSEEK_V4_MODEL_NAME', 'deepseek-v4-flash'),
                    'api_key': os.getenv('DEEPSEEK_V4_API_KEY', ''),
                    'api_url': os.getenv('DEEPSEEK_V4_API_URL', 'https://api.deepseek.com/v1/chat/completions'),
                    'env_prefix': 'DEEPSEEK_V4'
                }
            
            result, total_time_ms, ttft_ms, token_count = call_deepseek_v4(
                query_with_context,
                model_config,
                model_config_found['api_url'],
                model_config_found['name'],
                context_messages,
                model_config_found['api_key']
            )
            result_data = {
                'result': result,
                'total_time_ms': total_time_ms,
                'ttft_ms': ttft_ms,
                'token_count': token_count,
                'model_name': f'DeepSeek ({model_config_found["name"]})'
            }
        else:
            return JsonResponse({'code': 400, 'msg': '不支持的模型类型！', 'data': {}})

        ModelCallLog.objects.create(
            user=request.user,
            model_name=result_data.get('model_name', 'Unknown'),
            call_mode='cloud',
            input_text=message[:500],
            output_text=result_data['result'][:500] if result_data['result'] else '',
            total_time_ms=result_data['total_time_ms'],
            ttft_ms=result_data['ttft_ms'],
            token_count=result_data['token_count'],
            success=not result_data['result'].startswith('错误')
        )

        return JsonResponse({'code': 200, 'msg': 'success', 'data': result_data})

    except Exception as e:
        logger.error(f"[云端模型聊天错误] {str(e)}", exc_info=True)
        return JsonResponse({'code': 500, 'msg': f'服务器错误: {str(e)}', 'data': {}})


@login_required(login_url='login')
@require_http_methods(['GET'])
def logout_view(request):
    """退出登录视图：清除登录状态并跳转登录页"""
    logout(request)
    messages.success(request, "已退出登录！")
    return redirect('login')

# 核心业务接口 - 对话与消息管理

@login_required(login_url='login')
@require_http_methods(['GET'])
def get_user_knowledge_bases(request):
    """获取用户的知识库列表"""
    try:
        kbs = KnowledgeBase.objects.filter(user=request.user)
        kb_list = [{
            'kb_id': kb.kb_id,
            'name': kb.name,
            'content': kb.content,
            'chunk_size': kb.chunk_size,
            'chunk_overlap': kb.chunk_overlap,
            'similarity_threshold': kb.similarity_threshold,
            'top_n': kb.top_n
        } for kb in kbs]
        return JsonResponse({'code': 200, 'msg': 'success', 'data': kb_list})
    except Exception as e:
        return JsonResponse({'code': 500, 'msg': f'获取失败：{str(e)}', 'data': []})


@login_required(login_url='login')
@require_http_methods(['POST'])
def get_chat_sessions(request):
    """获取用户历史对话列表"""
    try:
        chat_sessions = Chat.objects.filter(user=request.user).order_by('-create_time')
        data = [{
            "chat_id": session.chat_id,
            "chat_title": session.chat_title,
            "create_time": timezone.localtime(session.create_time).strftime('%Y-%m-%d %H:%M:%S'),
            "update_time": timezone.localtime(session.update_time).strftime('%Y-%m-%d %H:%M:%S'),
            "model_name": session.model_name
        } for session in chat_sessions]
        return JsonResponse({'code': 200, 'msg': '加载成功', 'data': data})
    except Exception as e:
        print(f"加载历史对话失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '加载历史对话失败', 'data': []})


@login_required(login_url='login')
@require_http_methods(['POST'])
def get_chat_sessions_by_model(request):
    """获取指定模型的对话列表"""
    try:
        data = json.loads(request.body)
        model_name = data.get('model_name')
        model_type = data.get('model_type', 'doubao')
        
        if not model_name:
            return JsonResponse({'code': 400, 'msg': '缺少模型名称参数', 'data': []})
        
        if model_type == 'doubao':
            model_chats = DoubaoChat.objects.filter(
                user=request.user, 
                model_name=model_name
            ).select_related('chat').order_by('create_time')
            chat_sessions = [mc.chat for mc in model_chats]
        elif model_type == 'deepseek':
            model_chats = DeepSeekChat.objects.filter(
                user=request.user, 
                model_name=model_name
            ).select_related('chat').order_by('create_time')
            chat_sessions = [mc.chat for mc in model_chats]
        elif model_type == 'local':
            model_chats = Qwen3_4BChat.objects.filter(
                user=request.user, 
                model_name=model_name
            ).select_related('chat').order_by('create_time')
            chat_sessions = [mc.chat for mc in model_chats]
        else:
            chat_sessions = Chat.objects.filter(
                user=request.user, 
                model_name=model_name
            ).order_by('-update_time')
        
        data = [{
            "chat_id": session.chat_id,
            "chat_title": session.chat_title,
            "create_time": timezone.localtime(session.create_time).strftime('%Y-%m-%d %H:%M:%S'),
            "update_time": timezone.localtime(session.update_time).strftime('%Y-%m-%d %H:%M:%S'),
            "model_name": session.model_name
        } for session in chat_sessions]
        
        return JsonResponse({'code': 200, 'msg': '加载成功', 'data': data})
    except Exception as e:
        print(f"加载模型对话失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '加载模型对话失败', 'data': []})


@login_required(login_url='login')
@require_http_methods(['POST'])
def get_chat_messages(request):
    """获取指定对话的消息列表"""
    try:
        # print('=== [后端调试] get_chat_messages 开始 ===')
        data = json.loads(request.body)
        chat_id = data.get('chat_id')
        call_mode = data.get('call_mode')
        cloud_model_id = data.get('cloud_model_id')
        
        #print(f'请求参数 - chat_id: {chat_id}, call_mode: {call_mode}, cloud_model_id: {cloud_model_id}')
        
        if not chat_id:
            return JsonResponse({'code': 400, 'msg': '缺少对话ID', 'data': []})

        chat = Chat.objects.get(chat_id=chat_id, user=request.user)
        #print(f'找到对话: {chat.chat_id}, 标题: {chat.chat_title}')
        
        messages_query = Message.objects.filter(chat=chat, chat__user=request.user)
        #print(f'过滤前消息数量: {messages_query.count()}')
        
        if call_mode:
            messages_query = messages_query.filter(call_mode=call_mode)
            print(f'call_mode={call_mode} 过滤后消息数量: {messages_query.count()}')
        # for msg in messages_query.values('msg_id', 'role', 'cloud_model_id'):
        #         print(f' 过滤前： msg_id={msg["msg_id"]}, role={msg["role"]}, cloud_model_id="{msg["cloud_model_id"]}", 类型={type(msg["cloud_model_id"])}')
        if cloud_model_id:
            messages_query = messages_query.filter(
                Q(role=1, cloud_model_id=cloud_model_id) | Q(role=2, cloud_model_id=cloud_model_id)
            )
            # for msg in messages_query.values('msg_id', 'role', 'cloud_model_id'):
            #     print(f' 过滤后： msg_id={msg["msg_id"]}, role={msg["role"]}, cloud_model_id="{msg["cloud_model_id"]}", 类型={type(msg["cloud_model_id"])}')
            # print(f'cloud_model_id={cloud_model_id} 过滤后消息数量: {messages_query.count()}')
        
        messages = messages_query.order_by('send_time')

        # 模型ID到名称的映射表（从环境变量加载）
        def get_cloud_model_name(model_id):
            if not model_id:
                return None
            import os
            model_name_map = {
                'doubao': os.getenv('DOUBAO_MODEL_NAME', '豆包'),
                'deepseek': os.getenv('DEEPSEEK_MODEL_NAME', 'DeepSeek'),
                'deepseek-v4': os.getenv('DEEPSEEK_V4_MODEL_NAME', 'DeepSeek-V4-Flash')
            }
            return model_name_map.get(model_id)

        msg_list = [{
            'msg_id': msg.msg_id,
            'role': msg.role,
            'content': msg.content,
            'time': timezone.localtime(msg.send_time).strftime('%H:%M'),
            'round_num': msg.round_num,
            'model_type': msg.model_type,
            'cloud_model_id': msg.cloud_model_id,
            'call_mode': msg.call_mode,
            'model_name': get_cloud_model_name(msg.cloud_model_id) if msg.role == 2 and msg.model_type == 'cloud' else None,
            'performance': {
                'total_time_ms': msg.total_time_ms,
                'ttft_ms': msg.ttft_ms,
                'token_count': msg.token_count
            } if msg.role == 2 else None
        } for msg in messages]

        return JsonResponse({
            'code': 200,
            'msg': '加载成功',
            'data': msg_list,
            'memory_rounds': chat.memory_rounds
        })
    except Chat.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '对话不存在', 'data': []})
    except Exception as e:
        print(f"加载对话消息失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '加载对话消息失败', 'data': []})


@login_required(login_url='login')
@require_http_methods(['POST'])
def create_chat_session(request):
    """创建新对话会话"""
    try:
        data = json.loads(request.body) if request.body else {}
        model_type = data.get('model_type')
        model_name = data.get('model_name')
        
        new_chat = Chat.objects.create(user=request.user, chat_title="新对话")
        
        # 如果提供了模型信息，创建模型关联记录
        if model_type and model_name:
            try:
                if model_type == 'doubao':
                    DoubaoChat.objects.create(
                        user=request.user,
                        chat=new_chat,
                        model_name=model_name
                    )
                elif model_type == 'deepseek':
                    DeepSeekChat.objects.create(
                        user=request.user,
                        chat=new_chat,
                        model_name=model_name
                    )
                elif model_type == 'local':
                    Qwen3_4BChat.objects.create(
                        user=request.user,
                        chat=new_chat,
                        model_name=model_name
                    )
            except Exception as e:
                print(f"创建模型关联记录失败：{str(e)}")
        
        format_time = timezone.localtime(new_chat.create_time).strftime('%Y-%m-%d %H:%M:%S')
        return JsonResponse({
            'code': 200,
            'msg': '新建对话成功',
            'data': {
                'chat_id': new_chat.chat_id,
                'chat_title': new_chat.chat_title,
                'create_time': format_time
            }
        })
    except Exception as e:
        print(f"新建对话失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '新建对话失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def set_chat_memory_rounds(request):
    """设置对话上下文记忆轮数"""
    try:
        data = json.loads(request.body)
        chat_id = data.get('chat_id')
        memory_rounds = data.get('memory_rounds', 0)

        if not chat_id or not isinstance(memory_rounds, int) or memory_rounds < 0:
            return JsonResponse({'code': 400, 'msg': '参数错误（记忆轮数≥0）', 'data': None})

        chat = Chat.objects.get(chat_id=chat_id, user=request.user)
        chat.memory_rounds = memory_rounds
        chat.save()

        return JsonResponse({'code': 200, 'msg': '设置成功', 'data': {'memory_rounds': chat.memory_rounds}})
    except Chat.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '对话不存在', 'data': None})
    except Exception as e:
        print(f"设置记忆轮数失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '设置失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def save_chat_message(request):
    """保存对话消息到数据库，计算会话轮次"""
    try:
        data = json.loads(request.body)
        required_params = ['chat_id', 'role', 'content']
        for param in required_params:
            if not data.get(param):
                return JsonResponse({'code': 400, 'msg': f'缺少参数：{param}', 'data': None})

        role = data['role']
        chat = Chat.objects.get(chat_id=data['chat_id'], user=request.user)

        # 获取当前调用模式
        call_mode = data.get('call_mode', 'local')

        # 计算当前轮次（按模式独立计算）
        all_messages = Message.objects.filter(chat=chat, call_mode=call_mode).order_by('send_time')
        current_round = 1
        if all_messages.exists():
            last_round = all_messages.last().round_num
            current_round = last_round + 1 if role == 'user' else last_round

        # 转换角色标识
        role_int = 1 if role == 'user' else 2

        # 判断是否超出记忆轮数（按模式独立判断）
        memory_rounds = chat.memory_rounds
        new_topic_triggered = memory_rounds > 0 and current_round > memory_rounds

        # 保存消息
        cloud_model_id_value = data.get('cloud_model_id') or data.get('model_name')
        print(f'保存消息 - cloud_model_id: {cloud_model_id_value}, model_name: {data.get("model_name")}')
        new_msg = Message.objects.create(
            chat=chat,
            user=request.user,
            role=role_int,
            content=data['content'],
            send_time=timezone.now(),
            round_num=current_round,
            model_type=data.get('model_type'),
            cloud_model_id=cloud_model_id_value,
            call_mode=call_mode,
            total_time_ms=data.get('total_time_ms'),
            ttft_ms=data.get('ttft_ms'),
            token_count=data.get('token_count')
        )
        print(f'消息已保存 - msg_id: {new_msg.msg_id}, cloud_model_id: {new_msg.cloud_model_id}')

        # 更新对话信息
        chat.update_time = timezone.now()
        if role == 'user' and chat.chat_title == "新对话":
            chat.chat_title = data['content'][:20] + '...' if len(data['content']) > 20 else data['content']
        
        # 更新对话的模型名称（如果提供了）
        if data.get('model_name'):
            chat.model_name = data['model_name']
        
        chat.save()
        
        # 根据模型类型创建或更新模型专属对话关联表记录
        model_name = data.get('model_name')
        model_type = data.get('model_type', 'local')
        cloud_model_id = data.get('cloud_model_id')
        
        if model_name and role == 'user':
            try:
                # 根据model_type或cloud_model_id判断具体模型类型
                if model_type == 'doubao' or (model_type == 'cloud' and cloud_model_id and 'doubao' in cloud_model_id.lower()):
                    DoubaoChat.objects.update_or_create(
                        user=request.user,
                        chat=chat,
                        model_name=model_name,
                        defaults={'create_time': timezone.now()}
                    )
                elif model_type == 'deepseek' or (model_type == 'cloud' and cloud_model_id and 'deepseek' in cloud_model_id.lower()):
                    DeepSeekChat.objects.update_or_create(
                        user=request.user,
                        chat=chat,
                        model_name=model_name,
                        defaults={'create_time': timezone.now()}
                    )
                elif model_type == 'local' or (model_type == 'cloud' and cloud_model_id and ('qwen' in cloud_model_id.lower() or 'local' in cloud_model_id.lower())):
                    Qwen3_4BChat.objects.update_or_create(
                        user=request.user,
                        chat=chat,
                        model_name=model_name,
                        defaults={'create_time': timezone.now()}
                    )
                elif model_type == 'cloud' and not cloud_model_id:
                    # 如果是云端模式但没有指定具体模型，根据model_name判断
                    if model_name and ('豆包' in model_name or 'doubao' in model_name.lower()):
                        DoubaoChat.objects.update_or_create(
                            user=request.user,
                            chat=chat,
                            model_name=model_name,
                            defaults={'create_time': timezone.now()}
                        )
                    elif model_name and ('deepseek' in model_name.lower()):
                        DeepSeekChat.objects.update_or_create(
                            user=request.user,
                            chat=chat,
                            model_name=model_name,
                            defaults={'create_time': timezone.now()}
                        )
            except Exception as e:
                print(f"更新模型专属对话关联表失败：{str(e)}")

        format_send_time = timezone.localtime(new_msg.send_time).strftime('%Y-%m-%d %H:%M:%S')
        return JsonResponse({
            'code': 200,
            'msg': '保存消息成功',
            'data': {
                'msg_id': new_msg.msg_id,
                'send_time': format_send_time,
                'current_round': current_round,
                'memory_rounds': memory_rounds,
                'new_topic_triggered': new_topic_triggered
            }
        })
    except Chat.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '对话不存在或无操作权限', 'data': None})
    except Exception as e:
        print(f"保存消息失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '保存消息失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def delete_chat(request):
    """删除指定对话，级联删除关联消息和收藏"""
    try:
        data = json.loads(request.body)
        chat_id = data.get('chat_id')
        if not chat_id:
            return JsonResponse({'code': 400, 'msg': '缺少对话ID', 'data': None})

        chat = Chat.objects.get(chat_id=chat_id, user=request.user)
        chat.delete()

        return JsonResponse({'code': 200, 'msg': '对话删除成功', 'data': None})
    except Chat.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '对话不存在或无操作权限', 'data': None})
    except Exception as e:
        print(f"删除对话失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '删除失败，请重试', 'data': None})


# 消息收藏相关接口

@login_required(login_url='login')
@require_http_methods(['POST'])
def collect_message(request):
    """收藏消息"""
    try:
        data = json.loads(request.body)
        msg_id = data.get('msg_id')
        if not msg_id:
            return JsonResponse({'code': 400, 'msg': '缺少消息ID', 'data': None})

        message = Message.objects.get(msg_id=msg_id)
        if message.chat.user != request.user:
            return JsonResponse({'code': 403, 'msg': '无权操作该消息', 'data': None})

        collection, created = MessageCollection.objects.get_or_create(
            user=request.user,
            message=message
        )

        return JsonResponse({
            'code': 200,
            'msg': '收藏成功' if created else '已收藏该消息，无需重复操作',
            'data': {'collected': True}
        })
    except Message.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '消息不存在或无操作权限', 'data': None})
    except IntegrityError:
        return JsonResponse({'code': 400, 'msg': '已收藏该消息', 'data': {'collected': True}})
    except Exception as e:
        print(f"收藏消息失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '收藏失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def cancel_collect_message(request):
    """取消收藏消息"""
    try:
        data = json.loads(request.body)
        msg_id = data.get('msg_id')
        if not msg_id:
            return JsonResponse({'code': 400, 'msg': '缺少消息ID', 'data': None})

        message = Message.objects.get(msg_id=msg_id)
        if message.chat.user != request.user:
            return JsonResponse({'code': 403, 'msg': '无权操作该消息', 'data': None})

        deleted_count, _ = MessageCollection.objects.filter(user=request.user, message=message).delete()

        return JsonResponse({
            'code': 200,
            'msg': '取消收藏成功' if deleted_count > 0 else '未收藏该消息，无需操作',
            'data': {'collected': False}
        })
    except Message.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '消息不存在或无操作权限', 'data': None})
    except Exception as e:
        print(f"取消收藏失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '取消收藏失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def check_message_collected(request):
    """检查消息是否收藏"""
    try:
        data = json.loads(request.body)
        msg_id = data.get('msg_id')
        if not msg_id:
            return JsonResponse({'code': 400, 'msg': '缺少消息ID', 'data': None})

        message = Message.objects.get(msg_id=msg_id)
        if message.chat.user != request.user:
            return JsonResponse({'code': 403, 'msg': '无权操作该消息', 'data': None})

        collected = MessageCollection.objects.filter(user=request.user, message=message).exists()
        return JsonResponse({'code': 200, 'msg': '查询成功', 'data': {'collected': collected}})
    except Message.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '消息不存在或无操作权限', 'data': None})
    except Exception as e:
        print(f"查询收藏状态失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '查询失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['GET'])
def get_collected_messages(request):
    """获取用户收藏的消息列表"""
    try:
        collections = MessageCollection.objects.filter(user=request.user) \
            .select_related('message', 'message__chat') \
            .order_by('-collect_time')

        collected_msgs = [{
            "collection_id": col.collection_id,
            "msg_id": col.message.msg_id,
            "role": col.message.role,
            "content": col.message.content,
            "send_time": timezone.localtime(col.message.send_time).strftime('%Y-%m-%d %H:%M:%S'),
            "collect_time": timezone.localtime(col.collect_time).strftime('%Y-%m-%d %H:%M:%S'),
            "chat_id": col.message.chat.chat_id,
            "chat_title": col.message.chat.chat_title,
            "call_mode": col.message.call_mode,
            "model_type": col.message.model_type,
            "cloud_model_id": col.message.cloud_model_id
        } for col in collections]

        return JsonResponse({'code': 200, 'msg': '加载成功', 'data': collected_msgs})
    except Exception as e:
        print(f"加载收藏消息失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '加载失败', 'data': []})


@login_required(login_url='login')
@require_http_methods(['POST'])
def batch_collect_messages(request):
    """批量收藏消息"""
    try:
        data = json.loads(request.body)
        msg_ids = data.get('msg_ids', [])
        if not isinstance(msg_ids, list) or len(msg_ids) == 0:
            return JsonResponse({'code': 400, 'msg': '请选择至少一条消息', 'data': None})

        messages = Message.objects.filter(msg_id__in=msg_ids, chat__user=request.user)
        valid_msg_ids = [msg.msg_id for msg in messages]
        if len(valid_msg_ids) == 0:
            return JsonResponse({'code': 400, 'msg': '无有效消息可收藏', 'data': None})

        collected_count = 0
        for msg_id in valid_msg_ids:
            _, created = MessageCollection.objects.get_or_create(
                user=request.user,
                message=messages.get(msg_id=msg_id)
            )
            if created:
                collected_count += 1

        return JsonResponse({
            'code': 200,
            'msg': '收藏成功！',
            'data': {'collected_count': collected_count, 'total_count': len(valid_msg_ids)}
        })
    except Message.DoesNotExist:
        return JsonResponse({'code': 404, 'msg': '部分消息不存在或无操作权限', 'data': None})
    except Exception as e:
        print(f"批量收藏失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '批量收藏失败', 'data': None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def batch_cancel_collect_messages(request):
    """批量取消收藏消息"""
    try:
        data = json.loads(request.body)
        msg_ids = data.get('msg_ids', [])
        if not isinstance(msg_ids, list) or len(msg_ids) == 0:
            return JsonResponse({'code': 400, 'msg': '请选择至少一条消息', 'data': None})

        messages = Message.objects.filter(msg_id__in=msg_ids, chat__user=request.user)
        valid_msg_ids = [msg.id for msg in messages]
        if len(valid_msg_ids) == 0:
            return JsonResponse({'code': 400, 'msg': '无有效消息可取消收藏', 'data': None})

        deleted_count, _ = MessageCollection.objects.filter(
            user=request.user,
            message_id__in=valid_msg_ids
        ).delete()

        return JsonResponse({
            'code': 200,
            'msg': f'批量取消收藏成功！共取消{deleted_count}条消息收藏',
            'data': {'deleted_count': deleted_count, 'total_count': len(valid_msg_ids)}
        })
    except Exception as e:
        print(f"批量取消收藏失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': '批量取消收藏失败', 'data': None})


@login_required(login_url='login')
def collection_view(request):
    """收藏夹页面视图"""
    return render(request, 'myapp/message_collection.html')


# 加载轻量相似度模型（首次运行自动下载，约100MB）
similarity_model = SentenceTransformer('./all-MiniLM-L6-v2')


def split_text_by_chunk(text, chunk_size=300, chunk_overlap=50):
    """按指定大小和重叠长度拆分文本为片段"""
    print("="*60)
    print(f"[RAG流程-1/5] 开始文本分段处理")
    print(f"[文本分段] 原始文本长度: {len(text)} 字符")
    print(f"[文本分段] 分段配置 - chunk_size: {chunk_size}, chunk_overlap: {chunk_overlap}")

    chunks = []
    start = 0
    text_length = len(text)
    chunk_index = 1

    while start < text_length:
        end = min(start + chunk_size, text_length)
        chunk = text[start:end]
        if not chunk:
            break
        chunks.append(chunk)
        print(f"[文本分段] 片段{chunk_index:02d}: 位置[{start}-{end}], 长度={len(chunk)}字符")
        print(f"[文本分段] 内容预览: {chunk[:30]}..." if len(chunk) > 30 else f"[文本分段] 内容: {chunk}")
        next_start = end - chunk_overlap
        if next_start <= start:
            break
        start = next_start
        chunk_index += 1

    print(f"[文本分段] 分段完成！共生成 {len(chunks)} 个片段")
    print("="*60)
    return chunks


def extract_keywords(text, top_k=5):
    """提取文本关键词（基于词频，过滤停用词）"""
    stop_words = {'的', '了', '是', '在', '有', '和', '就', '不', '人', '都', '一', '要', '我', '他', '她', '它', '这', '那', '此', '其', '从', '到', '与', '及', '而', '于', '对', '向', '以', '为', '把', '被', '给', '跟', '同', '比', '因', '由', '用', '使', '让', '叫', '被', '所', '的', '地', '得', '着', '过', '来', '去', '上', '下', '进', '出', '开', '关', '走', '跑', '跳', '飞', '坐', '站', '躺', '睡', '吃', '喝', '穿', '戴', '拿', '放', '做', '干', '学', '看', '听', '说', '读', '写'}
    
    print("-"*60)
    print(f"[RAG流程-2/5] 开始关键词提取")
    print(f"[关键词提取] 输入文本长度: {len(text)} 字符")
    
    words = jieba.lcut(text)
    print(f"[关键词提取] 分词结果（部分展示）: {words[:10]}...")
    
    filtered_words = [word for word in words if word not in stop_words and len(word) > 1]
    print(f"[关键词提取] 过滤停用词后: {len(filtered_words)} 个词")
    
    word_freq = {}
    for word in filtered_words:
        word_freq[word] = word_freq.get(word, 0) + 1
    
    sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:top_k]
    keywords = [word for word, freq in sorted_words]
    
    print(f"[关键词提取] Top-{top_k} 关键词:")
    for i, (word, freq) in enumerate(sorted_words, 1):
        print(f"            {i}. {word} (词频: {freq})")
    print("-"*60)
    
    return keywords


def calculate_rag_score(query, chunk, rag_config):
    """计算文本片段与查询的相似度得分（综合语义相似度、关键词匹配、长度比例）"""
    print("-"*60)
    print(f"[RAG流程-3/5] 开始相似度评分计算")
    print(f"[评分计算] 用户问题: {query[:50]}..." if len(query) > 50 else f"[评分计算] 用户问题: {query}")
    print(f"[评分计算] 知识库片段: {chunk[:50]}..." if len(chunk) > 50 else f"[评分计算] 知识库片段: {chunk}")
    
    # 1. 语义相似度计算
    print("\n[评分计算] 步骤1: 语义相似度计算")
    query_embedding = similarity_model.encode(query, convert_to_tensor=True)
    chunk_embedding = similarity_model.encode(chunk, convert_to_tensor=True)
    similarity_score = max(0, min(1, util.cos_sim(query_embedding, chunk_embedding).item()))
    print(f"            语义相似度得分: {similarity_score:.4f}")
    
    # 2. 关键词匹配计算
    print("\n[评分计算] 步骤2: 关键词匹配计算")
    query_keywords = extract_keywords(query)
    chunk_keywords = extract_keywords(chunk)
    print(f"            用户问题关键词: {query_keywords}")
    print(f"            知识库片段关键词: {chunk_keywords}")
    
    matched_keywords = set(query_keywords) & set(chunk_keywords)
    keyword_score = len(matched_keywords) / len(query_keywords) if query_keywords else 0
    print(f"            匹配关键词: {list(matched_keywords)}")
    print(f"            关键词匹配得分: {keyword_score:.4f}")
    
    # 3. 长度比例计算
    print("\n[评分计算] 步骤3: 长度比例计算")
    query_length, chunk_length = len(query), len(chunk)
    length_score = min(query_length, chunk_length) / max(query_length, chunk_length)
    print(f"            用户问题长度: {query_length} 字符")
    print(f"            知识库片段长度: {chunk_length} 字符")
    print(f"            长度比例得分: {length_score:.4f}")
    
    # 4. 综合评分
    print("\n[评分计算] 步骤4: 综合评分")
    sim_weight = rag_config.get('similarity_weight', 0.6)
    kw_weight = rag_config.get('keyword_weight', 0.3)
    len_weight = rag_config.get('length_weight', 0.1)
    
    final_score = (
        similarity_score * sim_weight +
        keyword_score * kw_weight +
        length_score * len_weight
    )
    
    print(f"            权重配置: 语义相似度({sim_weight}) + 关键词匹配({kw_weight}) + 长度比例({len_weight})")
    print(f"            综合得分 = ({similarity_score:.4f} * {sim_weight}) + ({keyword_score:.4f} * {kw_weight}) + ({length_score:.4f} * {len_weight})")
    print(f"            = {similarity_score * sim_weight:.4f} + {keyword_score * kw_weight:.4f} + {length_score * len_weight:.4f}")
    print(f"            = {final_score:.4f}")
    
    print("\n[评分计算] 计算完成！")
    print("-"*60)
    
    return final_score


# 大模型文本生成接口
logger = logging.getLogger(__name__)


@login_required(login_url='login')
@csrf_exempt
@require_http_methods(['POST'])
def generate_text_api(request):
    """RAG增强文本生成接口：支持知识库检索和上下文记忆
    """
    try:
        data = json.loads(request.body)
        user_query = data.get("user_query") or data.get("message")
        kb_id = data.get("kb_id", "default")
        model_config = data.get("model_config", {})
        model_name = data.get("model_name", "qwen3-4b-lora-fp8")  # 新增：模型名称参数
        rag_config = data.get("rag_config", {
            "chunk_size": 300,
            "chunk_overlap": 50,
            "similarity_threshold": 0.3,
            "similarity_weight": 0.6,
            "keyword_weight": 0.3,
            "length_weight": 0.1,
            "top_n": 3
        })

        # ==================== 阶段0：参数接收 ====================
        print("\n" + "="*80)
        print("                    【RAG智能问答系统 - 完整流程演示】                    ")
        print("="*80)
        print(f"\n[阶段0/5] 参数接收")
        print(f"┌─────────────────────────────────────────────────────────────┐")
        print(f"│ 用户提问: {user_query[:60]}..." if len(user_query) > 60 else f"│ 用户提问: {user_query}")
        print(f"│ 知识库ID: {kb_id}")
        print(f"│ 是否使用知识库: {'是' if kb_id != 'default' else '否（通用问答）'}")
        print(f"└─────────────────────────────────────────────────────────────┘")
        
        print("\n[RAG配置参数]")
        print(f"  ├─ 文本分段: chunk_size={rag_config['chunk_size']}, overlap={rag_config['chunk_overlap']}")
        print(f"  ├─ 相似度阈值: {rag_config['similarity_threshold']}")
        print(f"  ├─ 权重分配: 语义相似度({rag_config['similarity_weight']}) + 关键词({rag_config['keyword_weight']}) + 长度({rag_config['length_weight']})")
        print(f"  └─ Top-N检索: {rag_config['top_n']}")
        
        print("\n[模型配置参数]")
        print(f"  ├─ temperature: {model_config.get('temperature', 0.2)}")
        print(f"  ├─ max_tokens: {model_config.get('max_tokens', 8192)}")
        print(f"  ├─ top_p: {model_config.get('top_p', 0.9)}")
        print(f"  └─ frequency_penalty: {model_config.get('frequency_penalty', 0.1)}")

        if not user_query:
            return JsonResponse({"code": 400, "msg": "请输入问题！", "data": {}})

        reference_knowledge = user_query
        
        # ==================== 阶段1-4：RAG检索流程 ====================
        if kb_id != "default":
            try:
                kb = KnowledgeBase.objects.get(kb_id=kb_id, user=request.user)
                print(f"\n{'='*80}")
                print(f"[阶段1/5] 加载知识库")
                print(f"知识库名称: {kb.name}")
                print(f"知识库内容长度: {len(kb.content)} 字符")
                
                # 阶段2: 文本分段
                chunks = split_text_by_chunk(
                    kb.content,
                    chunk_size=rag_config["chunk_size"],
                    chunk_overlap=rag_config["chunk_overlap"]
                )
                print(f"\n{'='*80}")
                print(f"[阶段2/5] 文本分段完成")
                print(f"生成片段总数: {len(chunks)} 个")

                if not chunks:
                    reference_knowledge = f"知识库无有效内容，用户问题：{user_query}"
                else:
                    # 阶段3: 相似度计算（已在calculate_rag_score中输出）
                    chunk_scores = []
                    print(f"\n{'='*80}")
                    print(f"[阶段3/5] 相似度匹配")
                    print(f"正在对 {len(chunks)} 个片段进行评分计算...")
                    
                    for idx, chunk in enumerate(chunks, 1):
                        print(f"\n处理片段 {idx}/{len(chunks)}...")
                        score = calculate_rag_score(user_query, chunk, rag_config)
                        if score >= rag_config["similarity_threshold"]:
                            chunk_scores.append((chunk, score))
                            print(f" ✓ 片段{idx} 得分: {score:.4f} ≥ 阈值{rag_config['similarity_threshold']}, 保留")
                        else:
                            print(f" ✗ 片段{idx} 得分: {score:.4f} < 阈值{rag_config['similarity_threshold']}, 过滤")

                    # 阶段4: Top-N选择
                    print(f"\n{'='*80}")
                    print(f"[阶段4/5] Top-N 选择")
                    print(f"符合阈值的片段数: {len(chunk_scores)}")
                    
                    chunk_scores.sort(key=lambda x: x[1], reverse=True)
                    top_chunks = [chunk for chunk, score in chunk_scores[:rag_config["top_n"]]]
                    top_scores = [score for chunk, score in chunk_scores[:rag_config["top_n"]]]

                    print(f"最终选中 Top-{len(top_chunks)} 片段:")
                    for i, (chunk, score) in enumerate(zip(top_chunks, top_scores), 1):
                        print(f"\n  [{i}] 得分: {score:.4f}")
                        print(f"    内容预览: {chunk[:80]}...")

                    if top_chunks:
                        reference_knowledge = f"""以下是参考知识库的相关片段（按相关性排序）：
{chr(10).join([f"片段{idx + 1}：{chunk}" for idx, chunk in enumerate(top_chunks)])}

请基于上述片段回答用户问题，优先使用知识库信息，若片段无相关内容可结合自身知识回答，回答需简洁准确：
用户问题：{user_query}"""
                    else:
                        reference_knowledge = f"知识库中无相关内容，用户问题：{user_query}"
            except KnowledgeBase.DoesNotExist:
                print(f"\n[警告] 未找到知识库ID: {kb_id}，将使用通用问答模式")

        # ==================== 阶段5：大模型调用 ====================
        print(f"\n{'='*80}")
        print(f"[阶段5/5] 大模型推理")
        print(f"┌─────────────────────────────────────────────────────────────┐")
        print(f"│ 输入提示词长度: {len(reference_knowledge)} 字符")
        print(f"│ 模型地址: https://u964814-962f-1298fd9f.bjb1.seetacloud.com:8443/v1")
        print(f"│ 模型名称: qwen3-4b")
        print(f"└─────────────────────────────────────────────────────────────┘")
        print(f"\n正在调用大模型生成回答...")
        
        import time
        start = time.perf_counter()
        result = local_chat(reference_knowledge, model_config, model_name)
        mdtime = int((time.perf_counter() - start) * 1000)
        
        # 输出大模型调用结果
        print(f"\n[大模型推理结果]")
        print(f"┌─────────────────────────────────────────────────────────────┐")
        print(f"│ 推理耗时: {mdtime} ms")
        print(f"│ 回答长度: {len(result)} 字符")
        print(f"└─────────────────────────────────────────────────────────────┘")
        print(f"\n[回答内容]:")
        print(f"{result[:500]}..." if len(result) > 500 else result)
        print(f"\n{'='*80}")
        print("                      【流程结束】                      ")
        print("="*80)
        
        return JsonResponse({"code": 200, "msg": "success", "data": {"result": result, "mdtime": mdtime}})
    except Exception as e:
        print(f"\n{'='*80}")
        print(f"                    【错误】大模型推理失败                    ")
        print(f"{'='*80}")
        print(f"错误信息: {str(e)}")
        print(f"{'='*80}")
        logger.error(f"[RAG检索错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": f"服务器错误，生成失败: {str(e)}", "data": {}})


# ==================== 多模型调用API ====================

@login_required(login_url='login')
@csrf_exempt
@require_http_methods(['POST'])
def multi_model_generate_api(request):
    """多模型调用接口：支持本地、云端、对比模式"""
    try:
        data = json.loads(request.body)
        user_query = data.get("user_query") or data.get("message")
        kb_id = data.get("kb_id", "default")
        model_config = data.get("model_config", {})
        call_mode = data.get("call_mode", "local")  # local, cloud, compare
        cloud_model_id = data.get("cloud_model_id")
        
        if not user_query:
            return JsonResponse({"code": 400, "msg": "请输入问题！", "data": {}})

        # 获取参考知识（RAG检索）
        reference_knowledge = user_query
        if kb_id != "default":
            try:
                kb = KnowledgeBase.objects.get(kb_id=kb_id, user=request.user)
                chunks = split_text_by_chunk(kb.content, chunk_size=300, chunk_overlap=50)
                if chunks:
                    chunk_scores = []
                    rag_config = {"similarity_threshold": 0.3, "similarity_weight": 0.6, "keyword_weight": 0.3, "length_weight": 0.1}
                    for chunk in chunks:
                        score = calculate_rag_score(user_query, chunk, rag_config)
                        if score >= 0.3:
                            chunk_scores.append((chunk, score))
                    chunk_scores.sort(key=lambda x: x[1], reverse=True)
                    top_chunks = [chunk for chunk, score in chunk_scores[:3]]
                    if top_chunks:
                        reference_knowledge = f"""以下是参考知识库的相关片段：
{chr(10).join([f"片段{idx + 1}：{chunk}" for idx, chunk in enumerate(top_chunks)])}

请基于上述片段回答用户问题：{user_query}"""
            except KnowledgeBase.DoesNotExist:
                pass

        result_data = {}

        # 本地模型调用
        def call_local_model(model_name="qwen3-4b-lora-fp8"):
            import time
            start = time.perf_counter()
            result = local_chat(reference_knowledge, model_config, model_name)
            total_time_ms = int((time.perf_counter() - start) * 1000)
            ttft_ms = total_time_ms // 2  # 估算TTFT
            token_count = len(result)
            return {
                "result": result,
                "total_time_ms": total_time_ms,
                "ttft_ms": ttft_ms,
                "token_count": token_count
            }

        # 云端模型调用
        def call_cloud_model(cloud_model_id):
            print(f"调用云端模型: {cloud_model_id}")
            try:
                # 优先使用统一的多模型配置解析器
                matched_model = None
                
                # 先尝试按模型名称匹配（前端传递的是模型名称）
                for model_type, models in CLOUD_MODELS.items():
                    for model in models:
                        if model['name'] == cloud_model_id:
                            matched_model = model
                            matched_model['type'] = model_type
                            break
                    if matched_model:
                        break
                print(f"匹配到的模型: {matched_model.values()}")
                # 如果没找到，尝试按环境变量前缀匹配（旧格式）
                if not matched_model:
                    print(f"未匹配到模型: {cloud_model_id}")
                    old_configs = {
                        "doubao": {
                            "name": "豆包",
                            "key_env": "DOUBAO_API_KEY",
                            "url_env": "DOUBAO_API_URL",
                            "name_env": "DOUBAO_MODEL_NAME",
                            "default_url": "https://api.doubao.com/v1/chat/completions",
                            "default_name": "Doubao-3.5-Turbo",
                            "call_func": call_doubao
                        },
                        "deepseek": {
                            "name": "DeepSeek",
                            "key_env": "DEEPSEEK_API_KEY",
                            "url_env": "DEEPSEEK_API_URL",
                            "name_env": "DEEPSEEK_MODEL_NAME",
                            "default_url": "https://api.deepseek.com/v1/chat/completions",
                            "default_name": "deepseek-chat",
                            "call_func": call_deepseek
                        },
                        "deepseek-v4": {
                            "name": "DeepSeek-V4-Flash",
                            "key_env": "DEEPSEEK_V4_API_KEY",
                            "url_env": "DEEPSEEK_V4_API_URL",
                            "name_env": "DEEPSEEK_V4_MODEL_NAME",
                            "default_url": "https://api.deepseek.com/v1/chat/completions",
                            "default_name": "deepseek-chat",
                            "call_func": call_deepseek_v4
                        }
                    }
                    
                    # 检查是否是旧格式的模型ID
                    config = old_configs.get(cloud_model_id)
                    print(f"旧格式模型配置: {config.values()}")
                    if config:
                        api_key = os.getenv(config["key_env"])
                        if not api_key:
                            return {"result": f"错误：{config['name']} API密钥未配置！", "total_time_ms": 0, "ttft_ms": 0, "token_count": 0, "cloud_model_id": cloud_model_id}
                        api_url = os.getenv(config["url_env"], config["default_url"])
                        model_name = os.getenv(config["name_env"], config["default_name"])
                        
                        result, total_time_ms, ttft_ms, token_count = config["call_func"](
                            reference_knowledge, 
                            model_config, 
                            api_url, 
                            model_name
                        )
                        
                        return {
                            "result": result,
                            "total_time_ms": total_time_ms,
                            "ttft_ms": ttft_ms,
                            "token_count": token_count,
                            "model_name": config["name"],
                            "cloud_model_id": cloud_model_id
                        }
                    
                    return {"result": "错误：不支持的模型类型", "total_time_ms": 0, "ttft_ms": 0, "token_count": 0, "cloud_model_id": cloud_model_id}
                
                # 使用新格式的模型配置
                api_key = matched_model.get('api_key')
                print(f"=== 调试信息 ===")
                print(f"matched_model 完整内容: {matched_model}")
                print(f"api_key 值: '{api_key}'")
                print(f"api_key 类型: {type(api_key)}")
                print(f"api_key 是否为空: {not api_key}")
                print(f"api_key 长度: {len(api_key) if api_key else 0}")
                if not api_key or api_key.strip() == '':
                    return {"result": f"错误：{matched_model['name']} API密钥未配置！", "total_time_ms": 0, "ttft_ms": 0, "token_count": 0, "cloud_model_id": cloud_model_id}
                
                api_url = matched_model['api_url']
                if not api_url:
                    # 设置默认URL
                    if matched_model['type'] == 'doubao':
                        api_url = 'https://api.doubao.com/v1/chat/completions'
                    else:
                        api_url = 'https://api.deepseek.com/v1/chat/completions'
                
                model_name = matched_model['name']
                
                # 根据模型类型选择调用函数
                if matched_model['type'] == 'doubao':
                    call_func = call_doubao
                elif matched_model['type'] == 'deepseek':
                    call_func = call_deepseek_v4
                else:
                    return {"result": "错误：不支持的模型类型", "total_time_ms": 0, "ttft_ms": 0, "token_count": 0, "cloud_model_id": cloud_model_id}
                
                # 调用模型服务（传递api_key参数）
                result, total_time_ms, ttft_ms, token_count = call_func(
                    reference_knowledge, 
                    model_config, 
                    api_url, 
                    model_name,
                    None,  # context_messages
                    api_key  # api_key
                )
                
                return {
                    "result": result,
                    "total_time_ms": total_time_ms,
                    "ttft_ms": ttft_ms,
                    "token_count": token_count,
                    "model_name": model_name,
                    "cloud_model_id": cloud_model_id
                }
            except Exception as e:
                return {"result": f"错误：{str(e)}", "total_time_ms": 0, "ttft_ms": 0, "token_count": 0, "cloud_model_id": cloud_model_id}

        # 根据模式调用模型
        if call_mode == "local":
            local_result = call_local_model()
            result_data["local"] = local_result
            # 记录日志
            ModelCallLog.objects.create(
                user=request.user,
                model_name="Qwen3-4B-FP8",
                call_mode="local",
                input_text=user_query[:500],
                output_text=local_result["result"][:500] if local_result["result"] else "",
                total_time_ms=local_result["total_time_ms"],
                ttft_ms=local_result["ttft_ms"],
                token_count=local_result["token_count"],
                success=not local_result["result"].startswith("错误")
            )

        elif call_mode == "cloud":
            if not cloud_model_id:
                return JsonResponse({"code": 400, "msg": "请选择云端模型！", "data": {}})
            cloud_result = call_cloud_model(cloud_model_id)
            result_data["cloud"] = cloud_result
            # 记录日志
            ModelCallLog.objects.create(
                user=request.user,
                model_name=cloud_result.get("model_name", "Unknown"),
                call_mode="cloud",
                input_text=user_query[:500],
                output_text=cloud_result["result"][:500] if cloud_result["result"] else "",
                total_time_ms=cloud_result["total_time_ms"],
                ttft_ms=cloud_result["ttft_ms"],
                token_count=cloud_result["token_count"],
                success=not cloud_result["result"].startswith("错误")
            )

        elif call_mode == "compare":
            if not cloud_model_id:
                return JsonResponse({"code": 400, "msg": "请选择云端模型！", "data": {}})
            # 并行调用两个模型
            local_result = call_local_model()
            cloud_result = call_cloud_model(cloud_model_id)
            result_data["local"] = local_result
            result_data["cloud"] = cloud_result
            # 记录日志
            ModelCallLog.objects.create(
                user=request.user,
                model_name="Qwen3-4B-FP8",
                call_mode="compare",
                input_text=user_query[:500],
                output_text=local_result["result"][:500] if local_result["result"] else "",
                total_time_ms=local_result["total_time_ms"],
                ttft_ms=local_result["ttft_ms"],
                token_count=local_result["token_count"],
                success=not local_result["result"].startswith("错误")
            )
            ModelCallLog.objects.create(
                user=request.user,
                model_name=cloud_result.get("model_name", "Unknown"),
                call_mode="compare",
                input_text=user_query[:500],
                output_text=cloud_result["result"][:500] if cloud_result["result"] else "",
                total_time_ms=cloud_result["total_time_ms"],
                ttft_ms=cloud_result["ttft_ms"],
                token_count=cloud_result["token_count"],
                success=not cloud_result["result"].startswith("错误")
            )

        return JsonResponse({"code": 200, "msg": "success", "data": result_data})

    except Exception as e:
        logger.error(f"[多模型调用错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": f"服务器错误: {str(e)}", "data": {}})


@login_required(login_url='login')
@require_http_methods(['GET'])
def get_cloud_models(request):
    """获取可用云端模型列表（混合架构：环境变量 + 数据库）"""
    try:
        model_list = []
        existing_ids = set()
        
        # 1. 从环境变量读取配置（优先）
        cloud_models = CLOUD_MODELS
        
        # 处理豆包模型
        for model in cloud_models.get('doubao', []):
            model_id = model['env_prefix']
            model_list.append({
                "model_id": model_id,
                "name": model['name'],
                "model_type": 'doubao',
                "model_type_display": model['name'],
                "api_url": model['api_url'],
                "model_name": model['name']
            })
            existing_ids.add(model_id)
        
        # 处理DeepSeek模型
        for model in cloud_models.get('deepseek', []):
            model_id = model['env_prefix']
            model_list.append({
                "model_id": model_id,
                "name": model['name'],
                "model_type": 'deepseek-v4',
                "model_type_display": model['name'],
                "api_url": model['api_url'],
                "model_name": model['name']
            })
            existing_ids.add(model_id)
        
        # 2. 从旧格式环境变量读取（后备）
        if not existing_ids:
            import os
            old_configs = [
                {
                    "model_id": "doubao",
                    "name": "豆包",
                    "type": "doubao",
                    "key_env": "DOUBAO_API_KEY",
                    "url_env": "DOUBAO_API_URL",
                    "name_env": "DOUBAO_MODEL_NAME",
                    "default_url": "https://api.doubao.com/v1/chat/completions",
                    "default_name": "Doubao-3.5-Turbo"
                },
                {
                    "model_id": "deepseek",
                    "name": "DeepSeek",
                    "type": "deepseek",
                    "key_env": "DEEPSEEK_API_KEY",
                    "url_env": "DEEPSEEK_API_URL",
                    "name_env": "DEEPSEEK_MODEL_NAME",
                    "default_url": "https://api.deepseek.com/v1/chat/completions",
                    "default_name": "deepseek-chat"
                },
                {
                    "model_id": "deepseek-v4",
                    "name": "DeepSeek-V4-Flash",
                    "type": "deepseek-v4",
                    "key_env": "DEEPSEEK_V4_API_KEY",
                    "url_env": "DEEPSEEK_V4_API_URL",
                    "name_env": "DEEPSEEK_V4_MODEL_NAME",
                    "default_url": "https://api.deepseek.com/v1/chat/completions",
                    "default_name": "deepseek-chat"
                }
            ]
            
            for config in old_configs:
                api_key = os.getenv(config["key_env"])
                if api_key and config["model_id"] not in existing_ids:
                    api_url = os.getenv(config["url_env"], config["default_url"])
                    model_name = os.getenv(config["name_env"], config["default_name"])
                    model_list.append({
                        "model_id": config["model_id"],
                        "name": model_name,
                        "model_type": config["type"],
                        "model_type_display": model_name,
                        "api_url": api_url,
                        "model_name": model_name
                    })
                    existing_ids.add(config["model_id"])
        
        # 3. 从数据库读取配置（补充，环境变量未配置的模型）
        db_models = CloudModel.objects.filter(is_active=True)
        for db_model in db_models:
            if db_model.model_id not in existing_ids:
                # 尝试从环境变量获取API密钥
                import os
                api_key_env = f"{db_model.model_type.upper()}_{db_model.name.upper().replace('-', '_')}_API_KEY"
                api_key = os.getenv(api_key_env)
                
                if api_key or db_model.api_key:  # 有密钥才加入
                    model_list.append({
                        "model_id": db_model.model_id,
                        "name": db_model.name,
                        "model_type": db_model.model_type,
                        "model_type_display": db_model.name,
                        "api_url": db_model.api_url or '',
                        "model_name": db_model.model_name or db_model.name
                    })
        
        return JsonResponse({"code": 200, "msg": "success", "data": model_list})
    except Exception as e:
        return JsonResponse({"code": 500, "msg": str(e), "data": []})


# ==================== 统计分析API ====================

@login_required(login_url='login')
@require_http_methods(['POST'])
def get_statistics_time(request):
    """获取TTFT统计数据：平均首Token时间"""
    try:
        data = json.loads(request.body)
        start_date = data.get("start_date")
        end_date = data.get("end_date")
        
        # 从Message表获取真实会话数据，只统计assistant消息（模型回复）中包含ttft_ms的消息
        messages = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            model_type__isnull=False,
            ttft_ms__isnull=False
        )
        
        if start_date:
            messages = messages.filter(send_time__gte=start_date)
        if end_date:
            messages = messages.filter(send_time__lte=end_date)

        daily_ttft = {}
        model_ttft_stats = {}
        
        for msg in messages:
            date_str = msg.send_time.strftime("%Y-%m-%d")
            model_name = msg.chat.model_name or msg.model_type  # 优先使用chat中的model_name，否则使用model_type
            
            if date_str not in daily_ttft:
                daily_ttft[date_str] = {}
            if model_name not in daily_ttft[date_str]:
                daily_ttft[date_str][model_name] = {"total": 0, "count": 0}
            
            daily_ttft[date_str][model_name]["total"] += msg.ttft_ms
            daily_ttft[date_str][model_name]["count"] += 1
            
            if model_name not in model_ttft_stats:
                model_ttft_stats[model_name] = {"total": 0, "count": 0}
            model_ttft_stats[model_name]["total"] += msg.ttft_ms
            model_ttft_stats[model_name]["count"] += 1

        daily_avg_ttft = []
        for date in sorted(daily_ttft.keys()):
            day_data = {"date": date}
            for model, stats in daily_ttft[date].items():
                day_data[model] = round(stats["total"] / stats["count"], 2)
            daily_avg_ttft.append(day_data)

        model_avg_ttft = []
        for model, stats in model_ttft_stats.items():
            model_avg_ttft.append({
                "model_name": model,
                "avg_ttft_ms": round(stats["total"] / stats["count"], 2) if stats["count"] > 0 else 0,
                "call_count": stats["count"]
            })

        return JsonResponse({"code": 200, "msg": "success", "data": {
            "daily_avg_ttft": daily_avg_ttft,
            "model_avg_ttft": model_avg_ttft
        }})
    except Exception as e:
        logger.error(f"[TTFT统计错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": str(e), "data": {}})


@login_required(login_url='login')
@require_http_methods(['POST'])
def get_statistics_tokens(request):
    """获取Token统计数据：按模型和日期汇总"""
    try:
        data = json.loads(request.body)
        time_range = data.get("time_range", "week")  # day, week, month
        
        now = timezone.now()
        if time_range == "day":
            start_time = now - timedelta(days=1)
        elif time_range == "week":
            start_time = now - timedelta(days=7)
        else:
            start_time = now - timedelta(days=30)
        
        # 从Message表获取真实会话数据，只统计assistant消息（模型回复）中包含token_count的消息
        messages = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            send_time__gte=start_time,
            model_type__isnull=False,
            token_count__isnull=False
        )

        daily_tokens = {}
        model_daily_tokens = {}
        model_total_tokens = {}
        
        for msg in messages:
            date_str = msg.send_time.strftime("%Y-%m-%d")
            model_name = msg.chat.model_name or msg.model_type  # 优先使用chat中的model_name，否则使用model_type
            token_count = msg.token_count or 0
            
            if date_str not in daily_tokens:
                daily_tokens[date_str] = 0
            daily_tokens[date_str] += token_count
            
            if model_name not in model_daily_tokens:
                model_daily_tokens[model_name] = {}
            if date_str not in model_daily_tokens[model_name]:
                model_daily_tokens[model_name][date_str] = 0
            model_daily_tokens[model_name][date_str] += token_count
            
            if model_name not in model_total_tokens:
                model_total_tokens[model_name] = 0
            model_total_tokens[model_name] += token_count

        # 构建每日Token数据（按模型分组）
        daily_token_data = []
        all_dates = sorted(set(daily_tokens.keys()))
        for date_str in all_dates:
            day_data = {"date": date_str}
            for model_name in model_total_tokens.keys():
                day_data[model_name] = model_daily_tokens.get(model_name, {}).get(date_str, 0)
            daily_token_data.append(day_data)

        total_all = sum(model_total_tokens.values())
        model_ratio = []
        for model_name, total in model_total_tokens.items():
            model_ratio.append({
                "model_name": model_name,
                "total_tokens": total,
                "percentage": round(total / total_all * 100, 2) if total_all > 0 else 0
            })

        return JsonResponse({"code": 200, "msg": "success", "data": {
            "daily_token_data": daily_token_data,
            "model_ratio": model_ratio,
            "total_all": total_all,
            "time_range": time_range
        }})
    except Exception as e:
        logger.error(f"[Token统计错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": str(e), "data": {}})


@login_required(login_url='login')
@require_http_methods(['POST'])
def get_statistics_preference(request):
    """获取模型使用偏好统计"""
    try:
        data = json.loads(request.body)
        time_range = data.get("time_range", "week")  # day, week, month
        
        now = timezone.now()
        if time_range == "day":
            start_time = now - timedelta(days=1)
        elif time_range == "week":
            start_time = now - timedelta(days=7)
        else:
            start_time = now - timedelta(days=30)
        
        # 从Message表获取真实会话数据，只统计assistant消息（模型回复）
        messages = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            send_time__gte=start_time,
            model_type__isnull=False
        )
        
        # 按日期和模型名称分组统计每日调用次数
        daily_calls = {}
        model_daily_calls = {}
        model_total_calls = {}
        
        for msg in messages:
            date_str = msg.send_time.strftime("%Y-%m-%d")
            model_name = msg.chat.model_name or msg.model_type  # 优先使用chat中的model_name，否则使用model_type
            
            if date_str not in daily_calls:
                daily_calls[date_str] = 0
            daily_calls[date_str] += 1
            
            if model_name not in model_daily_calls:
                model_daily_calls[model_name] = {}
            if date_str not in model_daily_calls[model_name]:
                model_daily_calls[model_name][date_str] = 0
            model_daily_calls[model_name][date_str] += 1
            
            if model_name not in model_total_calls:
                model_total_calls[model_name] = 0
            model_total_calls[model_name] += 1
        
        # 构建每日调用数据（按模型分组）
        daily_call_data = []
        all_dates = sorted(set(daily_calls.keys()))
        for date_str in all_dates:
            day_data = {"date": date_str}
            for model_name in model_total_calls.keys():
                day_data[model_name] = model_daily_calls.get(model_name, {}).get(date_str, 0)
            daily_call_data.append(day_data)
        
        # 模型调用占比
        total_calls = sum(model_total_calls.values())
        model_ratio = []
        for model_name, count in model_total_calls.items():
            model_ratio.append({
                "model_name": model_name,
                "count": count,
                "percentage": round(count / total_calls * 100, 2) if total_calls > 0 else 0
            })

        return JsonResponse({"code": 200, "msg": "success", "data": {
            "daily_call_data": daily_call_data,
            "model_ratio": model_ratio,
            "time_range": time_range,
            "total_calls": total_calls
        }})
    except Exception as e:
        logger.error(f"[偏好统计错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": str(e), "data": {}})


@login_required(login_url='login')
@require_http_methods(['GET'])
def get_user_habit_suggestion(request):
    """获取用户习惯建议"""
    try:
        now = timezone.now()
        last_7_days = now - timedelta(days=7)
        last_30_days = now - timedelta(days=30)
        
        # 从Message表获取真实会话数据，只统计assistant消息（模型回复）
        messages_7d = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            send_time__gte=last_7_days,
            model_type__isnull=False
        )
        
        messages_30d = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            send_time__gte=last_30_days,
            model_type__isnull=False
        )
        
        # 统计7天各模型调用次数和Token
        model_stats_7d = {}
        model_stats_30d = {}
        total_calls_7d = 0
        total_tokens_7d = 0
        total_ttft_7d = {}
        
        for msg in messages_7d:
            model_name = msg.chat.model_name or msg.model_type  # 优先使用chat中的model_name，否则使用model_type
            if model_name not in model_stats_7d:
                model_stats_7d[model_name] = {"calls": 0, "tokens": 0, "ttft_list": []}
            model_stats_7d[model_name]["calls"] += 1
            model_stats_7d[model_name]["tokens"] += msg.token_count or 0
            if msg.ttft_ms:
                model_stats_7d[model_name]["ttft_list"].append(msg.ttft_ms)
            total_calls_7d += 1
            total_tokens_7d += msg.token_count or 0
            
            if model_name not in total_ttft_7d:
                total_ttft_7d[model_name] = []
            if msg.ttft_ms:
                total_ttft_7d[model_name].append(msg.ttft_ms)
        
        for msg in messages_30d:
            model_name = msg.chat.model_name or msg.model_type
            if model_name not in model_stats_30d:
                model_stats_30d[model_name] = {"calls": 0, "tokens": 0}
            model_stats_30d[model_name]["calls"] += 1
            model_stats_30d[model_name]["tokens"] += msg.token_count or 0
        
        # 计算各模型平均TTFT
        model_avg_ttft = {}
        for model_name, ttft_list in total_ttft_7d.items():
            if ttft_list:
                model_avg_ttft[model_name] = round(sum(ttft_list) / len(ttft_list) / 1000, 2)
        
        # 找出TTFT最快和最慢的模型
        fastest_model = min(model_avg_ttft.items(), key=lambda x: x[1]) if model_avg_ttft else None
        slowest_model = max(model_avg_ttft.items(), key=lambda x: x[1]) if model_avg_ttft else None
        
        # 找出Token消耗最高和最低的模型
        most_token_model = max(model_stats_7d.items(), key=lambda x: x[1]["tokens"]) if model_stats_7d else None
        least_token_model = min(model_stats_7d.items(), key=lambda x: x[1]["tokens"]) if model_stats_7d else None
        
        suggestions = []
        
        if total_calls_7d == 0:
            suggestion = "您还没有使用记录，开始体验吧！"
        else:
            # 基于使用频率的建议
            if total_calls_7d >= 50:
                suggestions.append(f"近7天您共使用了{total_calls_7d}次问答服务")
            elif total_calls_7d >= 20:
                suggestions.append(f"近7天您使用了{total_calls_7d}次问答服务")
            else:
                suggestions.append(f"近7天您使用了{total_calls_7d}次问答服务，还可以多尝试")
            
            # 基于Token消耗的建议
            if most_token_model and most_token_model[1]["tokens"] > 0:
                suggestions.append(f"Token消耗最高的模型是{most_token_model[0]}（{most_token_model[1]['tokens']}个）")
                if most_token_model[0] != fastest_model[0] if fastest_model else False:
                    suggestions.append(f"{fastest_model[0]}模型响应最快（{fastest_model[1]}s），可优先使用")
            
            # 基于TTFT的建议
            if fastest_model and slowest_model and fastest_model[0] != slowest_model[0]:
                ttft_diff = slowest_model[1] - fastest_model[1]
                if ttft_diff > 1:
                    suggestions.append(f"建议：{fastest_model[0]}响应最快（{fastest_model[1]}s），适合快速问答")
                if ttft_diff > 2:
                    suggestions.append(f"注意：{slowest_model[0]}响应较慢（{slowest_model[1]}s），复杂任务可考虑其他模型")
            
            # 基于模型使用的建议
            model_usage = list(model_stats_7d.items())
            if len(model_usage) > 1:
                most_used = max(model_usage, key=lambda x: x[1]["calls"])
                suggestions.append(f"您最常使用{most_used[0]}（{most_used[1]['calls']}次）")
                
                # 对比不同模型效果
                if most_used[1]["calls"] >= 5:
                    avg_ttft = model_avg_ttft.get(most_used[0], 0)
                    suggestions.append(f"继续保持良好的使用习惯！")
            
            # 综合建议
            if fastest_model and model_avg_ttft.get(fastest_model[0], 0) < 2:
                suggestions.append("本地模型响应速度快，适合日常简单问答")
            
            if not suggestions:
                suggestions.append("继续探索不同的问答方式，找到最适合您的模式！")
            
            suggestion = "；".join(suggestions)

        return JsonResponse({"code": 200, "msg": "success", "data": {
            "suggestion": suggestion,
            "model_stats_7d": model_stats_7d,
            "model_avg_ttft": model_avg_ttft,
            "total_calls_7d": total_calls_7d,
            "total_tokens_7d": total_tokens_7d
        }})
    except Exception as e:
        logger.error(f"[习惯建议错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": str(e), "data": {}})


# ==================== 模型对比分析API ====================

@login_required(login_url='login')
@require_http_methods(['POST'])
def get_model_comparison(request):
    """获取模型对比分析数据 - 从Message表获取真实性能指标数据（与用户习惯建议页面使用相同的数据查询逻辑）"""
    try:
        data = json.loads(request.body)
        model_names = data.get('models', [])
        time_range = data.get('time_range', 'week')  # day, week, month
        
        if not model_names or len(model_names) < 2:
            return JsonResponse({"code": 400, "msg": "请选择至少2个模型进行对比", "data": {}})
        
        # 根据时间范围设置查询条件（与用户习惯建议页面保持一致）
        now = timezone.now()
        if time_range == "day":
            start_time = now - timedelta(days=1)
        elif time_range == "month":
            start_time = now - timedelta(days=30)
        else:  # week
            start_time = now - timedelta(days=7)
        
        # 判断模型类型
        def get_model_type(model_name):
            if 'qwen' in model_name.lower():
                return 'local'
            return 'cloud'
        
        # 从Message表获取真实会话数据（与用户习惯建议页面使用相同的数据来源）
        # 只统计assistant消息（模型回复），且查询最近7天的数据
        messages = Message.objects.filter(
            user=request.user,
            role=2,  # assistant消息
            send_time__gte=start_time,
            model_type__isnull=False
        ).select_related('chat')  # 预加载chat，避免N+1查询
        
        # 按模型分组统计性能指标（与用户习惯建议页面完全一致）
        # 使用 msg.chat.model_name or msg.model_type 来确定模型名称
        model_stats_temp = {}
        
        for msg in messages:
            # 与用户习惯建议页面保持一致：优先使用chat中的model_name，否则使用model_type
            msg_model_name = msg.chat.model_name or msg.model_type
            
            if msg_model_name not in model_stats_temp:
                model_stats_temp[msg_model_name] = {
                    "calls": 0, 
                    "tokens": 0, 
                    "ttft_list": [],
                    "total_time_list": [],
                    "success_count": 0
                }
            
            model_stats_temp[msg_model_name]["calls"] += 1
            model_stats_temp[msg_model_name]["tokens"] += msg.token_count or 0
            if msg.ttft_ms:
                model_stats_temp[msg_model_name]["ttft_list"].append(msg.ttft_ms)
            if msg.total_time_ms:
                model_stats_temp[msg_model_name]["total_time_list"].append(msg.total_time_ms)
            # 计算成功率（排除错误消息）
            if not msg.content or not msg.content.startswith('错误'):
                model_stats_temp[msg_model_name]["success_count"] += 1
        
        # 计算总Token数用于计算占比
        total_tokens_all = sum(stats["tokens"] for stats in model_stats_temp.values())
        
        # 构建最终结果（只包含用户选择的模型）
        model_stats = {}
        
        for model_name in model_names:
            # 查找匹配的模型数据（支持精确匹配和包含匹配）
            matched_key = None
            for key in model_stats_temp.keys():
                if key == model_name or model_name in key or key in model_name:
                    matched_key = key
                    break
            
            if matched_key is None:
                # 如果没有找到匹配数据，返回空统计
                model_stats[model_name] = {
                    'display_name': model_name,
                    'call_count': 0,
                    'avg_total_time_ms': 0,
                    'avg_ttft_ms': 0,
                    'total_tokens': 0,
                    'token_percentage': 0,
                    'success_rate': 0,
                    'model_type': get_model_type(model_name)
                }
                continue
            
            stats = model_stats_temp[matched_key]
            call_count = stats["calls"]
            ttft_list = stats["ttft_list"]
            total_time_list = stats["total_time_list"]
            
            avg_total_time_ms = round(sum(total_time_list) / len(total_time_list), 2) if total_time_list else 0
            avg_ttft_ms = round(sum(ttft_list) / len(ttft_list), 2) if ttft_list else 0
            success_rate = round(stats["success_count"] / call_count * 100, 2) if call_count > 0 else 0
            token_percentage = round(stats["tokens"] / total_tokens_all * 100, 2) if total_tokens_all > 0 else 0
            
            model_stats[model_name] = {
                'display_name': model_name,
                'call_count': call_count,
                'avg_total_time_ms': avg_total_time_ms,
                'avg_ttft_ms': avg_ttft_ms,
                'total_tokens': stats["tokens"],
                'token_percentage': token_percentage,
                'success_rate': success_rate,
                'model_type': get_model_type(model_name)
            }
        
        return JsonResponse({"code": 200, "msg": "success", "data": model_stats})
    
    except Exception as e:
        logger.error(f"[模型对比分析错误] {str(e)}", exc_info=True)
        return JsonResponse({"code": 500, "msg": str(e), "data": {}})


# ==================== 云端模型管理API ====================

@login_required(login_url='login')
@require_http_methods(['POST'])
def create_cloud_model(request):
    """创建云端模型配置（管理员）"""
    if not request.user.is_staff:
        return JsonResponse({"code": 403, "msg": "无权限", "data": None})
    
    try:
        data = json.loads(request.body)
        name = data.get("name")
        model_type = data.get("model_type")
        api_key = data.get("api_key")
        api_url = data.get("api_url")
        model_name = data.get("model_name")
        
        if not name:
            return JsonResponse({"code": 400, "msg": "模型名称不能为空", "data": None})
        
        cloud_model = CloudModel.objects.create(
            name=name,
            model_type=model_type or "other",
            api_key=api_key,
            api_url=api_url,
            model_name=model_name
        )
        
        return JsonResponse({"code": 200, "msg": "创建成功", "data": {
            "model_id": cloud_model.model_id,
            "name": cloud_model.name
        }})
    except Exception as e:
        return JsonResponse({"code": 500, "msg": str(e), "data": None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def update_cloud_model(request):
    """更新云端模型配置（管理员）"""
    if not request.user.is_staff:
        return JsonResponse({"code": 403, "msg": "无权限", "data": None})
    
    try:
        data = json.loads(request.body)
        model_id = data.get("model_id")
        name = data.get("name")
        model_type = data.get("model_type")
        api_key = data.get("api_key")
        api_url = data.get("api_url")
        model_name = data.get("model_name")
        is_active = data.get("is_active")
        
        cloud_model = CloudModel.objects.get(model_id=model_id)
        
        if name is not None:
            cloud_model.name = name
        if model_type is not None:
            cloud_model.model_type = model_type
        if api_key is not None:
            cloud_model.api_key = api_key
        if api_url is not None:
            cloud_model.api_url = api_url
        if model_name is not None:
            cloud_model.model_name = model_name
        if is_active is not None:
            cloud_model.is_active = is_active
        
        cloud_model.save()
        
        return JsonResponse({"code": 200, "msg": "更新成功", "data": None})
    except CloudModel.DoesNotExist:
        return JsonResponse({"code": 404, "msg": "模型不存在", "data": None})
    except Exception as e:
        return JsonResponse({"code": 500, "msg": str(e), "data": None})


@login_required(login_url='login')
@require_http_methods(['POST'])
def delete_cloud_model(request):
    """删除云端模型配置（管理员）"""
    if not request.user.is_staff:
        return JsonResponse({"code": 403, "msg": "无权限", "data": None})
    
    try:
        data = json.loads(request.body)
        model_id = data.get("model_id")
        
        cloud_model = CloudModel.objects.get(model_id=model_id)
        cloud_model.delete()
        
        return JsonResponse({"code": 200, "msg": "删除成功", "data": None})
    except CloudModel.DoesNotExist:
        return JsonResponse({"code": 404, "msg": "模型不存在", "data": None})
    except Exception as e:
        return JsonResponse({"code": 500, "msg": str(e), "data": None})


# 个人中心视图

@login_required(login_url='login')
@require_http_methods(['GET', 'POST'])
def edit_profile(request):
    """编辑个人资料：支持修改用户名和邮箱"""
    if request.method == 'POST':
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        new_username = request.POST.get('new_username', '').strip() or request.user.username
        new_email = request.POST.get('new_email', '').strip() or request.user.email
        original_email = request.POST.get('original_email', '').strip() or request.user.email
        email_code = request.POST.get('email_code', '').strip()

        if not all([new_username, new_email]):
            res = {'status': 'error', 'message': "用户名/邮箱不能为空！"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/edit_profile.html', {'user': request.user, **res})

        if User.objects.filter(username=new_username).exclude(id=request.user.id).exists():
            res = {'status': 'error', 'message': "用户名已被占用！"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/edit_profile.html', {'user': request.user, **res})

        if new_email != original_email:
            if not email_code:
                res = {'status': 'error', 'message': "修改邮箱必须填写验证码！"}
                return JsonResponse(res) if is_ajax else render(request, 'myapp/edit_profile.html', {'user': request.user, **res})
            try:
                verify_code = VerifyCode.objects.get(email=new_email, code=email_code, is_used=False, expire_time__gte=timezone.now())
                verify_code.is_used = True
                verify_code.save()
            except VerifyCode.DoesNotExist:
                res = {'status': 'error', 'message': "验证码错误/已过期！"}
                return JsonResponse(res) if is_ajax else render(request, 'myapp/edit_profile.html', {'user': request.user, **res})

        try:
            user = request.user
            user.username = new_username
            user.email = new_email
            user.save(update_fields=['username', 'email'])
            res = {'status': 'success', 'message': "个人资料修改成功！"}
            return JsonResponse(res) if is_ajax else redirect('index')
        except Exception as e:
            res = {'status': 'error', 'message': f"修改失败：{str(e)[:50]}"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/edit_profile.html', {'user': request.user, **res})

    return render(request, 'myapp/edit_profile.html', {'user': request.user})


@require_http_methods(['POST'])
def check_email_unique(request):
    """检查邮箱是否唯一（排除当前用户）"""
    if request.headers.get('X-Requested-With') != 'XMLHttpRequest':
        return JsonResponse({'valid': False, 'message': '请求方式错误'})

    new_email = request.POST.get('new_email', '').strip()
    user_id = request.POST.get('user_id', '')
    if not new_email:
        return JsonResponse({'valid': False, 'message': '邮箱不能为空'})

    exists = User.objects.filter(email=new_email).exclude(id=user_id).exists()
    return JsonResponse({'valid': not exists, 'message': '邮箱可用' if not exists else '该邮箱已被使用！'})


@login_required(login_url='login')
@require_http_methods(['GET', 'POST'])
def change_pwd(request):
    """修改密码：需验证旧密码和邮箱验证码"""
    if request.method == 'POST':
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        old_password = request.POST.get('old_password', '').strip()
        new_password = request.POST.get('new_password', '').strip()
        new_password2 = request.POST.get('new_password2', '').strip()
        email_code = request.POST.get('email_code', '').strip()

        if not email_code:
            res = {'status': 'error', 'message': "修改失败！验证码不能为空"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

        try:
            verify_code = VerifyCode.objects.get(
                email=request.user.email,
                code=email_code,
                is_used=False,
                expire_time__gte=timezone.now()
            )
        except VerifyCode.DoesNotExist:
            code_exist = VerifyCode.objects.filter(email=request.user.email, code=email_code).exists()
            msg = '修改失败！验证码已使用' if code_exist else '修改失败！验证码已过期'
            res = {'status': 'error', 'message': msg}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

        if not all([old_password, new_password, new_password2]):
            res = {'status': 'error', 'message': "修改失败！所有字段不能为空"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

        if new_password != new_password2:
            res = {'status': 'error', 'message': "修改失败！两次密码不一致"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

        if not authenticate(request, username=request.user.username, password=old_password):
            res = {'status': 'error', 'message': "修改失败！用户密码错误"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

        try:
            user = request.user
            user.set_password(new_password)
            user.save()
            logout(request)
            verify_code.is_used = True
            verify_code.save()
            res = {'status': 'success', 'message': "修改成功！请重新登录..."}
            return JsonResponse(res) if is_ajax else redirect('login')
        except Exception as e:
            res = {'status': 'error', 'message': f"修改失败：{str(e)}"}
            return JsonResponse(res) if is_ajax else render(request, 'myapp/change_pwd.html', res)

    return render(request, 'myapp/change_pwd.html')


@login_required(login_url='login')
@require_http_methods(['POST'])
def upload_avatar(request):
    """上传用户头像：支持JPG、JPEG、PNG、GIF格式，最大5MB"""
    try:
        if 'avatar_file' not in request.FILES:
            return JsonResponse({'code': 400, 'msg': '未选择头像文件'})

        avatar_file = request.FILES['avatar_file']
        allowed_extensions = ['jpg', 'jpeg', 'png', 'gif']
        file_ext = avatar_file.name.split('.')[-1].lower()
        if file_ext not in allowed_extensions:
            return JsonResponse({'code': 400, 'msg': '上传失败！仅支持 JPG、JPEG、PNG、GIF 格式图片'})

        max_size = 5 * 1024 * 1024
        if avatar_file.size > max_size:
            return JsonResponse({'code': 400, 'msg': '上传失败！图片大小不能超过 5 MB'})

        user_id = request.user.id
        avatar_root_dir = os.path.join(settings.STATICFILES_DIRS[0], 'images', 'user_avatars')
        user_avatar_dir = os.path.join(avatar_root_dir, str(user_id))

        if not os.path.exists(avatar_root_dir):
            os.makedirs(avatar_root_dir)
        if not os.path.exists(user_avatar_dir):
            os.makedirs(user_avatar_dir)

        filename = avatar_file.name
        file_path = os.path.join(user_avatar_dir, filename)

        with open(file_path, 'wb+') as destination:
            for chunk in avatar_file.chunks():
                destination.write(chunk)

        avatar_url = f'/static/images/user_avatars/{user_id}/{filename}'
        user = request.user
        user.avatar_url = avatar_url
        user.save(update_fields=['avatar_url'])

        return JsonResponse({
            'code': 200,
            'msg': '头像上传成功！',
            'data': {'avatar_url': avatar_url}
        })
    except Exception as e:
        print(f"头像上传失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': f'上传失败：{str(e)[:50]}'})

# 知识库管理视图

@login_required(login_url='login')
@require_http_methods(['GET'])
def knowledge_base_view(request):
    """知识库管理首页：展示用户知识库列表"""
    knowledge_bases = KnowledgeBase.objects.filter(user=request.user).order_by('-create_time')
    context = {'knowledge_bases': knowledge_bases, 'username': request.user.username}
    return render(request, 'myapp/knowledge_base.html', context)


def parse_file_content(file):
    """解析上传文件内容：支持txt、md、docx、doc格式"""
    file_name = file.name
    file_ext = os.path.splitext(file_name)[1].lower()
    content = ""

    try:
        file_content = file.read()
        file.seek(0)

        if file_ext == '.txt' or file_ext == '.md':
            detected = chardet.detect(file_content[:1000])
            encoding = detected['encoding'] or 'utf-8'
            try:
                content = file_content.decode(encoding)
            except UnicodeDecodeError:
                content = file_content.decode('gbk', errors='ignore')
        elif file_ext == '.docx':
            doc = Document(file)
            content = '\n'.join([para.text for para in doc.paragraphs])
        elif file_ext == '.doc':
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                temp_file_path = os.path.join(settings.MEDIA_ROOT, f"temp_{os.urandom(8).hex()}.doc")
                with open(temp_file_path, 'wb') as f:
                    f.write(file_content)
                doc = word.Documents.Open(temp_file_path)
                content = doc.Content.Text
                doc.Close(SaveChanges=0)
                os.remove(temp_file_path)
            finally:
                word.Quit()
                pythoncom.CoUninitialize()
        else:
            raise ValueError(f"不支持的文件格式：{file_ext}")

        content = content.strip().replace('\r\n', '\n').replace('\r', '\n')
        return content
    except Exception as e:
        raise ValueError(f"文件解析失败：{str(e)}")


@login_required(login_url='login')
@require_http_methods(['POST'])
def create_knowledge_base(request):
    """创建知识库：支持手动输入或上传文件"""
    kb_file = request.FILES.get('kb_file')
    knowledge_base_name = request.POST.get('kb_name', '').strip()
    knowledge_base_content = request.POST.get('kb_content', '').strip()

    if kb_file:
        try:
            knowledge_base_content = parse_file_content(kb_file)
            if not knowledge_base_name:
                knowledge_base_name = os.path.splitext(kb_file.name)[0]
        except ValueError as e:
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return JsonResponse({'status': 'error', 'message': str(e)})
            messages.error(request, str(e))
            return redirect('knowledge_base')

    if not knowledge_base_name:
        error_msg = "创建失败！知识库名称不能为空"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': error_msg})
        messages.error(request, error_msg)
        return redirect('knowledge_base')

    if not knowledge_base_content:
        error_msg = "创建失败！知识库内容不能为空"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': error_msg})
        messages.error(request, error_msg)
        return redirect('knowledge_base')

    if len(knowledge_base_name) > 50000:
        error_msg = "知识库名称长度不能超过50000个字符！"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': error_msg})
        messages.error(request, error_msg)
        return redirect('knowledge_base')

    if KnowledgeBase.objects.filter(user=request.user, name=knowledge_base_name).exists():
        error_msg = "创建失败！该知识库名称已存在"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': error_msg})
        messages.error(request, error_msg)
        return redirect('knowledge_base')

    knowledge_base_content = knowledge_base_content.replace('<', '&lt;').replace('>', '&gt;')
    try:
        KnowledgeBase.objects.create(
            user=request.user,
            name=knowledge_base_name,
            content=knowledge_base_content
        )
        success_msg = "知识库创建成功！"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'success', 'message': success_msg})
        messages.success(request, success_msg)
    except Exception as e:
        error_msg = f"创建失败：{str(e)}"
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return JsonResponse({'status': 'error', 'message': error_msg})
        messages.error(request, error_msg)
        logger.critical(f"用户[{request.user.username}]创建知识库失败：{error_msg}", exc_info=True)

    return redirect('knowledge_base')


@login_required(login_url='login')
@require_http_methods(['GET', 'POST'])
def edit_knowledge_base(request, kb_id):
    """编辑知识库：支持修改名称、内容，可上传文件替换内容"""
    try:
        kb = KnowledgeBase.objects.get(kb_id=kb_id, user=request.user)
    except KnowledgeBase.DoesNotExist:
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        if is_ajax:
            return JsonResponse({'status': 'error', 'message': "找不到该知识库！"})
        messages.error(request, "找不到该知识库！")
        return redirect('knowledge_base')

    if request.method == 'POST':
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        new_name = request.POST.get('kb_name', '').strip()
        new_content = request.POST.get('kb_content', '').strip()
        kb_file = request.FILES.get('kb_file')

        # 上传文件解析
        if kb_file:
            try:
                new_content = parse_file_content(kb_file)
            except ValueError as e:
                if is_ajax:
                    return JsonResponse({'status': 'error', 'message': str(e)})
                messages.error(request, str(e))
                return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})

        # 基础校验
        if not new_name:
            if is_ajax:
                return JsonResponse({'status': 'error', 'message': "知识库名称不能为空！"})
            messages.error(request, "知识库名称不能为空！")
            return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})

        if not new_content:
            if is_ajax:
                return JsonResponse({'status': 'error', 'message': "知识库内容不能为空！"})
            messages.error(request, "知识库内容不能为空！")
            return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})

        # 名称唯一性校验
        if KnowledgeBase.objects.filter(user=request.user, name=new_name).exclude(kb_id=kb_id).exists():
            if is_ajax:
                return JsonResponse({'status': 'error', 'message': "该知识库名称已存在！"})
            messages.error(request, "该知识库名称已存在！")
            return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})

        # 获取检索配置参数
        chunk_size = request.POST.get('chunk_size', kb.chunk_size)
        chunk_overlap = request.POST.get('chunk_overlap', kb.chunk_overlap)
        similarity_threshold = request.POST.get('similarity_threshold', kb.similarity_threshold)
        top_n = request.POST.get('top_n', kb.top_n)

        # 保存修改
        try:
            kb.name = new_name
            kb.content = new_content.replace('<', '&lt;').replace('>', '&gt;')
            kb.chunk_size = int(chunk_size)
            kb.chunk_overlap = int(chunk_overlap)
            kb.similarity_threshold = float(similarity_threshold)
            kb.top_n = int(top_n)
            kb.save()

            if is_ajax:
                return JsonResponse({'status': 'success', 'message': "知识库修改成功！"})
            messages.success(request, "知识库修改成功！")
            return redirect('knowledge_base')
        except Exception as e:
            if is_ajax:
                return JsonResponse({'status': 'error', 'message': f"修改失败：{str(e)[:50]}"})
            messages.error(request, f"修改失败：{str(e)}")
            return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})

    return render(request, 'myapp/edit_knowledge_base.html', {'kb': kb})


@login_required(login_url='login')
@require_http_methods(['GET'])
def delete_knowledge_base(request, kb_id):
    """删除指定知识库"""
    try:
        kb = KnowledgeBase.objects.get(kb_id=kb_id, user=request.user)
        kb.delete()
        return JsonResponse({'status': 'success', 'message': '知识库删除成功！'})
    except KnowledgeBase.DoesNotExist:
        return JsonResponse({'status': 'error', 'message': '找不到该知识库！'}, status=404)
    except Exception as e:
        return JsonResponse({'status': 'error', 'message': f"删除失败：{str(e)}"}, status=500)


# 语音识别视图

@login_required(login_url='login')
@csrf_exempt
@require_http_methods(['POST'])
def asr_recognize_api(request):
    """语音识别接口：调用ASR服务将音频转文本"""
    if 'audio_file' not in request.FILES:
        return JsonResponse({'code': 400, 'msg': '未上传音频文件', 'data': {'text': ''}})

    try:
        audio_file = request.FILES['audio_file']
        audio_data = audio_file.read()
        audio_format = audio_file.name.split('.')[-1].lower()
        audio_base64 = f"data:audio/{audio_format};base64,{base64.b64encode(audio_data).decode('utf-8')}"

        messages = [
            {"role": "system", "content": [{"text": ""}]},
            {"role": "user", "content": [{"audio": audio_base64}]}
        ]

        start = time.perf_counter()
        response = dashscope.MultiModalConversation.call(
            api_key=DASHSCOPE_API_KEY,
            model="qwen3-asr-flash",
            messages=messages,
            result_format="message",
            asr_options={"language": "zh", "enable_itn": False}
        )
        mdtime = int((time.perf_counter() - start) * 1000)

        recognize_text = ""
        if response.status_code == 200:
            try:
                choices = getattr(response, 'output', {}).get('choices', [])
                if choices:
                    content = choices[0].get('message', {}).get('content', [])
                    recognize_text = content[0].get('text', '') if content else ''
            except (AttributeError, IndexError, KeyError):
                recognize_text = "识别结果格式异常"

        return JsonResponse({
            'code': 200,
            'msg': '识别成功' if recognize_text else '未识别到有效语音内容',
            'data': {'text': recognize_text, "mdtime": mdtime}
        })
    except Exception as e:
        print(f"语音识别失败：{str(e)}")
        return JsonResponse({'code': 500, 'msg': f'识别失败：{str(e)}', 'data': {'text': ''}})