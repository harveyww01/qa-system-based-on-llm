"""
豆包服务：调用字节跳动豆包大模型API进行文本生成。
"""
import os
import requests
import json


def call_doubao(message, model_config, api_url=None, model_name=None, context_messages=None, api_key=None):
    """
    调用豆包API生成文本。
    
    Args:
        message: 用户输入消息
        model_config: 模型配置参数（temperature, max_tokens等）
        api_url: API地址
        model_name: 模型名称
        context_messages: 上下文消息列表
        api_key: API密钥
    
    Returns:
        tuple: (结果文本, 总耗时ms, TTFT ms, token数)
    """
    import time
    
    # 优先使用传入的api_key，否则从环境变量读取
    if not api_key:
        api_key = os.getenv("DOUBAO_API_KEY", "")
    if not api_key:
        return "错误：豆包API密钥未配置！", 0, 0, 0
    
    # 使用传入的URL或默认URL
    url = api_url or os.getenv("DOUBAO_API_URL", "https://api.doubao.com/v1/chat/completions")
    
    # 使用传入的模型名称或默认名称
    model = model_name or os.getenv("DOUBAO_MODEL_NAME", "Doubao-3.5-Turbo")
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    # 构建消息列表
    messages = [
        {"role": "system", "content": "你是一个专业的中文助手，请始终使用简体中文回答所有问题。"}
    ]
    
    # 添加上下文消息
    if context_messages:
        for ctx_msg in context_messages:
            role = "user" if ctx_msg.get("role") == "user" else "assistant"
            messages.append({"role": role, "content": ctx_msg.get("content", "")})
    
    # 添加当前用户消息
    messages.append({"role": "user", "content": message})
    
    data = {
        "model": model,
        "messages": messages,
        "temperature": model_config.get("temperature", 0.2),
        "max_tokens": model_config.get("max_tokens", 8192),
        "top_p": model_config.get("top_p", 0.9)
    }
    
    try:
        start_time = time.perf_counter()
        response = requests.post(url, headers=headers, json=data, timeout=120)
        total_time_ms = int((time.perf_counter() - start_time) * 1000)
        
        if response.status_code == 200:
            result = response.json()
            if result.get("choices"):
                content = result["choices"][0]["message"]["content"].strip()
                token_count = result.get("usage", {}).get("total_tokens", len(content))
                return content, total_time_ms, total_time_ms // 2, token_count
            else:
                return f"错误：{result.get('error', {}).get('message', '未知错误')}", total_time_ms, total_time_ms, 0
        else:
            return f"错误：API请求失败，状态码: {response.status_code}", total_time_ms, total_time_ms, 0
    
    except requests.exceptions.RequestException as e:
        return f"错误：请求异常 - {str(e)}", 0, 0, 0
    except Exception as e:
        return f"错误：服务器错误 - {str(e)}", 0, 0, 0