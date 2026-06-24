"""
模型层：定义应用的核心数据结构，包含用户认证、对话管理、知识库、模型调用日志等实体。
"""
import os
import shutil
import random
import string
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.contrib.auth import get_user_model
from django.utils import timezone
from django.conf import settings


class User(AbstractUser):
    """自定义用户模型：扩展默认用户模型，添加头像字段"""
    username = models.CharField(max_length=50, unique=True, null=False, blank=False, verbose_name="用户名")
    email = models.EmailField(max_length=100, unique=True, null=False, blank=False, verbose_name="邮箱")
    avatar_url = models.CharField(max_length=255, null=True, blank=True, default='/static/images/avatar-default.jpg', verbose_name="头像链接")
    register_time = models.DateTimeField(auto_now_add=True, verbose_name="注册时间")

    class Meta:
        db_table = "user"
        verbose_name = "用户"
        verbose_name_plural = "用户"
        ordering = ["-register_time"]

    def __str__(self):
        return f"[{self.id}]{self.username}"

    def delete(self, *args, **kwargs):
        """删除用户时自动清理头像文件目录"""
        # 获取头像存储目录路径
        avatar_root_dir = os.path.join(settings.STATICFILES_DIRS[0], 'images', 'user_avatars')
        user_avatar_dir = os.path.join(avatar_root_dir, str(self.id))

        # 如果是非默认头像且目录存在，删除头像目录
        if self.avatar_url and self.avatar_url != '/static/images/avatar-default.jpg':
            if os.path.exists(user_avatar_dir):
                try:
                    shutil.rmtree(user_avatar_dir)
                    print(f"[用户删除] 已清理用户 {self.username} 的头像目录: {user_avatar_dir}")
                except Exception as e:
                    print(f"[用户删除] 清理头像目录失败 {user_avatar_dir}: {str(e)}")

        # 调用父类的delete方法删除用户记录
        super().delete(*args, **kwargs)


class VerifyCode(models.Model):
    """验证码模型：存储邮箱验证码，支持注册、找回密码、修改密码场景"""
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, null=True, blank=True, verbose_name="关联用户")
    email = models.EmailField(verbose_name="接收邮箱")
    code = models.CharField(max_length=6, verbose_name="验证码")
    expire_time = models.DateTimeField(verbose_name="过期时间")
    is_used = models.BooleanField(default=False, verbose_name="是否已使用")
    created_time = models.DateTimeField(auto_now_add=True, verbose_name="创建时间")

    class Meta:
        verbose_name = "验证码"
        verbose_name_plural = "验证码"
        indexes = [models.Index(fields=["email", "is_used", "expire_time"])]
        unique_together = ["email", "code", "is_used"]

    def __str__(self):
        return f"{self.email} - {self.code} | 有效至：{self.expire_time.strftime('%Y-%m-%d %H:%M')}"

    @staticmethod
    def generate_code():
        """生成6位数字验证码"""
        return ''.join(random.choices(string.digits, k=6))


class Chat(models.Model):
    """对话会话模型：存储用户的对话会话，包含标题、创建时间、上下文记忆轮数"""
    chat_id = models.AutoField(primary_key=True, verbose_name="对话ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    chat_title = models.CharField(max_length=100, null=False, blank=False, verbose_name="对话标题")
    create_time = models.DateTimeField(auto_now_add=True, null=False, verbose_name="创建时间")
    update_time = models.DateTimeField(auto_now=True, null=False, verbose_name="最后消息时间")
    memory_rounds = models.IntegerField(default=0, verbose_name="上下文记忆轮数")
    model_name = models.CharField(max_length=100, null=True, blank=True, verbose_name="模型名称")

    class Meta:
        db_table = "chat"
        verbose_name = "对话"
        verbose_name_plural = "对话"
        ordering = ["-update_time"]
        indexes = [models.Index(fields=["user", "-update_time"]), models.Index(fields=["user", "model_name"])]

    def __str__(self):
        return f"[{self.user.username}]{self.chat_title}"


class Message(models.Model):
    """对话消息模型：存储对话中的单条消息，记录角色、内容、调用模式和性能指标"""
    CALL_MODE_CHOICES = [
        ("local", "本地模型模式"),
        ("cloud", "云端模型模式"),
        ("compare", "对比模式"),
    ]

    msg_id = models.BigAutoField(primary_key=True, verbose_name="消息ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    chat = models.ForeignKey(Chat, on_delete=models.CASCADE, db_column="chat_id", verbose_name="所属对话")
    role = models.IntegerField(choices=[(1, 'user'), (2, 'assistant')], default=1, null=False, blank=False, verbose_name="消息角色")
    content = models.TextField(null=False, blank=False, verbose_name="消息内容")
    send_time = models.DateTimeField(auto_now_add=True, null=False, verbose_name="发送时间")
    round_num = models.IntegerField(default=1, verbose_name="会话轮次")
    model_type = models.CharField(max_length=20, null=True, blank=True, default=None, verbose_name="模型类型")
    cloud_model_id = models.CharField(max_length=100, null=True, blank=True, default=None, verbose_name="云端模型ID")
    call_mode = models.CharField(max_length=20, choices=CALL_MODE_CHOICES, null=True, blank=True, verbose_name="调用模式")
    total_time_ms = models.FloatField(null=True, blank=True, default=None, verbose_name="总响应时间(ms)")
    ttft_ms = models.FloatField(null=True, blank=True, default=None, verbose_name="首Token时间(ms)")
    token_count = models.IntegerField(null=True, blank=True, default=None, verbose_name="Token数量")

    class Meta:
        db_table = "message"
        verbose_name = "消息"
        verbose_name_plural = "消息"
        ordering = ["send_time"]

    def __str__(self):
        return f"{self.chat.chat_title} - {self.get_role_display()}：{self.content[:20]}..."


class MessageCollection(models.Model):
    """消息收藏模型：记录用户收藏的消息，支持快速收藏和取消收藏"""
    collection_id = models.AutoField(primary_key=True, verbose_name="收藏ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    message = models.ForeignKey(Message, on_delete=models.CASCADE, db_column="msg_id", verbose_name="收藏的消息")
    collect_time = models.DateTimeField(auto_now_add=True, verbose_name="收藏时间")

    class Meta:
        db_table = "message_collection"
        verbose_name = "消息收藏"
        verbose_name_plural = "消息收藏"
        ordering = ["-collect_time"]
        unique_together = ["user", "message"]
        indexes = [models.Index(fields=["user", "-collect_time"])]

    def __str__(self):
        return f"[{self.user.username}]收藏了消息：{self.message.content[:20]}..."


class DoubaoChat(models.Model):
    """豆包对话关联模型：记录用户使用豆包模型的对话会话"""
    id = models.AutoField(primary_key=True, verbose_name="ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    chat = models.ForeignKey(Chat, on_delete=models.CASCADE, db_column="chat_id", verbose_name="关联对话")
    model_name = models.CharField(max_length=100, null=False, blank=False, verbose_name="模型名称")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="关联时间")

    class Meta:
        db_table = "doubao_chat"
        verbose_name = "豆包对话"
        verbose_name_plural = "豆包对话"
        ordering = ["-create_time"]
        unique_together = ["user", "chat", "model_name"]
        indexes = [models.Index(fields=["user", "-create_time"]), models.Index(fields=["chat", "model_name"])]

    def __str__(self):
        return f"[{self.user.username}]豆包对话：{self.chat.chat_title}"


class DeepSeekChat(models.Model):
    """DeepSeek对话关联模型：记录用户使用DeepSeek模型的对话会话"""
    id = models.AutoField(primary_key=True, verbose_name="ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    chat = models.ForeignKey(Chat, on_delete=models.CASCADE, db_column="chat_id", verbose_name="关联对话")
    model_name = models.CharField(max_length=100, null=False, blank=False, verbose_name="模型名称")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="关联时间")

    class Meta:
        db_table = "deepseek_chat"
        verbose_name = "DeepSeek对话"
        verbose_name_plural = "DeepSeek对话"
        ordering = ["-create_time"]
        unique_together = ["user", "chat", "model_name"]
        indexes = [models.Index(fields=["user", "-create_time"]), models.Index(fields=["chat", "model_name"])]

    def __str__(self):
        return f"[{self.user.username}]DeepSeek对话：{self.chat.chat_title}"


class Qwen3_4BChat(models.Model):
    """Qwen3_4B对话关联模型：记录用户使用Qwen3-4B本地模型的对话会话"""
    id = models.AutoField(primary_key=True, verbose_name="ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    chat = models.ForeignKey(Chat, on_delete=models.CASCADE, db_column="chat_id", verbose_name="关联对话")
    model_name = models.CharField(max_length=100, null=False, blank=False, verbose_name="模型名称")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="关联时间")

    class Meta:
        db_table = "qwen3_4b_chat"
        verbose_name = "Qwen3_4B对话"
        verbose_name_plural = "Qwen3_4B对话"
        ordering = ["-create_time"]
        unique_together = ["user", "chat", "model_name"]
        indexes = [models.Index(fields=["user", "-create_time"]), models.Index(fields=["chat", "model_name"])]

    def __str__(self):
        return f"[{self.user.username}]Qwen3_4B对话：{self.chat.chat_title}"


class KnowledgeBase(models.Model):
    """知识库模型：存储用户自定义知识库内容，支持RAG检索增强问答"""
    kb_id = models.AutoField(primary_key=True, verbose_name="知识库ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户", related_name="knowledge_bases")
    name = models.CharField(max_length=100, null=False, blank=False, verbose_name="知识库名称")
    content = models.TextField(null=False, blank=False, verbose_name="知识库内容", help_text="支持Markdown格式")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="创建时间")
    update_time = models.DateTimeField(auto_now=True, verbose_name="更新时间")
    similarity_threshold = models.FloatField(default=0.3, verbose_name="相似度阈值")
    top_n = models.IntegerField(default=3, verbose_name="Top-N")
    chunk_size = models.IntegerField(default=300, verbose_name="分段长度")
    chunk_overlap = models.IntegerField(default=50, verbose_name="重叠长度")

    class Meta:
        db_table = "knowledge_base"
        verbose_name = "知识库"
        verbose_name_plural = "知识库"
        ordering = ["-update_time"]
        indexes = [models.Index(fields=["user", "-update_time"])]

    def __str__(self):
        return f"[{self.user.username}]{self.name}"


class CloudModel(models.Model):
    """云端模型配置模型：管理云端大模型的API配置和状态"""
    MODEL_TYPE_CHOICES = [
        ("doubao", "豆包"),
        ("deepseek", "DeepSeek"),
        ("other", "其他"),
    ]

    model_id = models.AutoField(primary_key=True, verbose_name="模型ID")
    name = models.CharField(max_length=100, null=False, blank=False, verbose_name="模型名称")
    model_type = models.CharField(max_length=20, choices=MODEL_TYPE_CHOICES, default="other", verbose_name="模型类型")
    api_key = models.CharField(max_length=255, null=True, blank=True, verbose_name="API密钥")
    api_url = models.CharField(max_length=500, null=True, blank=True, verbose_name="API地址")
    model_name = models.CharField(max_length=100, null=True, blank=True, verbose_name="模型标识")
    is_active = models.BooleanField(default=True, verbose_name="是否启用")
    sort_order = models.IntegerField(default=0, verbose_name="排序顺序")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="创建时间")
    update_time = models.DateTimeField(auto_now=True, verbose_name="更新时间")

    class Meta:
        db_table = "cloud_model"
        verbose_name = "云端模型配置"
        verbose_name_plural = "云端模型配置"
        ordering = ["sort_order", "name"]

    def __str__(self):
        return f"{self.name} ({self.get_model_type_display()})"


class ModelCallLog(models.Model):
    """模型调用日志模型：记录模型调用的性能数据，支持统计分析和性能监控"""
    CALL_MODE_CHOICES = [
        ("local", "仅本地模式"),
        ("cloud", "仅云端模式"),
        ("compare", "对比模式"),
    ]

    log_id = models.AutoField(primary_key=True, verbose_name="日志ID")
    user = models.ForeignKey(get_user_model(), on_delete=models.CASCADE, db_column="user_id", verbose_name="所属用户")
    model_name = models.CharField(max_length=100, verbose_name="模型名称")
    call_mode = models.CharField(max_length=20, choices=CALL_MODE_CHOICES, verbose_name="调用模式")
    input_text = models.TextField(null=True, blank=True, verbose_name="输入文本")
    output_text = models.TextField(null=True, blank=True, verbose_name="输出文本")
    total_time_ms = models.IntegerField(default=0, verbose_name="总耗时(毫秒)")
    ttft_ms = models.IntegerField(default=0, verbose_name="首Token时间(毫秒)")
    token_count = models.IntegerField(default=0, verbose_name="输出Token数")
    success = models.BooleanField(default=True, verbose_name="是否成功")
    error_message = models.TextField(null=True, blank=True, verbose_name="错误信息")
    call_time = models.DateTimeField(auto_now_add=True, verbose_name="调用时间")

    class Meta:
        db_table = "model_call_log"
        verbose_name = "模型调用日志"
        verbose_name_plural = "模型调用日志"
        ordering = ["-call_time"]
        indexes = [models.Index(fields=["user", "-call_time"]), models.Index(fields=["model_name", "call_time"])]

    def __str__(self):
        return f"[{self.user.username}] {self.model_name} - {self.call_mode}"
