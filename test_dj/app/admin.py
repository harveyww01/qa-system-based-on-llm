"""
后台管理配置：注册模型并配置展示和权限。
"""
from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.admin.models import LogEntry
from .models import User, VerifyCode, Chat, Message, KnowledgeBase


class LogEntryAdmin(admin.ModelAdmin):
    """日志管理：仅超级管理员可见"""
    list_display = ('action_time', 'user', 'content_type', 'object_repr', 'action_flag', 'change_message')
    list_filter = ('action_flag', 'content_type')
    search_fields = ('user__username', 'object_repr', 'change_message')
    readonly_fields = ('action_time', 'user', 'content_type', 'object_repr', 'action_flag', 'change_message')

    def has_module_permission(self, request):
        return request.user.is_superuser

    def has_view_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


class CustomUserAdmin(UserAdmin):
    """用户管理：权限分离"""
    list_display = ('pk', 'username', 'email', 'avatar_url', 'register_time', 'last_login', 'is_staff', 'is_superuser')
    search_fields = ('username', 'email')
    list_filter = ('is_staff', 'is_superuser')
    readonly_fields = ('register_time',)

    def get_fieldsets(self, request, obj=None):
        """动态字段集：超级管理员可见权限分配部分"""
        basic_fieldsets = [
            ('个人信息', {'fields': ('username', 'email', 'password')}),
            ('状态设置', {'fields': ('is_active', 'is_staff')}),
        ]
        if request.user.is_superuser:
            basic_fieldsets.extend([
                ('超级权限', {'fields': ('is_superuser',)}),
                ('用户权限', {'fields': ('user_permissions',)}),
            ])
        return basic_fieldsets

    def has_view_permission(self, request, obj=None):
        return request.user.is_staff

    def has_change_permission(self, request, obj=None):
        """平台管理员不能修改超级管理员账号"""
        if obj and obj.is_superuser:
            return False
        return request.user.is_staff

    def has_add_permission(self, request):
        return request.user.is_superuser

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

    def changelist_view(self, request, extra_context=None):
        extra_context = extra_context or {}
        extra_context['title'] = "Django 站点管理员"
        return super().changelist_view(request, extra_context=extra_context)


class VerifyCodeAdmin(admin.ModelAdmin):
    """验证码管理"""
    list_display = ('id', 'email', 'code', 'expire_time', 'is_used', 'created_time', 'user')
    search_fields = ('email', 'code', 'user__username')
    list_filter = ('is_used',)
    readonly_fields = ('id', 'user', 'email', 'code', 'created_time', 'expire_time', 'is_used')


class ChatAdmin(admin.ModelAdmin):
    """对话管理"""
    list_display = ('chat_id', 'user', 'chat_title', 'create_time', 'update_time')
    search_fields = ('user__username', 'chat_title')
    readonly_fields = ('create_time', 'update_time')


class MessageAdmin(admin.ModelAdmin):
    """消息管理"""
    list_display = ('msg_id', 'user', 'chat', 'role', 'content', 'send_time')
    search_fields = ('chat__chat_title', 'content', 'model__model_name')
    list_filter = ('role',)
    readonly_fields = ('send_time',)


class KnowledgeBaseAdmin(admin.ModelAdmin):
    """知识库管理"""
    list_display = ('kb_id', 'user', 'name', 'create_time', 'update_time')
    search_fields = ('user__username', 'name', 'content')
    readonly_fields = ('create_time', 'update_time')


# 注册模型到后台
admin.site.register(LogEntry, LogEntryAdmin)
admin.site.register(User, CustomUserAdmin)
admin.site.register(VerifyCode, VerifyCodeAdmin)
admin.site.register(Chat, ChatAdmin)
admin.site.register(Message, MessageAdmin)
admin.site.register(KnowledgeBase, KnowledgeBaseAdmin)