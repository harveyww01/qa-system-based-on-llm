"""
应用配置：定义Django应用基本配置。
"""
from django.apps import AppConfig


class AppConfig(AppConfig):
    """主应用配置"""
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'app'
