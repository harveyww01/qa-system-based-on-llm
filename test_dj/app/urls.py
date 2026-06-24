"""
路由层：定义URL与视图的映射关系，按业务模块组织。
"""
from django.urls import path
from django.views.generic import RedirectView
from . import views
from . import views_statistics

# 基础路由：首页、认证、登出、模型页面
basic_patterns = [
    path('', RedirectView.as_view(url='/index/', permanent=True), name='root'),
    path('register/', views.register_view, name='register'),
    path('login/', views.login_view, name='login'),
    path('forgot-pwd/', views.forgot_pwd_view, name='forgot_pwd'),
    path('send-email-code/', views.send_email_code, name='send_email_code'),
    path('verify-email-code/', views.verify_email_code, name='verify_email_code'),
    path('index/', views.index_view, name='index'),
    path('chat/', views.chat_page_view, name='chat_page'),
    path('model/qwen3-4b/', views.qwen3_4b_page_view, name='qwen3_4b_page'),
    path('model/doubao/', views.doubao_page_view, name='doubao_page'),
    path('model/deepseek/', views.deepseek_page_view, name='deepseek_page'),
    path('logout/', views.logout_view, name='logout'),
    path('set-chat-memory-rounds/', views.set_chat_memory_rounds, name='set_chat_memory_rounds'),
    path('knowledge-base/', views.knowledge_base_view, name='knowledge_base'),
    # 统计中心路由
    path('statistics/', views_statistics.statistics_view, name='statistics'),
    path('statistics/time/', views_statistics.statistics_time_view, name='statistics_time'),
    path('statistics/tokens/', views_statistics.statistics_tokens_view, name='statistics_tokens'),
    path('statistics/preference/', views_statistics.statistics_preference_view, name='statistics_preference'),
    path('statistics/suggestion/', views_statistics.statistics_suggestion_view, name='statistics_suggestion'),
]

# API路由：大模型生成、对话管理、统计分析
api_patterns = [
    path('api/llm/generate/', views.generate_text_api, name='generate_text_api'),
    path('api/llm/multi-generate/', views.multi_model_generate_api, name='multi_model_generate_api'),
    path('api/llm/cloud-models/', views.get_cloud_models, name='get_cloud_models'),
    path('api/llm/cloud-chat/', views.cloud_model_chat_api, name='cloud_model_chat'),
    path('api/user/knowledge-bases/', views.get_user_knowledge_bases, name='get_user_knowledge_bases'),
    path('api/chat/sessions/', views.get_chat_sessions, name='get_chat_sessions'),
    path('api/chat/sessions-by-model/', views.get_chat_sessions_by_model, name='get_chat_sessions_by_model'),
    path('api/chat/messages/', views.get_chat_messages, name='get_chat_messages'),
    path('api/chat/create/', views.create_chat_session, name='create_chat_session'),
    path('api/chat/delete/', views.delete_chat, name='delete_chat'),
    path('api/chat/save-message/', views.save_chat_message, name='save_message'),
    path('api/asr/recognize/', views.asr_recognize_api, name='asr_recognize_api'),
    # 统计分析API
    path('api/statistics/time/', views.get_statistics_time, name='get_statistics_time'),
    path('api/statistics/tokens/', views.get_statistics_tokens, name='get_statistics_tokens'),
    path('api/statistics/preference/', views.get_statistics_preference, name='get_statistics_preference'),
    path('api/statistics/habit-suggestion/', views.get_user_habit_suggestion, name='get_user_habit_suggestion'),
    # 模型对比分析API
    path('api/model/comparison/', views.get_model_comparison, name='get_model_comparison'),
    # 云端模型管理API
    path('api/admin/cloud-model/create/', views.create_cloud_model, name='create_cloud_model'),
    path('api/admin/cloud-model/update/', views.update_cloud_model, name='update_cloud_model'),
    path('api/admin/cloud-model/delete/', views.delete_cloud_model, name='delete_cloud_model'),
]

# 知识库管理路由：创建、编辑、删除知识库
kb_patterns = [
    path('knowledge-base/create/', views.create_knowledge_base, name='create_knowledge_base'),
    path('knowledge-base/edit/<int:kb_id>/', views.edit_knowledge_base, name='edit_knowledge_base'),
    path('knowledge-base/delete/<int:kb_id>/', views.delete_knowledge_base, name='delete_knowledge_base'),
]

# 个人中心路由：资料编辑、密码修改、头像上传、消息收藏
profile_patterns = [
    path('profile/edit/', views.edit_profile, name='edit_profile'),
    path('profile/change_pwd/', views.change_pwd, name='change_pwd'),
    path('api/profile/check-email/', views.check_email_unique, name='check_email_unique'),
    path('api/profile/upload-avatar/', views.upload_avatar, name='upload_avatar'),
    path('collection/', views.collection_view, name='chat_collection'),
    path('api/message/collect/', views.collect_message, name='collect_message'),
    path('api/message/cancel-collect/', views.cancel_collect_message, name='cancel_collect_message'),
    path('api/message/check-collected/', views.check_message_collected, name='check_message_collected'),
    path('api/message/collected-list/', views.get_collected_messages, name='get_collected_messages'),
    path('api/message/batch-collect/', views.batch_collect_messages, name='batch_collect_messages'),
    path('api/message/batch-cancel-collect/', views.batch_cancel_collect_messages, name='batch_cancel_collect_messages'),
]

# 总路由
urlpatterns = basic_patterns + api_patterns + kb_patterns + profile_patterns