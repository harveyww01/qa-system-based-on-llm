"""
统计中心视图：提供模型使用统计、性能分析、用户习惯建议等可视化页面。
"""
from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.utils import timezone
from datetime import timedelta
import json

from app.models import ModelCallLog


@login_required(login_url='login')
def statistics_view(request):
    """统计中心首页"""
    return render(request, 'myapp/statistics/index.html')


@login_required(login_url='login')
def statistics_time_view(request):
    """输出时间统计页"""
    return render(request, 'myapp/statistics/time_statistics.html')


@login_required(login_url='login')
def statistics_tokens_view(request):
    """Token统计页"""
    return render(request, 'myapp/statistics/token_statistics.html')


@login_required(login_url='login')
def statistics_preference_view(request):
    """模型偏好统计页"""
    return render(request, 'myapp/statistics/preference_statistics.html')


@login_required(login_url='login')
def statistics_suggestion_view(request):
    """用户习惯建议页"""
    return render(request, 'myapp/statistics/suggestion.html')