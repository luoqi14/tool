#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Blueprint, render_template, send_from_directory, current_app
import os

# 获取当前模块的绝对路径
module_path = os.path.dirname(os.path.abspath(__file__))

# 创建蓝图
file_ysm_bp = Blueprint('file_ysm', __name__, 
                      url_prefix='/file/ysm',
                      template_folder=module_path,
                      static_folder=module_path,
                      static_url_path='')

@file_ysm_bp.route('/')
def index():
    """文件工具首页"""
    return send_from_directory(module_path, 'index.html')

# 提供静态文件
@file_ysm_bp.route('/<path:filename>')
def static_files(filename):
    return send_from_directory(module_path, filename)
