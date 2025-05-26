#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Blueprint, send_from_directory
import os
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("file_supplier")

# 获取当前模块的绝对路径
module_path = os.path.dirname(os.path.abspath(__file__))

# 创建蓝图
file_supplier_bp = Blueprint("file_supplier", __name__, 
                          url_prefix="/file/supplier",
                          template_folder=module_path,
                          static_folder=os.path.join(module_path, "static"),
                          static_url_path="/static")

@file_supplier_bp.route("/")
def index():
    """供应商账单拆分工具首页"""
    return send_from_directory(module_path, "index.html")

# 提供静态文件
@file_supplier_bp.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(module_path, filename)

# 注意：所有文件处理都在浏览器端完成，不需要服务器端API
