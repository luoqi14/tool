#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Flask, render_template, jsonify
from flask_cors import CORS
import os
import argparse

# 导入模块
from modules.email.routes import email_bp

# 创建应用
app = Flask(__name__)
CORS(app)  # 启用CORS，允许跨域请求

# 注册蓝图
app.register_blueprint(email_bp)

# 根路径重定向到工具列表页面
@app.route('/')
def index():
    """工具列表页面"""
    return jsonify({
        'status': 'ok',
        'message': 'Jarvis工具集API服务正在运行',
        'tools': [
            {
                'name': '邮件发送工具',
                'path': '/email',
                'description': '用于发送企业邮件的工具'
            }
            # 未来可以在这里添加更多工具
        ]
    })

# 错误处理
@app.errorhandler(404)
def page_not_found(e):
    return jsonify({
        'status': 'error',
        'message': '请求的资源不存在'
    }), 404

@app.errorhandler(500)
def internal_server_error(e):
    return jsonify({
        'status': 'error',
        'message': '服务器内部错误'
    }), 500

if __name__ == '__main__':
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='启动工具集应用')
    parser.add_argument('--port', type=int, default=3001, help='指定运行端口，默认为3001')
    parser.add_argument('--host', type=str, default='0.0.0.0', help='指定运行主机，默认为0.0.0.0')
    parser.add_argument('--debug', action='store_true', help='是否开启调试模式')
    args = parser.parse_args()
    
    # 启动Flask应用
    app.run(host=args.host, port=args.port, debug=True)
