#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Flask, render_template, jsonify
from flask_cors import CORS
import os
import argparse

# 导入模块
from modules.email.routes import email_bp
from modules.file_ysm.routes import file_ysm_bp
from modules.file_supplier.routes import file_supplier_bp

# 创建Flask应用
app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)  # 启用CORS，允许跨域请求

# 注册蓝图
app.register_blueprint(email_bp)
app.register_blueprint(file_ysm_bp)
app.register_blueprint(file_supplier_bp)

# 根路径显示工具列表页面
@app.route('/')
def index():
    """工具列表页面"""
    tools = [
        {
            'name': '邮件发送工具',
            'path': '/email',
            'description': '用于群发企业邮件的工具',
            'icon': 'bi-envelope',
            'image': 'email-tool.png'
        },
        {
            'name': '隐适美账单分割',
            'path': '/file/ysm',
            'description': 'excel账单文件拆分成需要的文件',
            'icon': 'bi-file-earmark',
            'image': 'file-tool.png'
        },
        {
            'name': '供应商账单拆分',
            'path': '/file/supplier',
            'description': '按供应商拆分线下服务费账单',
            'icon': 'bi-file-earmark-spreadsheet',
            'image': 'file-tool.png'
        }
        # 未来可以在这里添加更多工具
    ]
    return render_template('index.html', tools=tools)

# 提供API版本的工具列表
@app.route('/api')
def api_index():
    """工具列表API"""
    return jsonify({
        'status': 'ok',
        'message': 'Jarvis工具集API服务正在运行',
        'tools': [
            {
                'name': '邮件发送工具',
                'path': '/email',
                'description': '用于发送企业邮件的工具'
            },
            {
                'name': '隐适美账单分割',
                'path': '/file/ysm',
                'description': 'excel账单文件拆分成需要的文件'
            },
            {
                'name': '供应商账单拆分',
                'path': '/file/supplier',
                'description': '按供应商拆分线下服务费账单'
            }
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
