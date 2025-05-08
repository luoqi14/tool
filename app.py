#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import smtplib
import tempfile

app = Flask(__name__, static_folder='static', static_url_path='/static')
CORS(app)  # 启用CORS，允许跨域请求

# 注释掉这行，让应用在根路径运行
# app.config['APPLICATION_ROOT'] = '/email'

# 企业微信邮箱配置
EMAIL_CONFIG = {
    "sender_email": "qi.luo@jarvismedical.com",
    "auth_code": "pm6reY69tbbS7NhS",
    "smtp_server": "smtp.exmail.qq.com",
    "smtp_port": 465
}

@app.route('/')
@app.route('/index')
def index():
    """提供前端页面"""
    return send_from_directory('.', 'index.html')

@app.route('/static/<path:path>')
def serve_static(path):
    """提供静态文件"""
    return send_from_directory('static', path)

@app.route('/api/send-email', methods=['POST'])
def send_email():
    """处理邮件发送请求"""
    try:
        data = request.json
        
        # 验证必要参数
        if not data.get('recipients') or not data.get('subject') or not data.get('content'):
            return jsonify({
                'success': False,
                'message': '缺少必要参数：收件人、主题或内容'
            }), 400
        
        # 解析参数
        recipients = data['recipients'].split(',')
        recipients = [r.strip() for r in recipients]
        subject = data['subject']
        content = data['content']
        content_type = data.get('contentType', 'html')
        attachments = data.get('attachments', [])
        
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG['sender_email']
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = Header(subject, 'utf-8')
        
        # 添加邮件正文
        msg.attach(MIMEText(content, content_type, 'utf-8'))
        
        # 处理附件
        temp_files = []  # 用于跟踪临时文件以便后续删除
        
        for attachment in attachments:
            try:
                # 解析附件数据
                filename = attachment['name']
                file_data = attachment['data']
                
                # 如果数据是Base64编码的，需要解码
                if 'base64,' in file_data:
                    file_data = file_data.split('base64,')[1]
                    
                # 解码Base64数据
                decoded_data = base64.b64decode(file_data)
                
                # 创建临时文件存储附件数据
                temp_file = tempfile.NamedTemporaryFile(delete=False)
                temp_file.write(decoded_data)
                temp_file.close()
                temp_files.append(temp_file.name)
                
                # 添加附件到邮件
                with open(temp_file.name, 'rb') as f:
                    attachment_mime = MIMEApplication(f.read())
                    attachment_mime.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(attachment_mime)
                    
            except Exception as e:
                return jsonify({
                    'success': False,
                    'message': f'处理附件 {filename} 时出错: {str(e)}'
                }), 400
        
        # 发送邮件
        try:
            with smtplib.SMTP_SSL(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port']) as server:
                server.login(EMAIL_CONFIG['sender_email'], EMAIL_CONFIG['auth_code'])
                server.sendmail(EMAIL_CONFIG['sender_email'], recipients, msg.as_string())
            
            # 清理临时文件
            for temp_file in temp_files:
                try:
                    os.unlink(temp_file)
                except:
                    pass
                
            return jsonify({
                'success': True,
                'message': '邮件发送成功'
            })
            
        except Exception as e:
            # 清理临时文件
            for temp_file in temp_files:
                try:
                    os.unlink(temp_file)
                except:
                    pass
                
            return jsonify({
                'success': False,
                'message': f'邮件发送失败: {str(e)}'
            }), 500
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'请求处理失败: {str(e)}'
        }), 500

if __name__ == '__main__':
    # 确保静态文件夹存在
    os.makedirs('static', exist_ok=True)
    
    # 启动Flask应用
    app.run(host='0.0.0.0', port=3000, debug=True)
