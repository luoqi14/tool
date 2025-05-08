#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Blueprint, request, jsonify, send_from_directory
import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import smtplib
import tempfile
import pandas as pd
import re
import json

# 创建蓝图
email_bp = Blueprint('email', __name__, 
                    url_prefix='/email',
                    static_folder='static',
                    static_url_path='/static',
                    template_folder='templates')

# 企业微信邮箱配置
EMAIL_CONFIG = {
    "sender_email": "qi.luo@jarvismedical.com",
    "auth_code": "pm6reY69tbbS7NhS",
    "smtp_server": "smtp.exmail.qq.com",
    "smtp_port": 465
}

@email_bp.route('/')
@email_bp.route('')
def index():
    """提供前端页面"""
    return send_from_directory(os.path.join(os.path.dirname(__file__), 'templates'), 'index.html')

def extract_emails_from_excel(file_path):
    """从 Excel 文件中提取收件人邮箱地址
    查找“接收账单邮箱”单元格，并获取其下方单元格的值
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 查找“接收账单邮箱”单元格
        email_found = False
        email_address = None
        
        # 遍历DataFrame查找“接收账单邮箱”
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                if "接收账单邮箱" in cell_value:
                    # 如果找到了目标单元格，获取下方单元格的值
                    if i + 1 < len(df):
                        email_address = str(df.iloc[i, j + 1])
                        email_found = True
                        break
            if email_found:
                break
        
        # 验证邮箱格式
        if email_address and re.match(r"[^@]+@[^@]+\.[^@]+", email_address):
            return email_address
        else:
            return None
    except Exception as e:
        print(f"解析Excel文件时出错: {str(e)}")
        return None

@email_bp.route('/api/parse-excel', methods=['POST'])
def parse_excel():
    """解析上传的Excel文件以提取收件人邮箱"""
    try:
        data = request.json
        
        if not data.get('file_data'):
            return jsonify({
                'success': False,
                'message': '缺少文件数据'
            }), 400
        
        # 解析文件数据
        file_data = data['file_data']
        file_name = data.get('file_name', 'attachment.xlsx')
        
        # 如果数据是Base64编码的，需要解码
        if 'base64,' in file_data:
            file_data = file_data.split('base64,')[1]
            
        # 解码Base64数据
        decoded_data = base64.b64decode(file_data)
        
        # 创建临时文件存储数据
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        temp_file.write(decoded_data)
        temp_file.close()
        
        # 从 Excel 文件中提取邮箱地址
        email_address = extract_emails_from_excel(temp_file.name)
        
        # 删除临时文件
        try:
            os.unlink(temp_file.name)
        except:
            pass
        
        if email_address:
            return jsonify({
                'success': True,
                'email': email_address,
                'file_name': file_name
            })
        else:
            return jsonify({
                'success': False,
                'message': '无法从 Excel 文件中提取邮箱地址，请确保文件中包含“接收账单邮箱”字段'
            }), 400
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'解析文件时出错: {str(e)}'
        }), 500

@email_bp.route('/api/send-email', methods=['POST'])
def send_email():
    """处理邮件发送请求"""
    try:
        data = request.json
        
        # 验证必要参数
        if not data.get('subject') or not data.get('content') or not data.get('attachments'):
            return jsonify({
                'success': False,
                'message': '缺少必要参数：主题、内容或附件'
            }), 400
        
        # 解析参数
        subject = data['subject']
        content = data['content']
        attachments = data.get('attachments', [])
        
        # 批量发送结果
        results = []
        
        # 处理每个附件
        for attachment in attachments:
            try:
                # 解析附件数据
                filename = attachment['name']
                file_data = attachment['data']
                recipient_email = attachment.get('email')
                
                # 提取文件名（不包含扩展名）用于占位符替换
                filename_without_ext = os.path.splitext(filename)[0]
                
                if not recipient_email:
                    results.append({
                        'filename': filename,
                        'success': False,
                        'message': '缺少收件人邮箱地址'
                    })
                    continue
                
                # 如果数据是Base64编码的，需要解码
                if 'base64,' in file_data:
                    file_data = file_data.split('base64,')[1]
                    
                # 解码Base64数据
                decoded_data = base64.b64decode(file_data)
                
                # 创建临时文件存储附件数据
                temp_file = tempfile.NamedTemporaryFile(delete=False)
                temp_file.write(decoded_data)
                temp_file.close()
                
                # 替换占位符
                email_subject = subject.replace('[[文件名]]', filename_without_ext)
                email_content = content.replace('[[文件名]]', filename_without_ext)
                
                # 创建邮件对象
                msg = MIMEMultipart()
                msg['From'] = EMAIL_CONFIG['sender_email']
                msg['To'] = recipient_email
                msg['Subject'] = Header(email_subject, 'utf-8')
                
                # 添加邮件正文
                msg.attach(MIMEText(email_content, 'html', 'utf-8'))
                
                # 添加附件到邮件
                with open(temp_file.name, 'rb') as f:
                    attachment_mime = MIMEApplication(f.read())
                    attachment_mime.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(attachment_mime)
                
                # 发送邮件
                try:
                    with smtplib.SMTP_SSL(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port']) as server:
                        server.login(EMAIL_CONFIG['sender_email'], EMAIL_CONFIG['auth_code'])
                        server.sendmail(EMAIL_CONFIG['sender_email'], [recipient_email], msg.as_string())
                    
                    results.append({
                        'filename': filename,
                        'recipient': recipient_email,
                        'success': True,
                        'message': '邮件发送成功'
                    })
                    
                except Exception as e:
                    results.append({
                        'filename': filename,
                        'recipient': recipient_email,
                        'success': False,
                        'message': f'邮件发送失败: {str(e)}'
                    })
                
                # 清理临时文件
                try:
                    os.unlink(temp_file.name)
                except:
                    pass
                    
            except Exception as e:
                results.append({
                    'filename': attachment.get('name', '未知文件'),
                    'success': False,
                    'message': f'处理附件时出错: {str(e)}'
                })
        
        # 返回批量发送结果
        success_count = sum(1 for r in results if r['success'])
        total_count = len(results)
        
        return jsonify({
            'success': success_count > 0,
            'message': f'已成功发送 {success_count}/{total_count} 封邮件',
            'results': results
        })
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'请求处理失败: {str(e)}'
        }), 500
