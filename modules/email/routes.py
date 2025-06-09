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
import docx
import pathlib

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

def extract_emails_from_word(file_path):
    """从 Word 文档中提取收件人邮箱地址
    查找"接收邮箱："字段，并获取其后面的值
    """
    try:
        # 读取Word文档
        doc = docx.Document(file_path)
        email_address = None
        
        # 遍历所有段落
        for para in doc.paragraphs:
            text = para.text.strip()
            if "接收邮箱：" in text:
                # 提取邮箱地址
                parts = text.split("接收邮箱：", 1)
                if len(parts) > 1:
                    potential_email = parts[1].strip()
                    # 验证邮箱格式
                    if re.match(r"[^@]+@[^@]+\.[^@]+", potential_email):
                        email_address = potential_email
                        break
        
        return {
            'email': email_address,
            'is_word': True
        }
    except Exception as e:
        print(f"解析Word文档时出错: {str(e)}")
        return {
            'email': None,
            'is_word': True
        }

def extract_emails_from_excel(file_path):
    """从 Excel 文件中提取收件人邮箱地址和对账期间
    查找"接收账单邮箱"或"接收邮箱"和"对账期间"单元格，并获取其下方单元格的值
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 查找"接收账单邮箱"、"对账期间"、"结算期间"和"公司"单元格
        email_found = False
        period_found = False
        settlement_found = False
        company_found = False
        email_address = None
        period_value = None
        period_range_value = None
        company_value = None
        
        # 遍历DataFrame查找"接收账单邮箱"和"对账期间"
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                
                # 查找"接收账单邮箱"或"接收邮箱"
                if ("接收账单邮箱" in cell_value or "接收邮箱" in cell_value or "账单收件箱" in cell_value) and not email_found:
                    # 先检查右侧单元格
                    if j + 1 < len(df.columns):
                        right_value = str(df.iloc[i, j + 1])
                        if re.match(r"[^@]+@[^@]+\.[^@]+", right_value):
                            email_address = right_value
                            email_found = True
                            continue
                    
                    # 再检查下方单元格
                    if i + 1 < len(df) and not email_found:
                        down_value = str(df.iloc[i + 1, j])
                        if re.match(r"[^@]+@[^@]+\.[^@]+", down_value):
                            email_address = down_value
                            email_found = True
                
                # 查找"对账期间"
                if "对账期间" in cell_value and not period_found:
                    # 如果找到了目标单元格，获取下方单元格的值
                    if i + 1 < len(df):
                        period_value = str(df.iloc[i + 1, j])
                        period_found = True
                
                # 查找"结算期间"
                if "结算期间" in cell_value and not settlement_found:
                    # 如果找到了目标单元格，获取下方单元格的值
                    if i + 1 < len(df):
                        # 获取同行下一列的值作为年月日范围
                        period_range_value = ""
                        if j + 1 < len(df.columns):
                            period_range_value = str(df.iloc[i, j + 1])
                        settlement_found = True
                        
                # 查找"公司："
                if "公司：" in cell_value and not company_found:
                    # 提取"公司："后面的值
                    company_value = cell_value.split("公司：", 1)[1].strip()
                    company_found = True
                
                # 如果所有值都找到了，可以提前结束循环
                if email_found and period_found and settlement_found and company_found:
                    break
            if email_found and period_found and settlement_found and company_found:
                break
        
        # 验证邮箱格式
        if email_address and re.match(r"[^@]+@[^@]+\.[^@]+", email_address):
            return {
                'email': email_address,
                'period': period_value if period_found else '',
                'period_range': period_range_value if settlement_found else '',
                'company': company_value if company_found else ''
            }
        else:
            return {
                'email': None,
                'period': period_value if period_found else '',
                'period_range': period_range_value if settlement_found else '',
                'company': company_value if company_found else ''
            }
    except Exception as e:
        print(f"解析Excel文件时出错: {str(e)}")
        return {
            'email': None,
            'period': '',
            'period_range': '',
            'company': ''
        }

@email_bp.route('/api/parse-file', methods=['POST'])
def parse_file():
    """解析上传的Excel或Word文件以提取收件人邮箱和其他信息"""
    try:
        data = request.json
        if not data or not data.get('file_data') or not data.get('file_name'):
            return jsonify({
                'success': False,
                'message': '缺少文件数据或文件名'
            }), 400
            
        # 获取文件名和数据
        filename = data['file_name']
        file_data = data['file_data']
        
        # 如果数据是Base64编码的，需要解码
        if 'base64,' in file_data:
            file_data = file_data.split('base64,')[1]
            
        # 解码Base64数据
        decoded_data = base64.b64decode(file_data)
        
        # 创建临时文件存储数据
        temp_file = tempfile.NamedTemporaryFile(delete=False)
        temp_file.write(decoded_data)
        temp_file.close()
        
        # 根据文件扩展名决定如何解析
        file_ext = filename.split('.')[-1].lower()
        
        # 获取文件名（不包含扩展名）用于匹配
        filename_without_ext = os.path.splitext(filename)[0]
        
        if file_ext in ['doc', 'docx']:
            # 解析Word文档
            result = extract_emails_from_word(temp_file.name)
            result['filename'] = filename
            result['filename_without_ext'] = filename_without_ext
        else:
            # 解析Excel文件
            result = extract_emails_from_excel(temp_file.name)
            result['filename'] = filename
            result['filename_without_ext'] = filename_without_ext
        
        # 删除临时文件
        os.unlink(temp_file.name)
        
        if result['email']:
            return jsonify({
                'success': True,
                'email': result['email'],
                'period': result.get('period', ''),
                'period_range': result.get('period_range', ''),
                'company': result.get('company', ''),
                'is_word': result.get('is_word', False),
                'filename': filename,
                'filename_without_ext': filename_without_ext
            })
        else:
            file_type = 'Word文档' if file_ext in ['doc', 'docx'] else 'Excel文件'
            return jsonify({
                'success': False,
                'message': f'无法从{file_type}中提取邮箱地址，请确保文件中包含正确的邮箱字段'
            }), 400
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'解析文件时出错: {str(e)}'
        }), 500

# 保留旧的API端点以保持兼容性
@email_bp.route('/api/parse-excel', methods=['POST'])
def parse_excel():
    """解析上传的Excel文件以提取收件人邮箱和对账期间（兼容旧版本）"""
    return parse_file()

@email_bp.route('/api/send-email', methods=['POST'])
def send_email():
    """处理邮件发送请求"""
    try:
        data = request.json
        
        # 验证必要参数
        if not data.get('subject') or not data.get('content') or not data.get('attachments') or not data.get('senderEmail') or not data.get('senderPassword'):
            return jsonify({
                'success': False,
                'message': '缺少必要参数：主题、内容、附件、发送邮箱或密码'
            }), 400
        
        # 解析参数
        subject = data['subject']
        content = data['content']
        sender_email = data['senderEmail']
        sender_password = data['senderPassword']
        cc_sender = data.get('ccSender', True)  # 默认为 True，即发送副本给自己
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
                period_value = attachment.get('period', '')  # 获取对账期间
                is_word = attachment.get('is_word', False)  # 是否为Word文档
                word_attachment = attachment.get('word_attachment', None)  # 对应的Word附件
                
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
                email_subject = email_subject.replace('[[年月]]', period_value)
                email_subject = email_subject.replace('[[年月日范围]]', attachment.get('period_range', ''))
                email_subject = email_subject.replace('[[公司]]', attachment.get('company', ''))
                email_content = content.replace('[[文件名]]', filename_without_ext)
                email_content = email_content.replace('[[年月]]', period_value)
                email_content = email_content.replace('[[年月日范围]]', attachment.get('period_range', ''))
                email_content = email_content.replace('[[公司]]', attachment.get('company', ''))
                
                # 创建邮件对象
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = Header(email_subject, 'utf-8')
                # 如果用户选择将邮件副本发送给自己，添加发送方为密送收件人
                if cc_sender:
                    msg['Bcc'] = sender_email
                
                # 添加邮件正文
                msg.attach(MIMEText(email_content, 'html', 'utf-8'))
                
                # 添加主附件到邮件
                with open(temp_file.name, 'rb') as f:
                    attachment_mime = MIMEApplication(f.read())
                    attachment_mime.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(attachment_mime)
                
                # 如果有对应的Word附件，也添加到邮件中
                if word_attachment:
                    # 解码Word附件的Base64数据
                    word_file_data = word_attachment['data']
                    if 'base64,' in word_file_data:
                        word_file_data = word_file_data.split('base64,')[1]
                    word_decoded_data = base64.b64decode(word_file_data)
                    
                    # 创建临时文件存储Word附件数据
                    word_temp_file = tempfile.NamedTemporaryFile(delete=False)
                    word_temp_file.write(word_decoded_data)
                    word_temp_file.close()
                    
                    # 添加Word附件到邮件
                    with open(word_temp_file.name, 'rb') as f:
                        word_attachment_mime = MIMEApplication(f.read())
                        word_attachment_mime.add_header('Content-Disposition', 'attachment', filename=word_attachment['name'])
                        msg.attach(word_attachment_mime)
                    
                    # 删除临时文件
                    try:
                        os.unlink(word_temp_file.name)
                    except:
                        pass
                
                # 发送邮件
                try:
                    # 从邮箱地址推断 SMTP 服务器
                    smtp_domain = sender_email.split('@')[1]
                    smtp_server = f'smtp.{smtp_domain}'
                    
                    # 如果是企业微信邮箱，使用特定的服务器
                    if smtp_domain == 'exmail.qq.com' or smtp_domain == 'jarvismedical.com':
                        smtp_server = 'smtp.exmail.qq.com'
                    
                    with smtplib.SMTP_SSL(smtp_server, 465) as server:
                        server.login(sender_email, sender_password)
                        # 准备收件人列表
                        recipients = [recipient_email]
                        
                        # 如果用户选择将邮件副本发送给自己，将发送方也添加到收件人列表中
                        if cc_sender:
                            recipients.append(sender_email)
                            
                        server.sendmail(sender_email, recipients, msg.as_string())
                    
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
