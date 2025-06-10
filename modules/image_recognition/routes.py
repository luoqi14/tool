#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Blueprint, render_template, request, jsonify, Response, stream_with_context
import os
import tempfile
import json
import httpx
import re
from google.genai import Client

# 创建蓝图
image_recognition_bp = Blueprint('image_recognition', __name__, 
                          url_prefix='/image/recognition',
                          template_folder='templates',
                          static_folder='static')

# 配置Gemini API
API_KEY = "AIzaSyCgxs1UF3qv0d2AFm9Opl1vwroYIlOzW1g"

# 创建Gemini客户端
genai_client = Client(api_key=API_KEY)

# 设置模型
MODEL_NAME = "gemini-2.5-flash-preview-05-20"

@image_recognition_bp.route('/')
def index():
    """图像识别工具主页"""
    return render_template('image_recognition/index.html')

@image_recognition_bp.route('/process', methods=['POST'])
def process_image():
    """处理上传的多张图片并调用Gemini API，使用流式输出"""
    try:
        # 检查是否有文件上传
        if 'files' in request.files:
            # 获取所有上传的文件
            files = request.files.getlist('files')
            prompt = request.form.get('prompt', '请描述这些图片')
            
            if files and len(files) > 0:
                # 存储所有临时文件路径
                temp_file_paths = []
                uploaded_files = []
                
                try:
                    # 处理每个文件
                    for file in files:
                        if file and file.filename:
                            # 创建临时文件
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp:
                                temp_file_path = temp.name
                                file.save(temp_file_path)
                                temp_file_paths.append(temp_file_path)
                                
                                # 使用最新的Gemini API上传文件
                                uploaded_file = genai_client.files.upload(file=temp_file_path)
                                uploaded_files.append(uploaded_file)
                    
                    # 返回流式响应
                    return stream_response(MODEL_NAME, prompt, uploaded_files, temp_file_paths)
                except Exception as e:
                    # 清理临时文件
                    for path in temp_file_paths:
                        if path and os.path.exists(path):
                            os.unlink(path)
                    
                    # 删除上传到Gemini的文件
                    for uploaded_file in uploaded_files:
                        genai_client.files.delete(name=uploaded_file.name)
                    
                    # 重新抛出异常
                    raise e
            else:
                return jsonify({
                    'status': 'error',
                    'message': '未提供有效的文件'
                }), 400
        else:
            return jsonify({
                'status': 'error',
                'message': '未提供图片文件'
            }), 400
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'处理图片时出错: {str(e)}'
        }), 500


@image_recognition_bp.route('/fetch_product_images', methods=['GET', 'POST'])
def fetch_product_images():
    """从外部API获取产品图片URL"""
    try:
        # 获取请求数据 - 支持GET和POST方法
        if request.method == 'POST':
            data = request.json
            if not data or 'productId' not in data:
                return jsonify({
                    'status': 'error',
                    'message': '未提供产品ID'
                }), 400
            product_id = data['productId']
        else:  # GET方法
            product_id = request.args.get('product_id')
            if not product_id:
                return jsonify({
                    'status': 'error',
                    'message': '未提供产品ID'
                }), 400
        
        # 调用外部API
        api_url = 'http://notify.mmm920.com/api/notify_getProductPage'
        payload = {"uniformcodes": [product_id]}
        
        # 使用httpx发送请求，不使用代理
        with httpx.Client() as client:
            response = client.post(api_url, json=payload)
            
            # 检查响应
            if response.status_code != 200:
                return jsonify({
                    'success': False,
                    'message': f'外部API返回错误状态码: {response.status_code}'
                }), 500
            
            # 解析响应数据
            try:
                api_data = response.json()
                
                # 从响应中提取图片URL
                image_urls = []
                base_url = 'https://www.mmm920.com'
                
                # 检查响应数据结构
                if api_data.get('isSuccess') == True and api_data.get('data'):
                    data_array = api_data.get('data')
                    if not isinstance(data_array, list):
                        data_array = [data_array]
                    
                    # 遍历数据提取图片URL
                    for item in data_array:
                        if item and 'page' in item:
                            page_html = item['page']
                            if page_html:
                                # 使用正则表达式提取所有img标签的src属性
                                img_src_pattern = re.compile(r'<img[^>]+src=["\']([^"\'>]+)["\']')
                                img_srcs = img_src_pattern.findall(page_html)
                                
                                for src in img_srcs:
                                    # 如果是相对路径，拼接基础URL
                                    if src.startswith('/'):
                                        src = base_url + src
                                    image_urls.append(src)
                
                # 返回成功响应和图片URL列表
                return jsonify({
                    'success': True,
                    'images': image_urls,
                    'message': f'成功获取{len(image_urls)}张图片'
                })
                
            except Exception as e:
                return jsonify({
                    'success': False,
                    'message': f'解析API响应时出错: {str(e)}'
                }), 500
            
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'处理请求时出错: {str(e)}'
        }), 500

def stream_response(model_name, prompt, uploaded_files, temp_file_paths):
    """流式输出响应，支持多图处理"""
    @stream_with_context
    def generate():
        try:
            # 准备内容列表，包含提示词和所有上传的文件
            contents = [prompt]
            contents.extend(uploaded_files)
            
            # 使用流式生成内容
            response = genai_client.models.generate_content_stream(
                model=model_name,
                contents=contents
            )
            
            # 流式返回数据
            for chunk in response:
                if chunk.text:
                    yield f"data: {json.dumps({'text': chunk.text})}\n\n"
            
            # 发送完成信号
            yield f"data: {json.dumps({'done': True})}\n\n"
        finally:
            # 清理临时文件
            for path in temp_file_paths:
                if path and os.path.exists(path):
                    os.unlink(path)
            
            # 删除上传到Gemini的文件
            for uploaded_file in uploaded_files:
                genai_client.files.delete(name=uploaded_file.name)
    
    return Response(generate(), mimetype='text/event-stream')
