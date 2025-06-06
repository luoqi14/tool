#!/usr/bin/env python
# -*- coding: utf-8 -*-

from flask import Blueprint, render_template, request, jsonify, Response, stream_with_context
import os
import tempfile
import json
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

# 移除了标注相关函数

@image_recognition_bp.route('/')
def index():
    """图像识别工具主页"""
    return render_template('image_recognition/index.html')

@image_recognition_bp.route('/process', methods=['POST'])
def process_image():
    """处理上传的图片并调用Gemini API，使用流式输出"""
    try:
        # 检查是否有文件上传
        if 'file' in request.files:
            # 直接处理上传的文件
            file = request.files['file']
            prompt = request.form.get('prompt', '请描述这张图片')
            
            if file and file.filename:
                # 创建临时文件
                temp_file_path = None
                with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp:
                    temp_file_path = temp.name
                    file.save(temp_file_path)
                
                try:
                    # 使用最新的Gemini API上传文件方式
                    uploaded_file = genai_client.files.upload(file=temp_file_path)
                    
                    # 返回流式响应
                    return stream_response(MODEL_NAME, prompt, uploaded_file, temp_file_path)
                except Exception as e:
                    # 清理临时文件
                    if temp_file_path and os.path.exists(temp_file_path):
                        os.unlink(temp_file_path)
                    
                    # 删除上传到Gemini的文件
                    if 'uploaded_file' in locals():
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


def stream_response(model_name, prompt, uploaded_file, temp_file_path):
    """流式输出响应"""
    @stream_with_context
    def generate():
        try:
            # 使用流式生成内容
            response = genai_client.models.generate_content_stream(
                model=model_name,
                contents=[prompt, uploaded_file]
            )
            
            # 流式返回数据
            for chunk in response:
                if chunk.text:
                    yield f"data: {json.dumps({'text': chunk.text})}\n\n"
            
            # 发送完成信号
            yield f"data: {json.dumps({'done': True})}\n\n"
        finally:
            # 清理临时文件
            if temp_file_path and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
            
            # 删除上传到Gemini的文件
            genai_client.files.delete(name=uploaded_file.name)
    
    return Response(generate(), mimetype='text/event-stream')
