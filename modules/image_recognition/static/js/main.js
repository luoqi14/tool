document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const dropArea = document.getElementById('dropArea');
    const imageInput = document.getElementById('imageInput');
    const uploadPlaceholder = document.getElementById('uploadPlaceholder');
    const imagePreview = document.getElementById('imagePreview');
    const previewImg = document.getElementById('previewImg');
    const removeImage = document.getElementById('removeImage');
    const processButton = document.getElementById('processButton');
    const promptInput = document.getElementById('promptInput');
    const resultContainer = document.getElementById('resultContainer');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const statusIndicator = document.getElementById('status-indicator');
    
    // 设置默认提示词
    promptInput.value = `识别图中文字，生成三份内容：原样输出：再json结构化输出；最后按我给的字段json结构化输出：
    {
        "productName": null, // 商品名称
        "brandName": null, // 品牌名称
        "orderingInfo": null, // 规格信息
        "validPeriod": null, // 有效期
        "sterilizationValidity": null, // 灭菌有效期
        "guaranteePeriod": null, // 质保期
        "storage": null, // 储存方式
        "useLife": null, // 设备使用寿命
        "effectiveComponent": null, // 有效成分及含量
        "instruction": null, // 使用说明
        "clinicalOperation": null, // 临床操作
        "mountingsDesc": null, // 配件说明
        "maintaining": null, // 维护/保养
        "productDesc": null, // 商品介绍
        "sterilization": null, // 消毒灭菌
    }
    如果识别到值就填充，否则为null`;
    
    // 点击上传区域触发文件选择
    dropArea.addEventListener('click', function() {
        if (!imageInput.files[0]) {
            imageInput.click();
        }
    });
    
    // 处理文件选择
    imageInput.addEventListener('change', function() {
        handleFiles(this.files);
    });
    
    // 处理拖放 - 拖动进入
    dropArea.addEventListener('dragenter', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.add('drag-over');
    });
    
    // 处理拖放 - 拖动经过
    dropArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.add('drag-over');
    });
    
    // 处理拖放 - 拖动离开
    dropArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.remove('drag-over');
    });
    
    // 处理拖放 - 放置文件
    dropArea.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.remove('drag-over');
        
        const dt = e.dataTransfer;
        const files = dt.files;
        
        handleFiles(files);
    });
    
    // 处理文件
    function handleFiles(files) {
        if (files.length === 0) return;
        
        const file = files[0];
        
        // 检查是否为图片
        if (!file.type.match('image.*')) {
            alert('请上传图片文件');
            return;
        }
        
        // 直接创建URL显示预览
        const objectUrl = URL.createObjectURL(file);
        previewImg.src = objectUrl;
        uploadPlaceholder.style.display = 'none';
        imagePreview.style.display = 'block';
        processButton.disabled = false;
    }
    
    // 移除图片
    removeImage.addEventListener('click', function(e) {
        e.stopPropagation();
        previewImg.src = '';
        uploadPlaceholder.style.display = 'block';
        imagePreview.style.display = 'none';
        processButton.disabled = true;
        imageInput.value = '';
        resultContainer.innerHTML = '';
        
        // 释放之前创建的URL对象
        if (previewImg.src) {
            URL.revokeObjectURL(previewImg.src);
        }
    });
    
    // 处理图片按钮点击
    processButton.addEventListener('click', function() {
        if (!imageInput.files[0]) {
            alert('请先上传图片');
            return;
        }
        
        // 显示加载指示器
        loadingIndicator.classList.remove('d-none');
        processButton.disabled = true;
        
        // 更新状态指示器
        updateStatus('processing');
        
        // 清空之前的结果
        resultContainer.innerHTML = `
            <div class="alert alert-info streaming-text">
                <div id="streaming-result" class="result-text"></div>
                <div class="typing-indicator">
                    <span></span>
                    <span></span>
                    <span></span>
                </div>
            </div>
        `;
        
        const streamingResult = document.getElementById('streaming-result');
        
        // 获取提示词
        const prompt = promptInput.value || '请描述这张图片';
        
        // 使用FormData直接发送文件
        const formData = new FormData();
        formData.append('file', imageInput.files[0]);
        formData.append('prompt', prompt);
        
        // 使用fetch API的流式处理
        
        // 使用FormData直接发送文件
        fetch('/image/recognition/process', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }
            
            // 检查响应类型
            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('text/event-stream')) {
                // 处理流式响应
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                
                // 递归函数来处理数据流
                function readStream() {
                    return reader.read().then(({ done, value }) => {
                        if (done) {
                            // 流结束
                            loadingIndicator.classList.add('d-none');
                            processButton.disabled = false;
                            return;
                        }
                        
                        // 解码并处理数据
                        const chunk = decoder.decode(value, { stream: true });
                        
                        // 处理SSE格式的数据
                        const lines = chunk.split('\n\n');
                        for (const line of lines) {
                            if (line.startsWith('data: ')) {
                                try {
                                    const data = JSON.parse(line.substring(6));
                                    
                                    if (data.done) {
                                        // 流结束
                                        loadingIndicator.classList.add('d-none');
                                        processButton.disabled = false;
                                        // 移除打字指示器
                                        document.querySelector('.typing-indicator').remove();
                                        // 更新状态
                                        updateStatus('completed');
                                    } else if (data.text) {
                                        // 添加文本
                                        streamingResult.innerHTML += formatResult(data.text);
                                        // 滚动到底部
                                        resultContainer.scrollTop = resultContainer.scrollHeight;
                                    }
                                } catch (e) {
                                    console.error('解析SSE数据出错:', e);
                                }
                            }
                        }
                        
                        // 继续读取
                        return readStream();
                    });
                }
                
                // 开始处理流
                return readStream();
            } else {
                // 非流式响应，回退到普通JSON处理
                return response.json().then(handleResponse);
            }
        })
        .catch(handleError);
    });
    
    // 处理API响应
    function handleResponse(data) {
        // 隐藏加载指示器
        loadingIndicator.classList.add('d-none');
        processButton.disabled = false;
        
        if (data.status === 'success') {
            // 显示结果
            resultContainer.innerHTML = `
                <div class="alert alert-success">
                    <h5 class="alert-heading">识别成功</h5>
                    <hr>
                    <div class="result-text">${formatResult(data.result)}</div>
                </div>
            `;
        } else {
            // 显示错误
            resultContainer.innerHTML = `
                <div class="alert alert-danger">
                    <h5 class="alert-heading">处理失败</h5>
                    <hr>
                    <p>${data.message || '未知错误'}</p>
                </div>
            `;
        }
    }
    
    // 处理错误
    function handleError(error) {
        // 隐藏加载指示器
        loadingIndicator.classList.add('d-none');
        processButton.disabled = false;
        
        // 更新状态
        updateStatus('error');
        
        // 显示错误
        resultContainer.innerHTML = `
            <div class="alert alert-danger">
                <h5 class="alert-heading">请求失败</h5>
                <hr>
                <p>无法连接到服务器，请稍后再试</p>
                <p class="text-muted small">错误详情: ${error.message}</p>
            </div>
        `;
    }
    
    // 格式化结果文本，保留换行
    function formatResult(text) {
        if (!text) return '';
        return text
            .replace(/\n/g, '<br>')
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>');
    }
    
    // 更新状态指示器
    function updateStatus(status) {
        // 移除所有状态类
        const statusBadge = statusIndicator.querySelector('.badge');
        statusBadge.classList.remove('bg-secondary', 'bg-primary', 'processing', 'completed', 'error');
        
        // 根据状态设置样式和文本
        switch(status) {
            case 'idle':
                statusBadge.classList.add('bg-secondary');
                statusBadge.textContent = '尚未开始';
                break;
            case 'processing':
                statusBadge.classList.add('bg-primary', 'processing');
                statusBadge.textContent = '正在处理';
                break;
            case 'completed':
                statusBadge.classList.add('bg-success', 'completed');
                statusBadge.textContent = '处理完成';
                break;
            case 'error':
                statusBadge.classList.add('bg-danger', 'error');
                statusBadge.textContent = '处理错误';
                break;
        }
    }
});
