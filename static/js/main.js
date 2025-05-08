// 全局变量
let attachmentFiles = [];

// 页面加载完成后设置事件监听器
document.addEventListener('DOMContentLoaded', () => {
    // 表单提交
    const form = document.getElementById('emailForm');
    form.addEventListener('submit', handleFormSubmit);
    
    // 添加附件按钮
    const addAttachmentBtn = document.getElementById('addAttachmentBtn');
    addAttachmentBtn.addEventListener('click', addAttachment);
});

// 处理表单提交
async function handleFormSubmit(event) {
    event.preventDefault();
    
    // 显示加载中提示
    showLoading(true);
    
    try {
        // 获取表单数据
        const recipients = document.getElementById('recipients').value;
        const subject = document.getElementById('subject').value;
        const content = document.getElementById('content').value;
        const contentType = document.getElementById('contentType').value;
        
        // 验证表单
        if (!recipients || !subject || !content) {
            showAlert('请填写所有必填字段', 'warning');
            showLoading(false);
            return;
        }
        
        // 处理附件数据
        const attachments = [];
        for (const file of attachmentFiles) {
            try {
                // 读取文件内容并转换为Base64
                const base64Data = await readFileAsBase64(file);
                
                attachments.push({
                    name: file.name,
                    data: base64Data,
                    type: file.type
                });
            } catch (error) {
                console.error(`处理附件 ${file.name} 时出错:`, error);
                showAlert(`处理附件 ${file.name} 时出错: ${error.message}`, 'danger');
                showLoading(false);
                return;
            }
        }
        
        // 准备请求数据
        const requestData = {
            recipients,
            subject,
            content,
            contentType,
            attachments
        };
        
        // 发送请求
        const response = await fetch('/api/send-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        // 处理响应
        const result = await response.json();
        
        if (result.success) {
            showAlert(result.message, 'success');
            document.getElementById('emailForm').reset();
            clearAttachments();
        } else {
            showAlert(`邮件发送失败: ${result.message}`, 'danger');
        }
    } catch (error) {
        console.error('发送邮件时出错:', error);
        showAlert(`发送邮件时出错: ${error.message}`, 'danger');
    } finally {
        showLoading(false);
    }
}

// 读取文件并转换为Base64
function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = () => {
            resolve(reader.result);
        };
        
        reader.onerror = (error) => {
            reject(error);
        };
        
        reader.readAsDataURL(file);
    });
}

// 添加附件
function addAttachment() {
    const fileInput = document.getElementById('attachmentInput');
    
    if (fileInput.files.length > 0) {
        // 将FileList转换为数组并添加到附件列表
        Array.from(fileInput.files).forEach(file => {
            attachmentFiles.push(file);
        });
        
        // 更新附件列表显示
        updateAttachmentList();
        
        // 清空文件输入，以便可以再次选择相同的文件
        fileInput.value = '';
    }
}

// 更新附件列表显示
function updateAttachmentList() {
    const attachmentList = document.getElementById('attachmentList');
    attachmentList.innerHTML = '';
    
    attachmentFiles.forEach((file, index) => {
        // 创建附件项
        const attachmentItem = document.createElement('div');
        attachmentItem.className = 'attachment-item';
        
        // 文件名称和大小
        const nameSpan = document.createElement('span');
        nameSpan.textContent = `${file.name} (${formatFileSize(file.size)})`;
        
        // 删除按钮
        const removeBtn = document.createElement('button');
        removeBtn.className = 'btn btn-sm btn-outline-danger';
        removeBtn.textContent = '删除';
        removeBtn.onclick = () => removeAttachment(index);
        
        // 添加到附件项
        attachmentItem.appendChild(nameSpan);
        attachmentItem.appendChild(removeBtn);
        attachmentList.appendChild(attachmentItem);
    });
}

// 删除附件
function removeAttachment(index) {
    attachmentFiles.splice(index, 1);
    updateAttachmentList();
}

// 清空附件列表
function clearAttachments() {
    attachmentFiles = [];
    updateAttachmentList();
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
}

// 显示/隐藏加载中提示
function showLoading(show) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.style.visibility = show ? 'visible' : 'hidden';
}

// 显示提示信息
function showAlert(message, type) {
    const resultArea = document.getElementById('resultArea');
    resultArea.textContent = message;
    resultArea.className = `alert alert-${type}`;
    resultArea.classList.remove('d-none');
    
    // 5秒后自动隐藏成功消息
    if (type === 'success') {
        setTimeout(() => {
            resultArea.classList.add('d-none');
        }, 5000);
    }
}
