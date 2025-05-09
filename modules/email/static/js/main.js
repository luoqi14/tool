// 全局变量
let attachmentFiles = [];
let quillEditor;
let lastSendResults = [];

// 预设邮件模板
const EMAIL_TEMPLATES = {
    cleaning: {
        name: '洁牙引流',
        subject: '[[文件名]]',
        content: `老师好， 附件是[[年月]]洁牙对账单
1、金曜日开票信息： 详见对账单第三张表
2、开票要求： 发票单价、数量与对账单保持一致
3、重要的！！！结算需要的附件：
诊所开具电子发票： 提供电子版pdf原件+对账单盖章回执pdf版， 邮件反馈；
诊所开具纸质发票： 纸质发票邮寄信息以最新对账单为准~
4、核销问题， 非常重要！ 请及时核销， 不核销影响结算， 后果自负。
ps 账单请随时保存`,
        sender: 'finance01@jarvismedical.com'
    },
    // 可以在这里添加更多模板
};

// 页面加载完成后设置事件监听器
document.addEventListener('DOMContentLoaded', () => {
    // 初始化富文本编辑器
    initRichTextEditor();
    
    // 初始化模板功能
    initTemplateFeature();
    
    // 表单提交
    const form = document.getElementById('emailForm');
    form.addEventListener('submit', handleFormSubmit);
    
    // 附件输入框变化
    const attachmentInput = document.getElementById('attachmentInput');
    attachmentInput.addEventListener('change', handleFileInputChange);
});

// 初始化模板功能
function initTemplateFeature() {
    // 获取所有模板项
    const templateItems = document.querySelectorAll('.template-item');
    
    // 为每个模板项添加点击事件
    templateItems.forEach(item => {
        item.addEventListener('click', function() {
            // 移除所有选中状态
            templateItems.forEach(t => {
                t.setAttribute('data-selected', 'false');
                t.classList.remove('selected');
            });
            
            // 设置当前项为选中状态
            this.setAttribute('data-selected', 'true');
            this.classList.add('selected');
            
            // 获取模板名称
            const templateName = this.getAttribute('data-template');
            
            // 如果选中空白模板，清空内容
            if (!templateName) {
                document.getElementById('subject').value = '';
                quillEditor.root.innerHTML = '';
                document.getElementById('content').value = '';
                document.getElementById('senderEmail').value = '';
                return;
            }
            
            // 获取模板内容
            const template = EMAIL_TEMPLATES[templateName];
            if (template) {
                // 填充主题和内容
                document.getElementById('subject').value = template.subject;
                quillEditor.root.innerHTML = template.content.replaceAll('\n', '<p>');
                document.getElementById('content').value = template.content;
                
                // 如果模板有发送者邮箱，也填充
                if (template.sender) {
                    document.getElementById('senderEmail').value = template.sender;
                }
                
                // 显示成功提示
                showAlert(`已应用「${template.name}」模板`, 'success');
            }
        });
    });
}

// 初始化富文本编辑器
function initRichTextEditor() {
    // 工具栏选项
    const toolbarOptions = [
        ['bold', 'italic', 'underline', 'strike'],        // 文本格式
        ['blockquote', 'code-block'],                     // 引用和代码块
        [{ 'header': 1 }, { 'header': 2 }],               // 标题
        [{ 'list': 'ordered'}, { 'list': 'bullet' }],     // 列表
        [{ 'script': 'sub'}, { 'script': 'super' }],       // 上标/下标
        [{ 'indent': '-1'}, { 'indent': '+1' }],           // 缩进
        [{ 'direction': 'rtl' }],                          // 文本方向
        [{ 'size': ['small', false, 'large', 'huge'] }],   // 字体大小
        [{ 'header': [1, 2, 3, 4, 5, 6, false] }],        // 标题级别
        [{ 'color': [] }, { 'background': [] }],           // 字体颜色和背景色
        [{ 'font': [] }],                                  // 字体系列
        [{ 'align': [] }],                                 // 对齐方式
        ['clean'],                                         // 清除格式
        ['link', 'image']                                  // 链接和图片
    ];
    
    // 初始化Quill编辑器
    quillEditor = new Quill('#editor', {
        modules: {
            toolbar: toolbarOptions
        },
        placeholder: '请输入邮件内容...',
        theme: 'snow'
    });
    
    // 监听内容变化，更新隐藏输入框
    quillEditor.on('text-change', function() {
        document.getElementById('content').value = quillEditor.root.innerHTML;
    });
}

// 处理文件输入框变化
async function handleFileInputChange(event) {
    const fileInput = event.target;
    const fileTypeWarning = document.getElementById('fileTypeWarning');
    
    // 清除警告
    fileTypeWarning.classList.add('d-none');
    
    if (fileInput.files.length > 0) {
        // 检查文件类型
        let allValid = true;
        for (const file of fileInput.files) {
            const fileExt = file.name.split('.').pop().toLowerCase();
            if (fileExt !== 'xlsx' && fileExt !== 'xls') {
                allValid = false;
                break;
            }
        }
        
        if (!allValid) {
            fileTypeWarning.classList.remove('d-none');
            fileInput.value = '';
            return;
        }
        
        // 显示加载中提示
        showLoading(true, '正在解析Excel文件...');
        
        try {
            // 处理每个文件
            for (const file of fileInput.files) {
                try {
                    // 读取文件内容并转换为Base64
                    const base64Data = await readFileAsBase64(file);
                    
                    // 发送到后端解析
                    const response = await fetch('/email/api/parse-excel', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            file_name: file.name,
                            file_data: base64Data
                        })
                    });
                    
                    const result = await response.json();
                    
                    if (result.success) {
                        // 添加到附件列表
                        attachmentFiles.push({
                            file: file,
                            name: file.name,
                            data: base64Data,
                            email: result.email,
                            period: result.period || '', // 添加对账期间
                            size: file.size
                        });
                    } else {
                        // 显示错误
                        console.error(`解析文件 ${file.name} 失败:`, result.message);
                        attachmentFiles.push({
                            file: file,
                            name: file.name,
                            data: base64Data,
                            error: result.message,
                            size: file.size
                        });
                    }
                } catch (error) {
                    console.error(`处理文件 ${file.name} 时出错:`, error);
                    attachmentFiles.push({
                        file: file,
                        name: file.name,
                        error: error.message,
                        size: file.size
                    });
                }
            }
            
            // 更新附件列表显示
            updateAttachmentList();
        } catch (error) {
            console.error('解析文件时出错:', error);
            showAlert(`解析文件时出错: ${error.message}`, 'danger');
        } finally {
            showLoading(false);
        }
    }
}
// 处理表单提交
async function handleFormSubmit(event) {
    event.preventDefault();
    
    // 显示加载中提示
    showLoading(true, '正在发送邮件，请稍候...');
    
    try {
        // 获取表单数据
        const subject = document.getElementById('subject').value;
        const content = document.getElementById('content').value;
        const senderEmail = document.getElementById('senderEmail').value;
        const senderPassword = document.getElementById('senderPassword').value;
        
        // 验证发送邮箱和密码
        if (!senderEmail || !senderPassword) {
            showAlert('请填写发送邮箱和密码', 'warning');
            showLoading(false);
            return;
        }
        
        // 验证表单
        if (!subject || !content) {
            showAlert('请填写主题和内容', 'warning');
            showLoading(false);
            return;
        }
        
        // 验证附件
        if (attachmentFiles.length === 0) {
            showAlert('请上传至少一个Excel文件', 'warning');
            showLoading(false);
            return;
        }
        
        // 检查是否有有效的附件
        const validAttachments = attachmentFiles.filter(att => att.email && !att.error);
        if (validAttachments.length === 0) {
            showAlert('没有可用的附件，请确保文件中包含有效的邮箱地址', 'warning');
            showLoading(false);
            return;
        }
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            attachments: validAttachments.map(att => ({
                name: att.name,
                data: att.data,
                email: att.email,
                period: att.period || ''
            }))
        };
        
        // 发送请求
        const response = await fetch('/email/api/send-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        // 处理响应
        const result = await response.json();
        
        // 保存结果用于可能的重发
        lastSendResults = result.results || [];
        
        // 显示结果
        if (result.success) {
            showAlert(result.message, 'success');
            showDetailedResults(result.results);
        } else {
            showAlert(result.message, 'danger');
        }
    } catch (error) {
        console.error('发送邮件时出错:', error);
        showAlert(`发送邮件时出错: ${error.message}`, 'danger');
    } finally {
        showLoading(false);
    }
}

// 显示详细结果
function showDetailedResults(results) {
    // 保存结果以便重发
    lastSendResults = results;
    
    const detailedResults = document.getElementById('detailedResults');
    const resultsList = document.getElementById('resultsList');
    const resendFailedBtn = document.getElementById('resendFailedBtn');
    
    // 清空当前结果
    resultsList.innerHTML = '';
    
    // 检查是否有失败的邮件
    const hasFailedEmails = results.some(result => !result.success);
    
    // 显示/隐藏重发按钮
    if (hasFailedEmails) {
        resendFailedBtn.classList.remove('d-none');
    } else {
        resendFailedBtn.classList.add('d-none');
    }
    
    // 添加每个结果项
    results.forEach((result, index) => {
        const row = document.createElement('tr');
        
        // 设置行的背景色
        if (!result.success) {
            row.classList.add('table-danger');
        } else {
            row.classList.add('table-success');
        }
        
        // 文件名
        const fileCell = document.createElement('td');
        fileCell.textContent = result.filename;
        row.appendChild(fileCell);
        
        // 收件人
        const recipientCell = document.createElement('td');
        recipientCell.textContent = result.recipient || '-';
        row.appendChild(recipientCell);
        
        // 状态
        const statusCell = document.createElement('td');
        statusCell.innerHTML = result.success ? 
            '<span class="badge bg-success">成功</span>' : 
            '<span class="badge bg-danger">失败</span>';
        row.appendChild(statusCell);
        
        // 消息
        const messageCell = document.createElement('td');
        messageCell.textContent = result.message;
        row.appendChild(messageCell);
        
        // 操作
        const actionCell = document.createElement('td');
        if (!result.success) {
            const resendBtn = document.createElement('button');
            resendBtn.className = 'btn btn-sm btn-outline-primary';
            resendBtn.innerHTML = '<i class="bi bi-arrow-repeat"></i> 重发';
            resendBtn.onclick = () => resendEmail(index);
            actionCell.appendChild(resendBtn);
        } else {
            actionCell.textContent = '-';
        }
        row.appendChild(actionCell);
        
        resultsList.appendChild(row);
    });
    
    // 显示结果区域
    detailedResults.classList.remove('d-none');
    
    // 添加重发按钮的点击事件
    resendFailedBtn.onclick = resendAllFailedEmails;
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

// 更新附件列表显示
function updateAttachmentList() {
    const attachmentList = document.getElementById('attachmentList');
    attachmentList.innerHTML = '';
    
    attachmentFiles.forEach((attachment, index) => {
        // 创建附件项
        const attachmentItem = document.createElement('div');
        attachmentItem.className = 'attachment-item';
        
        // 附件信息区域
        const infoDiv = document.createElement('div');
        infoDiv.className = 'attachment-info';
        
        // 文件名称和大小
        const nameSpan = document.createElement('div');
        nameSpan.textContent = `${attachment.name} (${formatFileSize(attachment.size)})`;
        infoDiv.appendChild(nameSpan);
        
        // 邮箱或错误信息
        if (attachment.email) {
            const emailSpan = document.createElement('div');
            emailSpan.className = 'attachment-email';
            emailSpan.textContent = `收件人: ${attachment.email}`;
            infoDiv.appendChild(emailSpan);
        } else if (attachment.error) {
            const errorSpan = document.createElement('div');
            errorSpan.className = 'text-danger';
            errorSpan.textContent = `错误: ${attachment.error}`;
            infoDiv.appendChild(errorSpan);
        }
        
        // 按钮容器
        const btnContainer = document.createElement('div');
        btnContainer.className = 'attachment-buttons';
        
        // 预览按钮 - 只有当有邮箱地址时才显示
        if (attachment.email) {
            const previewBtn = document.createElement('button');
            previewBtn.className = 'btn btn-sm btn-outline-primary me-2';
            previewBtn.innerHTML = '<i class="bi bi-eye"></i> 预览';
            previewBtn.type = 'button'; // 明确设置为按钮类型，防止触发表单提交
            previewBtn.onclick = (e) => {
                e.preventDefault(); // 防止事件冒泡和默认行为
                e.stopPropagation(); // 阻止事件冒泡
                previewEmail(index);
            };
            btnContainer.appendChild(previewBtn);
        }
        
        // 删除按钮
        const removeBtn = document.createElement('button');
        removeBtn.className = 'btn btn-sm btn-outline-danger';
        removeBtn.innerHTML = '<i class="bi bi-trash"></i> 删除';
        removeBtn.onclick = () => removeAttachment(index);
        btnContainer.appendChild(removeBtn);
        
        // 添加到附件项
        attachmentItem.appendChild(infoDiv);
        attachmentItem.appendChild(btnContainer);
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
    
    // 清空详细结果
    const detailedResults = document.getElementById('detailedResults');
    detailedResults.classList.add('d-none');
    
    // 重置发送结果
    lastSendResults = [];
}

// 重发单个失败邮件
async function resendEmail(index) {
    // 获取失败的邮件信息
    const failedResult = lastSendResults[index];
    if (!failedResult || failedResult.success) return;
    
    // 查找对应的附件
    const attachment = attachmentFiles.find(att => att.name === failedResult.filename);
    if (!attachment) {
        showAlert(`找不到附件: ${failedResult.filename}`, 'warning');
        return;
    }
    
    // 显示加载中提示
    showLoading(true, '正在重新发送邮件...');
    
    try {
        // 获取表单数据
        const subject = document.getElementById('subject').value;
        const content = document.getElementById('content').value;
        
        // 获取发送邮箱和密码
        const senderEmail = document.getElementById('senderEmail').value;
        const senderPassword = document.getElementById('senderPassword').value;
        
        // 验证发送邮箱和密码
        if (!senderEmail || !senderPassword) {
            showAlert('请填写发送邮箱和密码', 'warning');
            showLoading(false);
            return;
        }
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            attachments: [{
                name: attachment.name,
                data: attachment.data,
                email: attachment.email,
                period: attachment.period || ''
            }]
        };
        
        // 发送请求
        const response = await fetch('/email/api/send-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        // 处理响应
        const result = await response.json();
        
        // 更新结果
        if (result.results && result.results.length > 0) {
            // 替换原来的结果
            lastSendResults[index] = result.results[0];
            // 重新显示结果
            showDetailedResults(lastSendResults);
        }
        
        // 显示结果消息
        showAlert(result.message, result.success ? 'success' : 'danger');
    } catch (error) {
        console.error('重发邮件时出错:', error);
        showAlert(`重发邮件时出错: ${error.message}`, 'danger');
    } finally {
        showLoading(false);
    }
}

// 重发所有失败邮件
async function resendAllFailedEmails() {
    // 过滤出失败的邮件索引
    const failedIndices = lastSendResults
        .map((result, index) => result.success ? -1 : index)
        .filter(index => index !== -1);
    
    if (failedIndices.length === 0) {
        showAlert('没有需要重发的邮件', 'warning');
        return;
    }
    
    // 显示加载中提示
    showLoading(true, `正在重新发送 ${failedIndices.length} 封失败邮件...`);
    
    try {
        // 获取表单数据
        const subject = document.getElementById('subject').value;
        const content = document.getElementById('content').value;
        const senderEmail = document.getElementById('senderEmail').value;
        const senderPassword = document.getElementById('senderPassword').value;
        
        // 验证发送邮箱和密码
        if (!senderEmail || !senderPassword) {
            showAlert('请填写发送邮箱和密码', 'warning');
            showLoading(false);
            return;
        }
        
        // 准备要重发的附件
        const failedAttachments = [];
        
        for (const index of failedIndices) {
            const failedResult = lastSendResults[index];
            const attachment = attachmentFiles.find(att => att.name === failedResult.filename);
            
            if (attachment && attachment.email) {
                failedAttachments.push({
                    name: attachment.name,
                    data: attachment.data,
                    email: attachment.email,
                    period: attachment.period || ''
                });
            }
        }
        
        if (failedAttachments.length === 0) {
            showAlert('找不到可重发的附件', 'warning');
            showLoading(false);
            return;
        }
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            attachments: failedAttachments
        };
        
        // 发送请求
        const response = await fetch('/email/api/send-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        // 处理响应
        const result = await response.json();
        
        // 更新结果
        if (result.results && result.results.length > 0) {
            // 更新每个重发的邮件结果
            result.results.forEach(newResult => {
                const index = lastSendResults.findIndex(r => r.filename === newResult.filename);
                if (index !== -1) {
                    lastSendResults[index] = newResult;
                }
            });
            
            // 重新显示结果
            showDetailedResults(lastSendResults);
        }
        
        // 显示结果消息
        showAlert(result.message, result.success ? 'success' : 'danger');
    } catch (error) {
        console.error('重发邮件时出错:', error);
        showAlert(`重发邮件时出错: ${error.message}`, 'danger');
    } finally {
        showLoading(false);
    }
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
}

// 预览邮件内容
function previewEmail(index) {
    const attachment = attachmentFiles[index];
    if (!attachment || !attachment.email) return;
    
    // 获取当前邮件主题和内容
    const subject = document.getElementById('subject').value;
    const content = document.getElementById('content').value;
    
    // 提取文件名（不包含扩展名）用于占位符替换
    const filename = attachment.name;
    const filename_without_ext = filename.split('.').slice(0, -1).join('.');
    
    // 替换占位符
    let previewSubject = subject.replace(/\[\[文件名\]\]/g, filename_without_ext);
    let previewContent = content.replace(/\[\[文件名\]\]/g, filename_without_ext);
    
    // 如果有对账期间，也替换[[年月]]占位符
    if (attachment.period) {
        previewSubject = previewSubject.replace(/\[\[年月\]\]/g, attachment.period);
        previewContent = previewContent.replace(/\[\[年月\]\]/g, attachment.period);
    }
    
    // 更新预览模态框内容
    document.getElementById('previewRecipient').textContent = attachment.email;
    document.getElementById('previewSubject').textContent = previewSubject;
    document.getElementById('previewContent').innerHTML = previewContent;
    document.getElementById('previewAttachmentName').textContent = filename;
    
    // 显示预览模态框
    const previewModal = new bootstrap.Modal(document.getElementById('emailPreviewModal'));
    previewModal.show();
}

// 显示/隐藏加载中提示
function showLoading(show, message = '正在发送邮件，请稍候...') {
    const loadingOverlay = document.getElementById('loadingOverlay');
    const loadingMessage = document.getElementById('loadingMessage');
    
    if (show) {
        loadingMessage.textContent = message;
        loadingOverlay.style.visibility = 'visible';
    } else {
        loadingOverlay.style.visibility = 'hidden';
    }
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
