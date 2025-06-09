// 全局变量
let attachmentFiles = [];
let quillEditor;
let lastSendResults = [];

// 预设邮件模板
const EMAIL_TEMPLATES = {
    cleaning: {
        name: '洁牙引流对账单',
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
    offline_supplier: {
        name: '线下供应商对账单',
        subject: '佳沃思平台服务费',
        content: `尊敬的[[公司]]：
您好！感谢贵司长期以来的支持与信任。现将[[年月日范围]]服务费发送至您处，请协助完成以下流程：
1、附件中的Excel 服务对账明细辛苦确认，确认后无误后，请邮件反馈技术服务费收取通知函盖章扫描件，我们会在收到反馈无误邮件5个工作日内开具相应的发票
2、以上核对无误后按发票金额及时回款，完成款项支付，避免影响后续服务

再次感谢您的配合！期待与贵司保持高效协作。
顺祝商祥`,
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
    const wordExcelWarning = document.getElementById('wordExcelWarning');
    
    // 清除警告
    fileTypeWarning.classList.add('d-none');
    wordExcelWarning.classList.add('d-none');
    
    if (fileInput.files.length > 0) {
        // 检查文件类型
        let allValid = true;
        for (const file of fileInput.files) {
            const fileExt = file.name.split('.').pop().toLowerCase();
            if (fileExt !== 'xlsx' && fileExt !== 'xls' && fileExt !== 'doc' && fileExt !== 'docx') {
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
        showLoading(true, '正在解析文件...');
        
        try {
            // 分类文件
            const excelFiles = [];
            const wordFiles = [];
            
            for (const file of fileInput.files) {
                const fileExt = file.name.split('.').pop().toLowerCase();
                if (fileExt === 'xlsx' || fileExt === 'xls') {
                    excelFiles.push(file);
                } else if (fileExt === 'doc' || fileExt === 'docx') {
                    wordFiles.push(file);
                }
            }
            
            // 先处理所有文件
            for (const file of fileInput.files) {
                try {
                    // 读取文件内容并转换为Base64
                    const base64Data = await readFileAsBase64(file);
                    const fileExt = file.name.split('.').pop().toLowerCase();
                    const isWord = (fileExt === 'doc' || fileExt === 'docx');
                    
                    // 发送到后端解析
                    const response = await fetch('/email/api/parse-file', {
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
                            period: result.period || '', 
                            period_range: result.period_range || '', 
                            company: result.company || '', 
                            is_word: isWord,
                            size: file.size,
                            filename_without_ext: result.filename_without_ext || ''
                        });
                    } else {
                        // 显示错误
                        console.error(`解析文件 ${file.name} 失败:`, result.message);
                        attachmentFiles.push({
                            file: file,
                            name: file.name,
                            data: base64Data,
                            error: result.message,
                            is_word: isWord,
                            size: file.size,
                            filename_without_ext: result.filename_without_ext || ''
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
            
            // 每次上传文件后重新检查所有文件的匹配关系
            matchFilesAndUpdateUI();
            
            // 更新附件列表显示
            updateAttachmentList();
            
            // 如果没有附件，确保发送按钮是启用的
            if (attachmentFiles.length === 0) {
                document.querySelector('.btn-send').disabled = false;
            }
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
        const validAttachments = attachmentFiles.filter(att => att.email && !att.error && !att.is_word);
        
        // 如果有Word文档，需要检查是否所有Excel文件都有匹配的Word文档
        const hasWordFiles = attachmentFiles.some(att => att.is_word && !att.error);
        if (hasWordFiles) {
            // 检查是否有未匹配的Excel文件
            const unmatchedExcel = validAttachments.some(att => !att.word_attachment);
            
            if (unmatchedExcel) {
                showAlert('存在未匹配的Excel文件，请确保每个Excel文件都有对应的Word文档', 'warning');
                showLoading(false);
                return;
            }
        }
        if (validAttachments.length === 0) {
            showAlert('没有可用的附件，请确保文件中包含有效的邮箱地址', 'warning');
            showLoading(false);
            return;
        }
        
        // 获取是否将邮件副本发送给自己的选项
        const ccSender = document.getElementById('ccSender').checked;
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            ccSender, // 添加是否将邮件副本发送给自己的选项
            attachments: validAttachments.map(att => {
                // 基本附件信息
                const attachmentData = {
                    name: att.name,
                    data: att.data,
                    email: att.email,
                    period: att.period || '',
                    period_range: att.period_range || '',
                    company: att.company || '',
                    is_word: att.is_word || false
                };
                
                // 如果是Excel文件且有匹配的Word文档，添加Word附件信息
                if (!att.is_word && att.word_attachment) {
                    attachmentData.word_attachment = {
                        name: att.word_attachment.name,
                        data: att.word_attachment.data
                    };
                }
                
                return attachmentData;
            })
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
    
    if (attachmentFiles.length === 0) {
        attachmentList.innerHTML = '<div class="text-center text-muted py-3">暂无上传文件</div>';
        return;
    }
    
    // 清空列表
    attachmentList.innerHTML = '';
    
    // 创建已处理文件的集合，用于跟踪已显示的文件
    const processedFiles = new Set();
    
    // 首先处理所有Excel文件及其匹配的Word文档
    attachmentFiles.forEach((attachment, index) => {
        // 如果已经处理过该文件或者是Word文档，则跳过
        if (processedFiles.has(index) || attachment.is_word) {
            return;
        }
        
        // 不再尝试基于邮箱匹配，只使用文件名匹配逻辑
        // 匹配逻辑已经在matchFilesAndUpdateUI函数中实现
        
        const fileItem = document.createElement('div');
        fileItem.className = 'attachment-item';
        
        // 根据是否有错误设置不同的样式
        if (attachment.error) {
            fileItem.classList.add('attachment-error');
        } else if (attachment.warning) {
            fileItem.classList.add('attachment-warning');
        }
        
        // 文件信息区域
        const fileInfo = document.createElement('div');
        fileInfo.className = 'attachment-info';
        
        // 检查是否有匹配的Word文档
        const wordAttachment = attachment.word_attachment;
        const hasWordMatch = wordAttachment && !wordAttachment.error;
        
        // 如果有匹配的Word文档，创建一个卡片式布局
        if (hasWordMatch) {
            // 标记该Word文档已处理
            const wordIndex = attachmentFiles.findIndex(att => att === wordAttachment);
            if (wordIndex !== -1) {
                processedFiles.add(wordIndex);
            }
            
            // 创建卡片式布局
            fileItem.classList.add('matched-files-card');
            
            // 邮箱信息
            if (attachment.email) {
                const emailInfo = document.createElement('div');
                emailInfo.className = 'matched-email-info';
                emailInfo.innerHTML = `<span class="text-success"><i class="bi bi-envelope"></i> ${attachment.email}</span>`;
                fileInfo.appendChild(emailInfo);
            }
            
            // Excel文件信息
            const excelFileInfo = document.createElement('div');
            excelFileInfo.className = 'matched-file-info';
            excelFileInfo.innerHTML = `
                <div class="file-icon"><i class="bi bi-file-earmark-excel"></i></div>
                <div class="file-details">
                    <div class="file-name">${attachment.name}</div>
                    <div class="file-size text-muted">${formatFileSize(attachment.size)}</div>
                </div>
            `;
            
            // Word文档信息
            const wordFileInfo = document.createElement('div');
            wordFileInfo.className = 'matched-file-info';
            wordFileInfo.innerHTML = `
                <div class="file-icon"><i class="bi bi-file-earmark-word"></i></div>
                <div class="file-details">
                    <div class="file-name">${wordAttachment.name}</div>
                    <div class="file-size text-muted">${formatFileSize(wordAttachment.size)}</div>
                </div>
            `;
            
            // 添加到文件信息区
            const filesContainer = document.createElement('div');
            filesContainer.className = 'matched-files-container';
            filesContainer.appendChild(excelFileInfo);
            filesContainer.appendChild(wordFileInfo);
            fileInfo.appendChild(filesContainer);
            
            // 操作按钮
            const fileActions = document.createElement('div');
            fileActions.className = 'attachment-actions';
            
            // 预览按钮
            if (attachment.email && !attachment.error) {
                const previewButton = document.createElement('button');
                previewButton.className = 'btn btn-sm btn-outline-primary me-2';
                previewButton.innerHTML = '<i class="bi bi-eye"></i>';
                previewButton.title = '预览邮件';
                previewButton.onclick = (event) => {
                    event.preventDefault(); // 防止页面滚动到顶部
                    previewEmail(index);
                    return false;
                };
                fileActions.appendChild(previewButton);
            }
            
            // 删除按钮
            const deleteButton = document.createElement('button');
            deleteButton.className = 'btn btn-sm btn-outline-danger';
            deleteButton.innerHTML = '<i class="bi bi-trash"></i>';
            deleteButton.title = '删除附件';
            deleteButton.onclick = () => removeAttachment(index);
            fileActions.appendChild(deleteButton);
            
            // 组装文件项
            fileItem.appendChild(fileInfo);
            fileItem.appendChild(fileActions);
            
            // 添加到列表
            attachmentList.appendChild(fileItem);
        } else {
            // 如果没有匹配的Word文档，使用原来的单文件布局
            
            // 文件图标
            const fileIcon = document.createElement('div');
            fileIcon.className = 'attachment-icon';
            fileIcon.innerHTML = '<i class="bi bi-file-earmark-excel"></i>';
            
            // 文件名称
            const fileName = document.createElement('div');
            fileName.className = 'attachment-name';
            fileName.textContent = attachment.name;
            
            // 文件详情
            const fileDetails = document.createElement('div');
            fileDetails.className = 'attachment-details';
            
            if (attachment.error) {
                // 如果有错误，显示错误信息
                fileDetails.innerHTML = `<span class="text-danger"><i class="bi bi-exclamation-triangle"></i> ${attachment.error}</span>`;
            } else if (attachment.warning) {
                // 如果有警告，显示警告信息
                fileDetails.innerHTML = `<span class="text-warning"><i class="bi bi-exclamation-circle"></i> ${attachment.warning}</span>`;
            } else {
                // 显示文件大小和收件人邮箱
                let detailsHTML = `<span class="text-muted">${formatFileSize(attachment.size)}</span>`;
                
                if (attachment.email) {
                    detailsHTML += `<span class="text-success ms-2"><i class="bi bi-envelope"></i> ${attachment.email}</span>`;
                }
                
                fileDetails.innerHTML = detailsHTML;
            }
            
            // 添加到文件信息区
            fileInfo.appendChild(fileName);
            fileInfo.appendChild(fileDetails);
            
            // 操作按钮
            const fileActions = document.createElement('div');
            fileActions.className = 'attachment-actions';
            
            // 预览按钮
            if (attachment.email && !attachment.error) {
                const previewButton = document.createElement('button');
                previewButton.className = 'btn btn-sm btn-outline-primary me-2';
                previewButton.innerHTML = '<i class="bi bi-eye"></i>';
                previewButton.title = '预览邮件';
                previewButton.onclick = (event) => {
                    event.preventDefault(); // 防止页面滚动到顶部
                    previewEmail(index);
                    return false;
                };
                fileActions.appendChild(previewButton);
            }
            
            // 删除按钮
            const deleteButton = document.createElement('button');
            deleteButton.className = 'btn btn-sm btn-outline-danger';
            deleteButton.innerHTML = '<i class="bi bi-trash"></i>';
            deleteButton.title = '删除附件';
            deleteButton.onclick = () => removeAttachment(index);
            fileActions.appendChild(deleteButton);
            
            // 组装文件项
            fileItem.appendChild(fileIcon);
            fileItem.appendChild(fileInfo);
            fileItem.appendChild(fileActions);
            
            // 添加到列表
            attachmentList.appendChild(fileItem);
        }
        
        // 标记该文件已处理
        processedFiles.add(index);
    });
    
    // 处理未匹配的Word文档
    attachmentFiles.forEach((attachment, index) => {
        if (processedFiles.has(index) || !attachment.is_word) {
            return;
        }
        
        // 不再尝试基于邮箱匹配，只使用文件名匹配逻辑
        // 匹配逻辑已经在matchFilesAndUpdateUI函数中实现
        
        const fileItem = document.createElement('div');
        fileItem.className = 'attachment-item';
        
        // 根据是否有错误设置不同的样式
        if (attachment.error) {
            fileItem.classList.add('attachment-error');
        }
        
        // 文件图标
        const fileIcon = document.createElement('div');
        fileIcon.className = 'attachment-icon';
        fileIcon.innerHTML = '<i class="bi bi-file-earmark-word"></i>';
        
        // 文件信息
        const fileInfo = document.createElement('div');
        fileInfo.className = 'attachment-info';
        
        // 文件名称
        const fileName = document.createElement('div');
        fileName.className = 'attachment-name';
        fileName.textContent = attachment.name;
        
        // 文件详情
        const fileDetails = document.createElement('div');
        fileDetails.className = 'attachment-details';
        
        if (attachment.error) {
            // 如果有错误，显示错误信息
            fileDetails.innerHTML = `<span class="text-danger"><i class="bi bi-exclamation-triangle"></i> ${attachment.error}</span>`;
        } else {
            // 显示文件大小和收件人邮箱
            let detailsHTML = `<span class="text-muted">${formatFileSize(attachment.size)}</span>`;
            
            if (attachment.email) {
                detailsHTML += `<span class="text-success ms-2"><i class="bi bi-envelope"></i> ${attachment.email}</span>`;
            }
            
            if (attachment.matched_excel) {
                detailsHTML += `<span class="text-info ms-2"><i class="bi bi-link"></i> 匹配: ${attachment.matched_excel}</span>`;
            } else {
                detailsHTML += `<span class="text-warning ms-2"><i class="bi bi-exclamation-circle"></i> 未匹配</span>`;
            }
            
            fileDetails.innerHTML = detailsHTML;
        }
        
        // 添加到文件信息区
        fileInfo.appendChild(fileName);
        fileInfo.appendChild(fileDetails);
        
        // 操作按钮
        const fileActions = document.createElement('div');
        fileActions.className = 'attachment-actions';
        
        // 删除按钮
        const deleteButton = document.createElement('button');
        deleteButton.className = 'btn btn-sm btn-outline-danger';
        deleteButton.innerHTML = '<i class="bi bi-trash"></i>';
        deleteButton.title = '删除附件';
        deleteButton.onclick = () => removeAttachment(index);
        fileActions.appendChild(deleteButton);
        
        // 组装文件项
        fileItem.appendChild(fileIcon);
        fileItem.appendChild(fileInfo);
        fileItem.appendChild(fileActions);
        
        // 添加到列表
        attachmentList.appendChild(fileItem);
        
        // 标记该文件已处理
        processedFiles.add(index);
    });
}

// 删除附件
function removeAttachment(index) {
    if (index >= 0 && index < attachmentFiles.length) {
        const attachment = attachmentFiles[index];
        const filesToRemove = [index]; // 要删除的文件索引列表
        
        // 如果删除的是Excel文件，同时删除与之匹配的Word文档
        if (!attachment.is_word && attachment.word_attachment) {
            const wordIndex = attachmentFiles.findIndex(att => att === attachment.word_attachment);
            if (wordIndex !== -1) {
                filesToRemove.push(wordIndex);
            }
        }
        
        // 如果删除的是Word文档，同时删除与之匹配的Excel文件
        if (attachment.is_word && attachment.matched_excel) {
            const excelIndex = attachmentFiles.findIndex(att => att.name === attachment.matched_excel);
            if (excelIndex !== -1) {
                filesToRemove.push(excelIndex);
            }
        }
        
        // 按索引从大到小排序，以便从后向前删除（避免删除时索引变化的问题）
        filesToRemove.sort((a, b) => b - a);
        
        // 删除所有相关附件
        for (const idx of filesToRemove) {
            attachmentFiles.splice(idx, 1);
        }
        
        updateAttachmentList();
        
        // 重新检查是否有重复邮箱
        checkDuplicateEmails();
        
        // 重新匹配文件
        matchFilesAndUpdateUI();
        
        // 重置文件输入框的值，使得用户可以重新上传相同的文件
        const fileInput = document.getElementById('attachmentInput');
        if (fileInput) {
            fileInput.value = '';
        }
    }
}

// 匹配文件并更新UI
function matchFilesAndUpdateUI() {
    const wordExcelWarning = document.getElementById('wordExcelWarning');
    
    // 获取所有Excel和Word文件，不考虑错误状态
    const excelAttachments = attachmentFiles.filter(att => !att.is_word);
    const wordAttachments = attachmentFiles.filter(att => att.is_word);
    
    // 重置所有匹配状态
    for (const att of attachmentFiles) {
        if (!att.is_word && att.word_attachment) {
            delete att.word_attachment;
        }
        if (att.is_word && att.matched_excel) {
            delete att.matched_excel;
        }
        // 清除匹配相关的错误和警告
        if (att.error === '找不到匹配的Excel文件') {
            delete att.error;
        }
        if (att.warning === '没有匹配的Word文档') {
            delete att.warning;
        }
    }
    
    // 检查是否有未匹配的Word文档
    let unmatchedWord = false;
    let unmatchedExcel = false;
    
    // 如果有Word文档和Excel文件，尝试匹配
    if (wordAttachments.length > 0 && excelAttachments.length > 0) {
        // 尝试匹配每个Word文档与Excel文件
        for (const wordAtt of wordAttachments) {
            // 获取Word文档的文件名（不含扩展名）
            const wordFilename = wordAtt.filename_without_ext || '';
            
            // 如果文件名为空，标记为未匹配
            if (!wordFilename) {
                unmatchedWord = true;
                wordAtt.error = '无法获取文件名';
                continue;
            }
            
            // 查找匹配的Excel文件（文件名除了最后三个字符外都相同）
            const matchedExcel = excelAttachments.find(excel => {
                const excelFilename = excel.filename_without_ext || '';
                if (!excelFilename) return false;
                
                // 获取文件名长度
                const wordLen = wordFilename.length;
                const excelLen = excelFilename.length;
                
                // 文件名太短，无法应用规则
                if (wordLen <= 3 || excelLen <= 3) return false;
                
                // 获取除了最后三个字符外的部分
                const wordPrefix = wordFilename.substring(0, wordLen - 3);
                const excelPrefix = excelFilename.substring(0, excelLen - 3);
                
                console.log(`比较: Word=${wordFilename}, Excel=${excelFilename}`);
                console.log(`前缀: Word=${wordPrefix}, Excel=${excelPrefix}`);
                console.log(`前缀长度: Word=${wordPrefix.length}, Excel=${excelPrefix.length}`);
                
                // 检查每个字符是否相同，帮助调试
                let isMatch = true;
                if (wordPrefix.length === excelPrefix.length) {
                    for (let i = 0; i < wordPrefix.length; i++) {
                        if (wordPrefix[i] !== excelPrefix[i]) {
                            console.log(`不匹配的字符在位置 ${i}: Word='${wordPrefix[i]}' (${wordPrefix.charCodeAt(i)}), Excel='${excelPrefix[i]}' (${excelPrefix.charCodeAt(i)})`);
                            isMatch = false;
                        }
                    }
                } else {
                    console.log('前缀长度不同，不匹配');
                    isMatch = false;
                }
                
                // 如果前缀相同，则认为匹配
                return isMatch;
            });
            
            if (matchedExcel) {
                // 建立双向关联
                wordAtt.matched_excel = matchedExcel.name;
                matchedExcel.word_attachment = wordAtt;
                // 清除任何错误标记
                delete wordAtt.error;
                console.log(`匹配成功: ${wordAtt.name} 匹配到 ${matchedExcel.name}`);
            } else {
                unmatchedWord = true;
                wordAtt.error = '找不到匹配的Excel文件';
                console.log(`未匹配: ${wordAtt.name} 没有找到匹配的Excel文件`);
            }
        }
        
        // 检查是否有未匹配的Excel文件
        for (const excelAtt of excelAttachments) {
            if (!excelAtt.word_attachment) {
                unmatchedExcel = true;
                excelAtt.warning = '没有匹配的Word文档';
                console.log(`未匹配: ${excelAtt.name} 没有匹配的Word文档`);
            } else {
                // 清除任何警告标记
                delete excelAtt.warning;
            }
        }
        
        // 如果有未匹配的文件，显示警告
        if (unmatchedWord || unmatchedExcel) {
            wordExcelWarning.classList.remove('d-none');
            showAlert('部分Word文档和Excel文件未能正确匹配，请检查文件名是否符合匹配规则', 'warning');
        } else {
            wordExcelWarning.classList.add('d-none');
        }
    } else if (wordAttachments.length > 0 && excelAttachments.length === 0) {
        // 如果只有Word文档没有Excel文件
        for (const wordAtt of wordAttachments) {
            if (wordAtt.email) {
                wordAtt.error = '找不到匹配的Excel文件';
            } else {
                wordAtt.error = '无法从文档中提取邮箱地址';
            }
        }
        wordExcelWarning.classList.remove('d-none');
        showAlert('没有匹配的Excel文件，请上传相应的Excel文件', 'warning');
    }
    
    // 更新附件列表显示
    updateAttachmentList();
    
    // 启用发送按钮
    document.querySelector('.btn-send').disabled = false;
}

// 检查文件匹配状态
function checkDuplicateEmails() {
}

// 清空附件列表
function clearAttachments() {
    attachmentFiles = [];
    updateAttachmentList();
    
    // 清除文件输入框
    const fileInput = document.getElementById('attachmentInput');
    if (fileInput) {
        fileInput.value = '';
    }
    
    // 隐藏警告
    document.getElementById('fileTypeWarning').classList.add('d-none');
    document.getElementById('wordExcelWarning').classList.add('d-none');
    
    // 启用发送按钮
    document.querySelector('.btn-send').disabled = false;
    
    // 重新检查是否有重复邮箱（实际上清空后应该没有重复，但为了保持一致性还是调用一下）
    checkDuplicateEmails();
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
        
        // 获取是否将邮件副本发送给自己的选项
        const ccSender = document.getElementById('ccSender').checked;
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            ccSender, // 添加是否将邮件副本发送给自己的选项
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
        
        // 获取是否将邮件副本发送给自己的选项
        const ccSender = document.getElementById('ccSender').checked;
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            ccSender, // 添加是否将邮件副本发送给自己的选项
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

// 当前预览的附件索引
let currentPreviewIndex = -1;

// 预览邮件内容
function previewEmail(index) {
    const attachment = attachmentFiles[index];
    if (!attachment || !attachment.email) return;
    
    // 记录当前预览的附件索引，用于直接发送
    currentPreviewIndex = index;
    
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
    
    // 如果有年月日范围，替换[[年月日范围]]占位符
    if (attachment.period_range) {
        previewSubject = previewSubject.replace(/\[\[年月日范围\]\]/g, attachment.period_range);
        previewContent = previewContent.replace(/\[\[年月日范围\]\]/g, attachment.period_range);
    }
    
    // 如果有公司信息，替换[[公司]]占位符
    if (attachment.company) {
        previewSubject = previewSubject.replace(/\[\[公司\]\]/g, attachment.company);
        previewContent = previewContent.replace(/\[\[公司\]\]/g, attachment.company);
    }
    
    // 更新预览模态框内容
    document.getElementById('previewRecipient').textContent = attachment.email;
    document.getElementById('previewSubject').textContent = previewSubject;
    document.getElementById('previewContent').innerHTML = previewContent;
    
    // 显示所有附件（Excel和Word）
    let attachmentText = filename;
    
    // 如果是Excel文件且有匹配的Word文档
    if (!attachment.is_word && attachment.word_attachment) {
        attachmentText = `${filename}, ${attachment.word_attachment.name}`;
    } 
    // 如果是Word文档且有匹配的Excel文件
    else if (attachment.is_word && attachment.matched_excel) {
        const excelAttachments = attachmentFiles.filter(att => !att.is_word && !att.error);
        const matchedExcel = excelAttachments.find(excel => excel.email === attachment.email);
        if (matchedExcel) {
            attachmentText = `${matchedExcel.name}, ${filename}`;
        }
    }
    
    document.getElementById('previewAttachmentName').textContent = attachmentText;
    
    // 显示预览模态框
    const previewModal = new bootstrap.Modal(document.getElementById('emailPreviewModal'));
    previewModal.show();
}

// 从预览界面发送单个邮件
async function sendEmailFromPreview() {
    // 关闭预览模态框
    const previewModal = bootstrap.Modal.getInstance(document.getElementById('emailPreviewModal'));
    previewModal.hide();
    
    // 显示加载中提示
    showLoading(true, '正在发送邮件，请稍候...');
    
    try {
        // 获取当前预览的附件
        const attachment = attachmentFiles[currentPreviewIndex];
        if (!attachment || !attachment.email) {
            showAlert('找不到有效的附件信息', 'warning');
            showLoading(false);
            return;
        }
        
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
        
        // 准备附件数据
        const attachmentData = {
            name: attachment.name,
            data: attachment.data,
            email: attachment.email,
            period: attachment.period || '',
            period_range: attachment.period_range || '',
            company: attachment.company || '',
            is_word: attachment.is_word || false
        };
        
        // 如果是Excel文件且有匹配的Word文档，添加Word附件信息
        if (!attachment.is_word && attachment.word_attachment) {
            attachmentData.word_attachment = {
                name: attachment.word_attachment.name,
                data: attachment.word_attachment.data
            };
        }
        
        // 获取是否将邮件副本发送给自己的选项
        const ccSender = document.getElementById('ccSender').checked;
        
        // 准备请求数据
        const requestData = {
            subject,
            content,
            senderEmail,
            senderPassword,
            ccSender,
            attachments: [attachmentData]
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
            showAlert('邮件发送成功！', 'success');
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
    // 移除旧的提示元素
    const oldToasts = document.querySelectorAll('.alert-toast');
    oldToasts.forEach(toast => {
        // 添加消失动画
        toast.style.animation = 'fadeOut 0.3s forwards';
        // 动画完成后移除元素
        setTimeout(() => toast.remove(), 300);
    });
    
    // 创建新的提示元素
    const toast = document.createElement('div');
    toast.className = `alert alert-${type} alert-toast`;
    toast.innerHTML = `
        <div class="d-flex align-items-center">
            <div class="me-3">
                ${type === 'success' ? '<i class="bi bi-check-circle-fill"></i>' : ''}
                ${type === 'warning' ? '<i class="bi bi-exclamation-triangle-fill"></i>' : ''}
                ${type === 'danger' ? '<i class="bi bi-x-circle-fill"></i>' : ''}
                ${type === 'info' ? '<i class="bi bi-info-circle-fill"></i>' : ''}
            </div>
            <div>${message}</div>
        </div>
    `;
    
    // 添加到文档中
    document.body.appendChild(toast);
    
    // 设置定时器移除提示
    setTimeout(() => {
        // 添加消失动画
        toast.style.animation = 'fadeOut 0.3s forwards';
        // 动画完成后移除元素
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}
