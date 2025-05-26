/**
 * 供应商账单拆分工具
 * 纯前端JavaScript逻辑，使用分批异步处理优化性能
 */

document.addEventListener('DOMContentLoaded', function() {
    // 全局变量
    const fileInput = document.getElementById('fileInput');
    const wordTemplateInput = document.getElementById('wordTemplateInput');
    const processBtn = document.getElementById('processBtn');
    const dropArea = document.getElementById('dropArea');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const resultSection = document.getElementById('resultSection');
    const resultMessage = document.getElementById('resultMessage');
    const resultTable = document.getElementById('resultTable');
    const downloadAllBtn = document.getElementById('downloadAllBtn');
    const logContainer = document.getElementById('logContainer');
    const errorMessage = document.getElementById('errorMessage');
    const logoImage = document.getElementById('logoImage');
    
    // Word模板文件路径
    const wordTemplatePath = 'static/通知函模板.docx';
    let wordTemplateContent = null;
    
    // 页面加载时预加载Word模板
    fetch(wordTemplatePath)
        .then(response => {
            if (!response.ok) {
                throw new Error(`无法加载Word模板文件: ${response.status} ${response.statusText}`);
            }
            return response.arrayBuffer();
        })
        .then(buffer => {
            wordTemplateContent = buffer;
            log('Word模板文件加载成功', 'success');
        })
        .catch(error => {
            log(`加载Word模板文件失败: ${error.message}`, 'error');
        });
    
    // 日志记录函数
    function log(message, type = 'info') {
        const timestamp = new Date().toLocaleTimeString();
        const logEntry = document.createElement('div');
        logEntry.className = `log-entry log-${type}`;
        logEntry.innerHTML = `<span class="log-time">[${timestamp}]</span> <span class="log-message">${message}</span>`;
        
        // 根据类型设置颜色
        switch(type) {
            case 'error':
                logEntry.style.color = '#ff5252';
                break;
            case 'success':
                logEntry.style.color = '#4caf50';
                break;
            case 'warning':
                logEntry.style.color = '#fb8c00';
                break;
            default:
                logEntry.style.color = '#8be9fd';
        }
        
        logContent.appendChild(logEntry);
        
        // 滚动到底部
        logContent.parentElement.scrollTop = logContent.parentElement.scrollHeight;
        
        // 显示日志区域
        logSection.classList.remove('d-none');
        
        // 在控制台也输出日志
        console.log(`[${type.toUpperCase()}] ${message}`);
    }
    
    // 存储拆分结果数据
    let splitResults = [];
    let workbook = null;
    let zip = null;
    
    // 清除结果和错误信息
    function clearResults() {
        // 清空日志内容
        logContent.innerHTML = '';
        
        // 隐藏结果和错误信息
        resultSection.classList.add('d-none');
        errorMessage.classList.add('d-none');
        
        // 重置结果数组
        splitResults = [];
        
        // 重置ZIP对象
        zip = new JSZip();
    }
    
    // 显示加载进度条
    function showLoading() {
        progressContainer.classList.remove('d-none');
        progressBar.style.width = '0%';
        progressBar.setAttribute('aria-valuenow', 0);
        log('开始处理文件...');
    }
    
    // 隐藏加载进度条
    function hideLoading() {
        progressContainer.classList.add('d-none');
    }
    
    // 更新进度条
    function updateProgress(current, total) {
        const percentage = Math.floor((current / total) * 100);
        progressBar.style.width = percentage + '%';
        progressBar.setAttribute('aria-valuenow', percentage);
        log(`处理进度: ${current}/${total} (${percentage}%)`);
    }
    
    // 数字转中文大写金额
    function numberToChinese(num) {
        if (isNaN(num)) return '';
        
        const chineseDigits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
        const chineseUnits = ['', '拾', '佰', '仟'];
        const chineseBigUnits = ['', '万', '亿', '兆'];
        
        // 将数字转换为字符串并处理小数点
        let numStr = num.toString();
        let integerPart = numStr;
        let decimalPart = '';
        
        if (numStr.includes('.')) {
            const parts = numStr.split('.');
            integerPart = parts[0];
            decimalPart = parts[1];
        }
        
        // 处理整数部分
        let result = '';
        if (integerPart === '0') {
            result = chineseDigits[0];
        } else {
            // 按照每4位分组处理
            const groups = [];
            for (let i = integerPart.length; i > 0; i -= 4) {
                groups.unshift(integerPart.substring(Math.max(0, i - 4), i));
            }
            
            for (let i = 0; i < groups.length; i++) {
                let groupResult = '';
                const group = groups[i];
                
                for (let j = 0; j < group.length; j++) {
                    const digit = parseInt(group[j]);
                    const unit = chineseUnits[group.length - j - 1];
                    
                    if (digit !== 0) {
                        groupResult += chineseDigits[digit] + unit;
                    } else {
                        // 处理连续的零
                        if (groupResult.length > 0 && groupResult[groupResult.length - 1] !== chineseDigits[0]) {
                            groupResult += chineseDigits[0];
                        }
                    }
                }
                
                // 去除末尾的零
                if (groupResult.endsWith(chineseDigits[0])) {
                    groupResult = groupResult.substring(0, groupResult.length - 1);
                }
                
                if (groupResult.length > 0) {
                    result += groupResult + chineseBigUnits[groups.length - i - 1];
                }
            }
        }
        
        // 处理小数部分
        if (decimalPart.length > 0) {
            result += '元';
            
            if (decimalPart[0] !== '0') {
                result += chineseDigits[parseInt(decimalPart[0])] + '角';
            }
            
            if (decimalPart.length > 1 && decimalPart[1] !== '0') {
                result += chineseDigits[parseInt(decimalPart[1])] + '分';
            }
        } else {
            result += '元整';
        }
        
        return result;
    }
    
    // 生成ZIP文件（不自动下载）
    function generateZipFile(zip) {
        log('开始生成ZIP文件...');
        // 显示进度条
        progressContainer.classList.remove('d-none');
        progressBar.style.width = '50%';
        progressBar.setAttribute('aria-valuenow', 50);
        
        // 生成ZIP文件
        zip.generateAsync({type: 'blob'})
            .then(function(content) {
                // 更新进度条到100%
                progressBar.style.width = '100%';
                progressBar.setAttribute('aria-valuenow', 100);
                
                const zipSize = (content.size / 1024 / 1024).toFixed(2);
                log(`ZIP文件生成成功，大小: ${zipSize} MB`, 'success');
                
                // 将生成的ZIP内容保存到全局变量，供下载按钮使用
                window.zipContent = content;
                window.zipFilename = `供应商账单拆分结果_${new Date().toISOString().slice(0, 10)}.zip`;
                
                // 显示结果
                showResults(splitResults);
            })
            .catch(function(error) {
                log(`生成ZIP文件时出错: ${error.message}`, 'error');
            });
    }
    
    // 文件选择事件
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    
    fileInput.addEventListener('change', function() {
        if (fileInput.files.length > 0) {
            const fileName = fileInput.files[0].name;
            // 更新文件名显示
            fileNameDisplay.textContent = fileName;
            fileNameDisplay.classList.add('text-primary');
            fileNameDisplay.classList.remove('text-muted');
            
            // 启用处理按钮
            processBtn.disabled = false;
            
            log(`已选择文件: ${fileName}`);
        } else {
            // 重置文件名显示
            fileNameDisplay.textContent = '未选择任何文件';
            fileNameDisplay.classList.remove('text-primary');
            fileNameDisplay.classList.add('text-muted');
            
            // 禁用处理按钮
            processBtn.disabled = true;
        }
    });
    
    // 注意：已移除Word模板文件上传功能，现在直接从固定路径读取
    
    // 拖放区域事件
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight() {
        dropArea.classList.add('drag-highlight');
        dropArea.style.backgroundColor = '#f8f9fa';
        dropArea.style.borderColor = '#0d6efd';
        dropArea.style.borderStyle = 'dashed';
    }
    
    function unhighlight() {
        dropArea.classList.remove('drag-highlight');
        dropArea.style.backgroundColor = '';
        dropArea.style.borderColor = '';
        dropArea.style.borderStyle = '';
    }
    
    dropArea.addEventListener('drop', handleDrop, false);
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            // 将拖放的文件设置到文件输入框
            fileInput.files = files;
            
            // 更新文件名显示
            const fileName = files[0].name;
            fileNameDisplay.textContent = fileName;
            fileNameDisplay.classList.add('text-primary');
            fileNameDisplay.classList.remove('text-muted');
            
            // 启用处理按钮
            processBtn.disabled = false;
            
            log(`已拖放文件: ${fileName}`);
        }
    }
    
    // 处理按钮点击事件
    processBtn.addEventListener('click', function() {
        handleFileUpload(null);
    });
    
    // 处理上传的Excel文件
    function handleFileUpload(event) {
        if (event) event.preventDefault();
        
        const fileInput = document.getElementById('fileInput');
        const file = fileInput.files[0];
        
        if (!file) {
            showError('请选择一个Excel文件');
            return;
        }
        
        if (!/\.(xlsx|xls)$/i.test(file.name)) {
            showError('请选择有效的Excel文件 (.xlsx 或 .xls)');
            return;
        }
        
        clearResults();
        showLoading();
        logSection.classList.remove('d-none');
        log(`开始处理文件: ${file.name}`);
        
        // 使用FileReader读取文件
        const reader = new FileReader();
        
        // 设置进度事件
        reader.onprogress = function(e) {
            if (e.lengthComputable) {
                const percentLoaded = Math.round((e.loaded / e.total) * 100);
                // 更新进度条
                progressBar.style.width = `${Math.min(percentLoaded, 30)}%`;
                progressBar.setAttribute('aria-valuenow', Math.min(percentLoaded, 30));
                
                // 每10%记录一次日志
                if (percentLoaded % 10 === 0) {
                    log(`读取文件进度: ${Math.round((e.loaded / e.total) * 100)}%`);
                }
            }
        };
        
        reader.onload = function(e) {
            try {
                // 更新进度条
                progressBar.style.width = '30%';
                progressBar.setAttribute('aria-valuenow', 30);
                
                log('文件读取完成，开始解析Excel数据');
                
                // 使用ExcelJS库解析Excel文件
                const data = new Uint8Array(e.target.result);
                workbook = new ExcelJS.Workbook();
                
                workbook.xlsx.load(data).then(function() {
                    log(`成功解析Excel文件，包含 ${workbook.worksheets.length} 个工作表`, 'success');
                    
                    // 处理Excel文件
                    processExcelWorkbook(workbook);
                }).catch(error => {
                    log(`解析Excel文件时出错: ${error.message}`, 'error');
                    showError(`解析Excel文件时出错: ${error.message}`);
                    progressContainer.classList.add('d-none');
                });
            } catch (error) {
                log(`处理Excel文件时出错: ${error.message}`, 'error');
                showError(`处理Excel文件时出错: ${error.message}`);
                progressContainer.classList.add('d-none');
            }
        };
        
        reader.onerror = function() {
            log('读取文件时出错', 'error');
            showError('读取文件时出错');
            progressContainer.classList.add('d-none');
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // 处理Excel工作簿
    function processExcelWorkbook(workbook) {
        try {
            // 更新进度条
            progressBar.style.width = '40%';
            progressBar.setAttribute('aria-valuenow', 40);
            
            log('开始处理Excel数据...');
            
            // 获取第一个工作表
            const worksheet = workbook.worksheets[0];
            
            log(`正在读取工作表: ${worksheet.name}`);
            
            // 将工作表转换为JSON
            // 使用setTimeout来避免UI卡顿
            setTimeout(() => {
                try {
                    const jsonData = [];
                    const headers = [];
                    
                    // 获取标题行
                    worksheet.getRow(1).eachCell((cell, colNumber) => {
                        headers[colNumber - 1] = cell.value;
                    });
                    
                    // 遍历数据行并转换为JSON
                    worksheet.eachRow((row, rowNumber) => {
                        if (rowNumber > 1) { // 跳过标题行
                            const rowData = {};
                            row.eachCell((cell, colNumber) => {
                                rowData[headers[colNumber - 1]] = cell.value;
                            });
                            jsonData.push(rowData);
                        }
                    });
                    
                    log(`成功读取 ${jsonData.length} 行数据`);
                    
                    if (jsonData.length === 0) {
                        throw new Error('文件不包含数据');
                    }
                    
                    // 检查是否有"供应商"列
                    if (!jsonData[0].hasOwnProperty('供应商')) {
                        throw new Error('文件中没有"供应商"列');
                    }
                    
                    log('文件包含有效的"供应商"列，开始分组数据', 'success');
                    
                    // 更新进度条
                    progressBar.style.width = '50%';
                    progressBar.setAttribute('aria-valuenow', 50);
                    
                    // 分批处理数据分组
                    groupDataByBatch(jsonData, workbook);
                } catch (error) {
                    log(`处理Excel数据时出错: ${error.message}`, 'error');
                    showError(`处理Excel文件时出错: ${error.message}`);
                    progressContainer.classList.add('d-none');
                }
            }, 100); // 给UI线程一点时间更新
        } catch (error) {
            log(`处理Excel文件时出错: ${error.message}`, 'error');
            showError(`处理Excel文件时出错: ${error.message}`);
            progressContainer.classList.add('d-none');
        }
    }
    
    // 分批处理数据分组
    function groupDataByBatch(jsonData, workbook) {
        const batchSize = 1000; // 每批处理的行数
        const totalRows = jsonData.length;
        let processedRows = 0;
        const supplierGroups = {};
        let rowCount = 0;
        
        log('正在按供应商分组数据...');
        
        // 分批处理函数
        function processBatch(startIndex) {
            const endIndex = Math.min(startIndex + batchSize, totalRows);
            
            // 处理当前批次的数据
            for (let i = startIndex; i < endIndex; i++) {
                const row = jsonData[i];
                
                // 每处理500行记录一次日志
                if (processedRows > 0 && processedRows % 500 === 0) {
                    log(`已处理 ${processedRows} 行数据...`);
                }
                
                const supplier = row['供应商'];
                if (supplier && supplier.trim() !== '') {
                    if (!supplierGroups[supplier]) {
                        supplierGroups[supplier] = [];
                    }
                    supplierGroups[supplier].push(row);
                    rowCount++;
                }
                
                processedRows++;
            }
            
            // 检查是否所有数据已处理完毕
            if (endIndex < totalRows) {
                // 还有数据需要处理，使用setTimeout安排下一批处理
                setTimeout(() => {
                    processBatch(endIndex);
                }, 0); // 使用 0 毫秒让浏览器有机会响应其他事件
            } else {
                // 所有数据处理完毕，继续处理供应商文件
                finishGroupingAndProcessSuppliers(supplierGroups, rowCount, workbook);
            }
        }
        
        // 开始处理第一批
        processBatch(0);
    }
    
    // 完成分组并处理供应商文件
    function finishGroupingAndProcessSuppliers(supplierGroups, rowCount, workbook) {
        // 创建ZIP实例
        zip = new JSZip();
        log('创建ZIP文件实例成功');
        
        // 清空结果数组
        splitResults = [];
        
        // 获取供应商列表
        const suppliers = Object.keys(supplierGroups);
        
        log(`找到 ${suppliers.length} 个不同的供应商，共 ${rowCount} 行有效数据`, 'success');
        
        // 计算每个供应商的进度比例
        const progressPerSupplier = 40 / suppliers.length;
        
        log('开始为每个供应商创建Excel文件...');
        
        // 分批处理供应商
        processSuppliersBatch(suppliers, 0, supplierGroups, progressPerSupplier, workbook);
    }
    
    // 生成Word通知函文档（使用Docxtemplater）
    function generateWordDocument(supplier, settlementPeriod, totalAmount, platformFeeRate, platformFee, email) {
        log('开始生成Word通知函文档...');
        
        // 使用dayjs处理日期
        const periodParts = settlementPeriod.split('-');
        
        // 解析结算期间结束日期
        const normalized = periodParts[1].replace(/年|月/g, "-").replace("日", "");
        let endDate = dayjs(normalized);
        
        // 计算下一个月的第一天
        const nextMonth = endDate.add(1, 'month').date(1);
        
        // 格式化为中文日期字符串
        const nextMonthFirstDay = `${nextMonth.year()}年${nextMonth.month() + 1}月1日`;
        log(`下一个月第一天: ${nextMonthFirstDay}`);
        
        // 当前日期
        const now = dayjs();
        const currentDate = `${now.year()}年${now.month() + 1}月${now.date()}日`;
        log(`当前日期: ${currentDate}`);
        
        // 准备数据字典，用于替换模板中的占位符
        const settlementPeriodEnd = periodParts[1]; // 结算期间的结束日期
        const platformFeeChinese = numberToChinese(parseFloat(platformFee)); // 平台费用的中文大写
        
        // 准备模板数据
        const templateData = {
            supplier: String(supplier || ''),
            settlementPeriod: String(settlementPeriod || ''),
            settlementPeriodEnd: String(settlementPeriodEnd || ''),
            totalAmount: String(totalAmount || ''),
            platformFeeRate: String(platformFeeRate || ''),
            platformFee: String(platformFee || ''),
            platformFeeChinese: String(platformFeeChinese || ''),
            email: typeof email === 'object' ? String(email.text || '') : String(email || ''),
            currentDate: String(currentDate || ''),
            nextMonthFirstDay: String(nextMonthFirstDay || '')
        };
        
        // 打印模板变量以便调试
        Object.keys(templateData).forEach(key => {
            log(`模板变量 ${key} = ${templateData[key]}`);
        });
        
        // 使用Docxtemplater处理Word模板
        return new Promise((resolve, reject) => {
            // 加载Word模板文件
            fetch(wordTemplatePath)
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`无法加载Word模板文件: ${response.status}`);
                    }
                    return response.arrayBuffer();
                })
                .then(templateBuffer => {
                    log('模板文件加载成功，开始处理...');
                    
                    // 检查PizZip是否可用
                    if (!window.PizZip) {
                        log('PizZip未定义，尝试使用备用库...', 'warning');
                        // 尝试使用备用库
                        window.PizZip = window.Pizzip || window.JSZip;
                        
                        if (!window.PizZip) {
                            throw new Error('PizZip库未加载，无法处理Word文档');
                        }
                    }
                    
                    // 使用PizZip加载Word文档
                    const zip = new window.PizZip(templateBuffer);
                    log('PizZip实例创建成功');
                    
                    // 检查docxtemplater是否可用
                    if (!window.docxtemplater) {
                        throw new Error('docxtemplater库未加载，无法处理Word模板');
                    }
                    
                    // 创建Docxtemplater实例
                    log('开始创建docxtemplater实例...');
                    
                    // 使用自定义分隔符来避免标签重复问题
                    const doc = new window.docxtemplater();
                    doc.loadZip(zip);
                    log('docxtemplater实例创建成功');
                    
                    // 打印日志以便调试
                    Object.keys(templateData).forEach(key => {
                        log(`设置模板变量 ${key} = ${templateData[key] || '(空)'}`);
                    });
                    
                    // 设置模板变量
                    doc.setData(templateData);
                    
                    // 渲染模板
                    doc.render();
                    log('模板渲染成功');
                    
                    // 获取生成的Word文档
                    const generatedDoc = doc.getZip().generate({
                        type: 'blob',
                        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });
                    log('Word文档生成成功');
                    resolve(generatedDoc);
                })
                .catch(error => {
                    log(`处理Word模板时出错: ${error.message}`, 'error');
                    reject(error);
                });
        });
    }
    
    // 分批处理供应商
    function processSuppliersBatch(suppliers, startIndex, supplierGroups, progressPerSupplier, excelWorkbook) {
        const batchSize = 5; // 每批处理的供应商数量
        const endIndex = Math.min(startIndex + batchSize, suppliers.length);
        let processedCount = 0; // 已处理的供应商数量
        
        log(`开始处理供应商批次: ${startIndex + 1} 至 ${endIndex} / ${suppliers.length}`);
        
        // 创建ZIP文件对象
        if (!zip) {
            zip = new JSZip();
        }
        
        // 处理当前批次的供应商
        for (let i = startIndex; i < endIndex; i++) {
            const supplier = suppliers[i];
            const index = i;
            log(`处理供应商 [${index+1}/${suppliers.length}]: ${supplier}`);
            
            const supplierData = supplierGroups[supplier];
            log(`该供应商共有 ${supplierData.length} 条订单数据`);
            
            // 获取结算期间
            const settlementPeriod = supplierData[0]['结算周期'] || '未知周期';
            
            // 计算汇总数据
            let totalAmount = 0;
            let platformFeeRate = 0;
            
            if (supplierData[0].hasOwnProperty('开票总价')) {
                totalAmount = supplierData.reduce((sum, row) => sum + (parseFloat(row['开票总价']) || 0), 0);
            }
            
            if (supplierData[0].hasOwnProperty('结算比例')) {
                platformFeeRate = parseFloat(supplierData[0]['结算比例']) || 0;
            }
            
            const platformFee = totalAmount * platformFeeRate;
            
            // 创建文件名
            const filename = `${settlementPeriod}${supplier}服务费对账表.xlsx`;
            log(`创建文件: ${filename}`);
            
            // 格式化金额，保留两位小数
            const formattedTotalAmount = totalAmount.toFixed(2);
            const formattedPlatformFee = platformFee.toFixed(2);
            
            // 使用ExcelJS创建新的工作簿
            const workbook = new ExcelJS.Workbook();
            
            // 创建结算表
            const worksheet1 = workbook.addWorksheet('结算表');
            
            // 设置列宽
            worksheet1.columns = [
                { width: 30 },
                { width: 40 },
                { width: 10 },
                { width: 20 },
                { width: 10 }
            ];
            
            // 设置行高
            for (let i = 1; i <= 8; i++) {
                worksheet1.getRow(i).height = 60;
            }
            
            // 添加结算表数据
            const settlementData = [
                ['', '佳沃思平台使用费结算单', '', '', ''],
                ['公司：', supplier, '', '', ''],
                ['结算期间：', settlementPeriod, '', '', ''],
                ['线下账款总交易额：', formattedTotalAmount, '', '', ''],
                ['计算平台使用费交易总额：', formattedTotalAmount, '', '', ''],
                ['服务费比例：', platformFeeRate, '', '', ''],
                ['总应付平台使用费：', formattedPlatformFee, '', '', ''],
                ['账单收件箱：', supplierData[0]['账单收件箱'] || '', '', '', '']
            ];
            
            // 将数据添加到工作表
            settlementData.forEach((row, rowIndex) => {
                worksheet1.getRow(rowIndex + 1).values = row;
            });
            
            // 将所有单元格设置为垂直居中和加粗
            for (let i = 1; i <= 8; i++) {
                for (let j = 1; j <= 5; j++) {
                    const cell = worksheet1.getCell(i, j);
                    cell.alignment = { vertical: 'middle' };
                    cell.font = { bold: true };
                }
            }
            
            // 设置标题单元格样式
            const titleCell = worksheet1.getCell('B1');
            titleCell.font = { bold: true, size: 18 }; // 将字体调整为18
            titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
            
            // 合并A2和B2单元格
            worksheet1.mergeCells('A2:B2');
            const companyCell = worksheet1.getCell('A2');
            companyCell.value = `公司：${supplier}`;
            companyCell.alignment = { vertical: 'middle' };
            
            // 设置B3右对齐和垂直居中
            worksheet1.getCell('B3').alignment = { vertical: 'middle', horizontal: 'right' };
            
            // 设置金额单元格右对齐
            worksheet1.getCell('B4').alignment = { vertical: 'middle', horizontal: 'right' };
            worksheet1.getCell('B5').alignment = { vertical: 'middle', horizontal: 'right' };
            worksheet1.getCell('B7').alignment = { vertical: 'middle', horizontal: 'right' };
            
            // 设置邮箱单元格样式
            const emailCell = worksheet1.getCell('B8');
            if (emailCell.value) {
                emailCell.font = { bold: true, color: { argb: '0000FF' }, underline: true };
                emailCell.alignment = { vertical: 'middle', horizontal: 'right' }; // 设置B8右对齐和垂直居中
            }
            
            // 添加logo图片
            try {
                const logoImage = document.getElementById('logoImage');
                
                if (logoImage && logoImage.complete) {
                    // 创建临时画布获取图片数据
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');
                    canvas.width = 200;
                    canvas.height = 50;
                    ctx.drawImage(logoImage, 0, 0, 200, 50);
                    
                    // 获取图片的base64数据
                    const imageBase64 = canvas.toDataURL('image/png').split(',')[1];
                    
                    // 添加图片到工作表
                    const logoId = workbook.addImage({
                        base64: imageBase64,
                        extension: 'png',
                    });
                    
                    // 将图片添加到A1单元格，并设置水平和垂直居中
                    worksheet1.addImage(logoId, {
                        tl: { col: 0, row: 0 },
                        ext: { width: 200, height: 50 },
                        editAs: 'oneCell' // 使图片作为单元格内容处理
                    });
                    
                    // 设置A1单元格居中
                    const logoCell = worksheet1.getCell('A1');
                    logoCell.alignment = { vertical: 'middle', horizontal: 'center' };
                    
                    log('成功添加Logo图片到结算表', 'success');
                } else {
                    log('未找到Logo图片或图片未加载完成', 'warning');
                }
            } catch (error) {
                log(`添加图片时出错: ${error.message}`, 'warning');
            }
            
            // 创建已发货表 - 保持原始表格的字段和格式
            log('开始生成已发货表，保持原始表格结构');
            
            // 创建已发货表
            const worksheet2 = workbook.addWorksheet('已发货表');
            
            // 获取原始表格的列名
            const originalHeaders = [];
            try {
                // 获取原始表头
                excelWorkbook.worksheets[0].getRow(1).eachCell((cell, colNumber) => {
                    originalHeaders[colNumber - 1] = cell.value;
                });
            } catch (error) {
                log(`获取原始表头时出错: ${error.message}`, 'error');
            }
            
            // 创建sheet2的数据，保持原始列结构
            // 添加表头到工作表
            worksheet2.addRow(originalHeaders);
            
            // 过滤出当前供应商的数据
            log(`开始过滤供应商 "${supplier}" 的数据行...`);
            let matchedRows = 0;
            
            // 遍历数据，只保留当前供应商的数据行
            for (const row of supplierData) {
                const rowData = [];
                
                // 按原始列顺序填充数据
                for (const header of originalHeaders) {
                    rowData.push(row[header] || '');
                }
                
                worksheet2.addRow(rowData);
                matchedRows++;
                
                // 每100行记录一次日志
                if (matchedRows % 100 === 0) {
                    log(`已过滤出 ${matchedRows} 行数据...`);
                }
            }
            
            log(`已发货表生成完成，共包含 ${matchedRows} 行数据`, 'success');
            
            log('正在生成Excel文件...');
            
            // 将Excel文件转换为二进制数据
            workbook.xlsx.writeBuffer().then(async excelBuffer => {
                // 添加Excel文件到ZIP
                zip.file(filename, excelBuffer);
                log(`已将Excel文件 ${filename} 添加到ZIP包中，文件大小: ${(excelBuffer.byteLength / 1024).toFixed(2)} KB`);
                
                // 生成Word文档（DOCX格式）
                const wordFilename = `${settlementPeriod}${supplier}服务费通知函.docx`;
                log(`开始生成Word文档: ${wordFilename}`);
                
                // 调用生成Word文档的函数
                const wordBuffer = await generateWordDocument(
                    supplier, 
                    settlementPeriod, 
                    formattedTotalAmount, 
                    platformFeeRate, 
                    formattedPlatformFee, 
                    supplierData[0]['账单收件箱']['text'] || supplierData[0]['账单收件箱'] || ''
                );
                
                if (wordBuffer) {
                    // 添加Word文档到ZIP
                    zip.file(wordFilename, wordBuffer);
                    log(`已将Word文档 ${wordFilename} 添加到ZIP包中，文件大小: ${(wordBuffer.size / 1024).toFixed(2)} KB`, 'success');
                }
                
                // 添加到结果列表
                splitResults.push({
                        supplier: supplier,
                        filename: filename,
                        word_filename: wordFilename,
                        order_count: supplierData.length,
                        total_amount: totalAmount,
                        platform_fee: platformFee,
                        excel_data: excelBuffer
                    });
                    
                    // 检查是否所有供应商都已处理完毕
                    processedCount++;
                    updateProgress(startIndex + processedCount, suppliers.length);
                    
                    if (processedCount === batchSize || startIndex + processedCount === suppliers.length) {
                        // 当前批次处理完毕或所有供应商都已处理完毕
                        if (startIndex + processedCount < suppliers.length) {
                            // 还有更多供应商需要处理
                            setTimeout(function() {
                                processSuppliersBatch(suppliers, startIndex + processedCount, supplierGroups, progressPerSupplier, excelWorkbook);
                            }, 100);
                        } else {
                            // 所有供应商都已处理完毕，生成ZIP文件
                            generateZipFile(zip);
                        }
                    }

                
                // 注意：showResults函数已在generateZipFile函数中调用，这里不需要重复调用
            }, 1000); // 等待1秒，给异步操作时间完成
        }
    }
    
    // 显示处理结果
    function showResults(data) {
        log('开始显示拆分结果...');
        
        // 显示结果区域
        resultSection.classList.remove('d-none');
        
        // 设置结果消息
        resultMessage.textContent = `成功拆分为 ${data.length} 个供应商文件`;
        
        // 清空结果表格
        resultTable.innerHTML = '';
        
        // 填充结果表格
        log('生成结果表格...');
        
        data.forEach(file => {
            const row = document.createElement('tr');
            
            // 格式化金额
            const formattedAmount = new Intl.NumberFormat('zh-CN', {
                style: 'currency',
                currency: 'CNY'
            }).format(file.total_amount);
            
            const formattedFee = new Intl.NumberFormat('zh-CN', {
                style: 'currency',
                currency: 'CNY'
            }).format(file.platform_fee);
            
            row.innerHTML = `
                <td>${file.supplier}</td>
                <td class="text-right">${file.order_count}</td>
                <td class="currency-column">${formattedAmount}</td>
                <td class="currency-column">${formattedFee}</td>
                <td>
                    <div class="btn-group">
                        <button class="btn btn-sm btn-outline-primary btn-download" data-index="${data.indexOf(file)}" data-type="excel" title="下载Excel文件">
                            <i class="bi bi-file-earmark-excel"></i> Excel
                        </button>
                        <button class="btn btn-sm btn-outline-info btn-download" data-index="${data.indexOf(file)}" data-type="word" title="下载Word文档">
                            <i class="bi bi-file-earmark-word"></i> Word
                        </button>
                    </div>
                </td>
            `;
            
            resultTable.appendChild(row);
        });
        
        // 添加下载按钮事件
        document.querySelectorAll('.btn-download').forEach(btn => {
            btn.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                const fileType = this.getAttribute('data-type') || 'excel';
                downloadSingleFile(index, fileType);
            });
        });
        
        log('结果显示完成，可以单独下载或打包下载所有文件', 'success');
        
        // 隐藏进度条
        setTimeout(() => {
            progressContainer.classList.add('d-none');
        }, 500);
    }
    
    // 下载所有文件
    downloadAllBtn.addEventListener('click', function() {
        if (!splitResults || splitResults.length === 0) {
            showError('没有可下载的文件');
            log('尝试下载所有文件失败，没有可用文件', 'error');
            return;
        }
        
        // 检查是否已经生成了ZIP文件
        if (window.zipContent) {
            // 如果已经生成了ZIP文件，直接下载
            saveAs(window.zipContent, window.zipFilename);
            log('保存ZIP文件成功', 'success');
            return;
        }
        
        // 如果还没有生成ZIP文件，则生成并下载
        if (!zip) {
            showError('没有可下载的文件');
            log('尝试下载所有文件失败，没有可用文件', 'error');
            return;
        }
        
        log(`开始打包所有 ${splitResults.length} 个文件...`);
        
        // 显示进度条
        progressContainer.classList.remove('d-none');
        progressBar.style.width = '50%';
        progressBar.setAttribute('aria-valuenow', 50);
        
        // 生成ZIP文件
        zip.generateAsync({type: 'blob'})
            .then(function(content) {
                // 更新进度条到100%
                progressBar.style.width = '100%';
                progressBar.setAttribute('aria-valuenow', 100);
                
                const zipSize = (content.size / 1024 / 1024).toFixed(2);
                log(`ZIP文件生成成功，大小: ${zipSize} MB`, 'success');
                
                // 保存到全局变量
                window.zipContent = content;
                window.zipFilename = `供应商账单拆分结果_${new Date().toISOString().slice(0, 10)}.zip`;
                
                // 保存ZIP文件
                saveAs(content, window.zipFilename);
                log('保存ZIP文件成功', 'success');
                
                // 隐藏进度条
                setTimeout(() => {
                    progressContainer.classList.add('d-none');
                }, 500);
            })
            .catch(function(error) {
                log(`生成ZIP文件时出错: ${error.message}`, 'error');
                progressContainer.classList.add('d-none');
            });
    });
    
    // 下载单个供应商文件
    function downloadSingleFile(index, fileType = 'excel') {
        if (index < 0 || index >= splitResults.length) {
            showError('文件索引无效');
            log('尝试下载单个文件失败，索引无效', 'error');
            return;
        }
        
        const fileData = splitResults[index];
        
        // 显示进度条
        progressContainer.classList.remove('d-none');
        progressBar.style.width = '50%';
        progressBar.setAttribute('aria-valuenow', 50);
        
        try {
            let blob, filename, fileSize, mimeType;
            
            if (fileType === 'excel') {
                // 下载Excel文件
                log(`开始下载供应商 "${fileData.supplier}" 的Excel文件...`);
                mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                blob = new Blob([fileData.excel_data], { type: mimeType });
                filename = fileData.filename;
                fileSize = (fileData.excel_data.byteLength / 1024).toFixed(2);
            } else if (fileType === 'word') {
                // 下载Word文档
                log(`开始下载供应商 "${fileData.supplier}" 的Word文档...`);
                
                // 从压缩包中获取Word文档
                return zip.file(fileData.word_filename).async('blob').then(content => {
                    mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                    blob = new Blob([content], { type: mimeType });
                    filename = fileData.word_filename;
                    fileSize = (content.size / 1024).toFixed(2);
                    
                    // 更新进度条到100%
                    progressBar.style.width = '100%';
                    progressBar.setAttribute('aria-valuenow', 100);
                    
                    log(`Word文档 "${filename}" 准备完成，大小: ${fileSize} KB`, 'success');
                    
                    // 使用FileSaver保存文件
                    saveAs(blob, filename);
                    log(`开始下载Word文档 "${filename}"`, 'success');
                    
                    // 隐藏进度条
                    setTimeout(() => {
                        progressContainer.classList.add('d-none');
                    }, 500);
                }).catch(error => {
                    log(`下载Word文档时出错: ${error.message}`, 'error');
                    showError('下载Word文档时出错: ' + error.message);
                    progressContainer.classList.add('d-none');
                });
            } else {
                throw new Error('不支持的文件类型');
            }
            
            // 更新进度条到100%
            progressBar.style.width = '100%';
            progressBar.setAttribute('aria-valuenow', 100);
            
            log(`文件 "${filename}" 准备完成，大小: ${fileSize} KB`, 'success');
            
            // 使用FileSaver保存文件
            saveAs(blob, filename);
            log(`开始下载文件 "${filename}"`, 'success');
            
            // 隐藏进度条
            setTimeout(() => {
                progressContainer.classList.add('d-none');
            }, 500);
        } catch (error) {
            log(`下载文件时出错: ${error.message}`, 'error');
            showError('下载文件时出错: ' + error.message);
            progressContainer.classList.add('d-none');
        }
    }
    
    // 显示错误信息
    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove('d-none');
        
        // 同时在日志中记录错误
        log(message, 'error');
        
        // 5秒后自动隐藏错误信息
        setTimeout(() => {
            errorMessage.classList.add('d-none');
        }, 5000);
    }
});
