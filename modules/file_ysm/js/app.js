// 全局变量
let processedFiles = [];
let zip = null;

// DOM元素
const excelFileInput = document.getElementById('excelFile');
const processButton = document.getElementById('processButton');
const progressSection = document.querySelector('.progress-section');
const progressBar = document.getElementById('progressBar');
const statusText = document.getElementById('statusText');
const logContent = document.getElementById('logContent');
const resultSection = document.getElementById('resultSection');
const resultSummary = document.getElementById('resultSummary');
const resultTable = document.getElementById('resultTable');
const downloadAllButton = document.getElementById('downloadAllButton');
// 添加对账月输入框元素
const billingMonthInput = document.getElementById('billingMonth');

// 初始化事件监听
document.addEventListener('DOMContentLoaded', function() {
    processButton.addEventListener('click', processExcelFile);
    downloadAllButton.addEventListener('click', downloadAllFiles);
});

// 日志函数
function log(message) {
    const now = new Date();
    const timeString = now.toLocaleTimeString();
    logContent.innerHTML += `[${timeString}] ${message}\n`;
    logContent.scrollTop = logContent.scrollHeight;
}

// 显示警告或错误消息
function showAlert(message, type = 'danger') {
    const alertContainer = document.getElementById('alertContainer');
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    `;
    alertContainer.appendChild(alertDiv);
    
    // 5秒后自动关闭
    setTimeout(() => {
        alertDiv.classList.remove('show');
        setTimeout(() => alertDiv.remove(), 150);
    }, 5000);
}

// 更新进度条
function updateProgress(percent, status) {
    progressBar.style.width = `${percent}%`;
    progressBar.textContent = `${Math.round(percent)}%`;
    progressBar.setAttribute('aria-valuenow', percent);
    statusText.textContent = status;
}

// 处理Excel文件
async function processExcelFile() {
    const file = excelFileInput.files[0];
    if (!file) {
        showAlert('请先选择Excel文件');
        return;
    }
    
    // 验证对账月是否已输入
    const billingMonth = billingMonthInput.value.trim();
    if (!billingMonth) {
        showAlert('请输入对账月');
        return;
    }
    
    try {
        // 重置状态
        processedFiles = [];
        zip = new JSZip();
        resultTable.innerHTML = '';
        resultSection.style.display = 'none';
        progressSection.style.display = 'block';
        processButton.disabled = true;
        updateProgress(0, '准备处理...');
        
        log('开始处理文件...');
        
        // 读取Excel文件
        log('正在读取Excel文件...');
        const data = await readExcelFile(file);
        
        // 查找诊所列
        let clinicColumn = null;
        for (const col of data.columns) {
            if (col.includes('诊所')) {
                clinicColumn = col;
                break;
            }
        }
        
        if (!clinicColumn) {
            throw new Error('未找到包含"诊所"的列名');
        }
        
        // 按诊所分组
        log('正在按诊所分组数据...');
        const grouped = groupByClinic(data, clinicColumn);
        
        const totalClinics = Object.keys(grouped).length;
        log(`共读取到 ${totalClinics} 个诊所的数据`);
        
        // 处理每个诊所的数据
        let processedCount = 0;
        for (const clinicName in grouped) {
            if (clinicName === 'undefined' || clinicName === '' || clinicName === 'null') continue;
            
            const clinicData = grouped[clinicName];
            updateProgress((processedCount / totalClinics) * 100, `正在处理 ${processedCount+1}/${totalClinics}`);
            
            // 生成Excel文件，传入对账月
            log(`正在生成诊所 '${clinicName}' 的Excel文件...`);
            const excelBlob = await generateExcelFile(clinicName, clinicData, data.columns, billingMonth);
            
            // 安全的文件名
            const safeClinicName = clinicName.replace(/[\\/:*?"<>|]/g, '');
            const fileName = `${safeClinicName}${billingMonth}.xlsx`;
            
            // 添加到ZIP
            zip.file(fileName, excelBlob);
            
            // 添加到处理结果
            processedFiles.push({
                name: clinicName,
                fileName: fileName,
                blob: excelBlob
            });
            
            processedCount++;
        }
        
        // 完成处理
        updateProgress(100, '处理完成');
        log(`所有Excel文件处理完成，共 ${processedFiles.length} 个文件`);
        
        // 显示结果
        displayResults();
        
    } catch (error) {
        log(`处理过程中出错: ${error.message}`);
        showAlert(`处理过程中出错: ${error.message}`);
        updateProgress(0, '处理失败');
    } finally {
        processButton.disabled = false;
    }
}

// 读取Excel文件
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 提取列名和数据
                const columns = jsonData[0];
                const rows = jsonData.slice(1);
                
                // 转换为对象数组
                const result = rows.map(row => {
                    const obj = {};
                    columns.forEach((col, i) => {
                        obj[col] = row[i];
                    });
                    return obj;
                });
                
                resolve({ data: result, columns });
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = function() {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// 按诊所分组数据
function groupByClinic(data, clinicColumn) {
    const grouped = {};
    
    data.data.forEach(row => {
        const clinicName = row[clinicColumn] || '未知诊所';
        if (!grouped[clinicName]) {
            grouped[clinicName] = [];
        }
        grouped[clinicName].push(row);
    });
    
    return grouped;
}

// 生成Excel文件
async function generateExcelFile(clinicName, data, columns, billingMonth) {
    // 创建新的工作簿
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(clinicName);
    
    // 设置列宽
    worksheet.columns = [
        { width: 8 },  // A列
        { width: 15 }, // B列
        { width: 15 }, // C列
        { width: 15 }, // D列
        { width: 15 }, // E列
        { width: 30 }, // F列
        { width: 15 }, // G列
        { width: 8 },  // H列
        { width: 8 }   // I列
    ];
    
    // 第一行
    const row1 = worksheet.addRow(['杭州佳沃思医疗科技有限公司']);
    worksheet.mergeCells('A1:G1');
    row1.font = { size: 22, bold: true };
    row1.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // 第二行：隐适美加工对账单（对账月）
    const row2 = worksheet.addRow([`隐适美加工对账单（${billingMonth}）`]);
    worksheet.mergeCells('A2:G2');
    row2.font = { size: 14 };
    row2.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // 第三行：客户名称
    const row3 = worksheet.addRow(['客户名称：' + clinicName]);
    worksheet.mergeCells('A3:G3');
    row3.font = { size: 11 };
    
    // 为前四行添加边框
    for (let i = 1; i <= 3; i++) {
        for (let j = 1; j <= 7; j++) {
            const cell = worksheet.getCell(i, j);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
    }
    
    // 表头
    const headerRow = worksheet.addRow(['序号', '发货日期', '病历号', '医生', '患者', '类型', '价格']);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    
    // 为表头添加边框
    for (let j = 1; j <= 7; j++) {
        const cell = headerRow.getCell(j);
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }
    
    // 填充数据
    let priceSum = 0;
    data.forEach((row, index) => {
        // 查找各列数据
        let dateValue = '-';
        for (const key in row) {
            if (key.includes('CCA时间')) {
                dateValue = row[key] || '-';
                
                // 处理日期格式 - 如果是数字，可能是Excel日期序列号
                if (typeof dateValue === 'number') {
                    try {
                        // Excel日期序列号转换为JavaScript日期
                        const excelEpoch = new Date(1900, 0, 1);
                        const jsDate = new Date(excelEpoch);
                        jsDate.setDate(excelEpoch.getDate() + dateValue - 2); // -2是Excel日期系统的修正
                        
                        // 格式化为YYYY-MM-DD
                        const year = jsDate.getFullYear();
                        const month = String(jsDate.getMonth() + 1).padStart(2, '0');
                        const day = String(jsDate.getDate()).padStart(2, '0');
                        dateValue = `${year}-${month}-${day}`;
                    } catch (e) {
                        // 如果转换失败，保留原始值
                        console.log("日期转换失败:", e);
                    }
                }
                
                break;
            }
        }
        
        let pidValue = '-';
        for (const key in row) {
            if (key.includes('PID') || key.includes('病历') || key.includes('编号')) {
                pidValue = row[key] || '-';
                break;
            }
        }
        
        let doctorValue = '-';
        for (const key in row) {
            if (key.includes('医生') || key.includes('Doctor')) {
                doctorValue = row[key] || '-';
                break;
            }
        }
        
        let patientValue = '-';
        for (const key in row) {
            if (key.includes('患者') || key.includes('Patient')) {
                patientValue = row[key] || '-';
                break;
            }
        }
        
        let typeValue = '-';
        for (const key in row) {
            if (key.includes('类型') || key.includes('规格型号')) {
                typeValue = row[key] || '-';
                break;
            }
        }
        
        let priceValue = '-';
        for (const key in row) {
            if (key.includes('价格') || key.includes('和门诊结算价')) {
                priceValue = row[key] || '-';
                if (typeof priceValue === 'number') {
                    priceSum += priceValue;
                } else if (typeof priceValue === 'string') {
                    const numValue = parseFloat(priceValue);
                    if (!isNaN(numValue)) {
                        priceSum += numValue;
                    }
                }
                break;
            }
        }
        
        // 添加数据行
        const dataRow = worksheet.addRow([
            index + 1, 
            dateValue, 
            pidValue, 
            doctorValue, 
            patientValue, 
            typeValue, 
            priceValue
        ]);
        
        // 设置对齐方式
        dataRow.getCell(1).alignment = { horizontal: 'center' };
        dataRow.getCell(2).alignment = { horizontal: 'center' };
        dataRow.getCell(3).alignment = { horizontal: 'center' };
        dataRow.getCell(4).alignment = { horizontal: 'center' };
        dataRow.getCell(5).alignment = { horizontal: 'center' };
        dataRow.getCell(6).alignment = { horizontal: 'center' };
        dataRow.getCell(7).alignment = { horizontal: 'right' };
        
        // 添加边框
        for (let j = 1; j <= 7; j++) {
            const cell = dataRow.getCell(j);
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
    });
    
    // 合计行
    const sumRow = worksheet.addRow(['', '', '', '', '', '<合计>', priceSum]);
    sumRow.getCell(6).alignment = { horizontal: 'center' };
    sumRow.getCell(7).alignment = { horizontal: 'right' };
    sumRow.getCell(7).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    
    // 佳沃思对公付款账户信息（一行显示）
    const accountRow = worksheet.addRow(['佳沃思对公付款账户:\n开户名称: 杭州佳沃思医疗科技有限公司\n开户行: 中信银行钱江支行\n银行账户: 8110801013500719720\n联行号: 302331033178']);
    worksheet.mergeCells(`A${accountRow.number}:G${accountRow.number}`);
    accountRow.height = 80; // 增加行高到80
    accountRow.alignment = { wrapText: true, vertical: 'middle' }; // 设置文本自动换行和垂直居中
    
    // 注释行
    const noteRow = worksheet.addRow(['注:你好，收到对账单后请及时核对账单及开票信息，如有异议联系我们；如无异于请于__20日_前回款，谢谢！']);
    worksheet.mergeCells(`A${noteRow.number}:H${noteRow.number}`);
    
    // 生成Excel文件
    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// 不再需要 formatExcelData 函数，因为我们直接在 generateExcelFile 中创建格式化的 Excel

// 显示处理结果
function displayResults() {
    resultSection.style.display = 'block';
    resultSummary.textContent = `成功处理了 ${processedFiles.length} 个诊所的数据`;
    
    // 清空表格
    resultTable.innerHTML = '';
    
    // 添加每个诊所的结果行
    processedFiles.forEach((file, index) => {
        const row = document.createElement('tr');
        
        // 序号
        const indexCell = document.createElement('td');
        indexCell.textContent = index + 1;
        row.appendChild(indexCell);
        
        // 诊所名称
        const nameCell = document.createElement('td');
        nameCell.textContent = file.name;
        row.appendChild(nameCell);
        
        // 操作按钮
        const actionCell = document.createElement('td');
        const downloadButton = document.createElement('button');
        downloadButton.className = 'btn btn-sm btn-outline-primary';
        downloadButton.textContent = '下载';
        downloadButton.addEventListener('click', () => {
            downloadSingleFile(file);
        });
        actionCell.appendChild(downloadButton);
        row.appendChild(actionCell);
        
        resultTable.appendChild(row);
    });
}

// 下载单个文件
function downloadSingleFile(file) {
    saveAs(file.blob, file.fileName);
}

// 下载所有文件
async function downloadAllFiles() {
    if (processedFiles.length === 0) {
        showAlert('没有可下载的文件');
        return;
    }
    
    try {
        log('正在生成ZIP文件...');
        
        // 生成ZIP文件
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        
        // 下载ZIP文件
        saveAs(zipBlob, '诊所Excel文件.zip');
        
        log('ZIP文件已生成并开始下载');
    } catch (error) {
        log(`生成ZIP文件时出错: ${error.message}`);
        showAlert(`生成ZIP文件时出错: ${error.message}`);
    }
}