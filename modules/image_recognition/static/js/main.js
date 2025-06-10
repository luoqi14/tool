document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const imagePreviewsContainer = document.getElementById('imagePreviews'); // For direct uploads, if any
    const resultContainer = document.getElementById('resultContainer');
    const processButton = document.getElementById('processButton');
    const promptInput = document.getElementById('promptInput');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const statusIndicator = document.getElementById('status-indicator');
    const productIdInput = document.getElementById('productIdInput');
    const downloadResultsButton = document.getElementById('downloadResults');

    // --- 样式 --- (保持不变)
    function addCustomStyles() {
        if (!document.getElementById('custom-ocr-styles')) {
            const style = document.createElement('style');
            style.id = 'custom-ocr-styles';
            style.textContent = `
                .image-display-section { margin-bottom: 5px; padding: 5px; border: 1px solid #e0e0e0; border-radius: 5px; background-color: #f9f9f9; }
                .image-display-section h6 { margin-top: 0; margin-bottom: 5px; color: #333; font-weight: bold; font-size: 14px; }
                .image-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(100px, 1fr)); gap: 5px; margin-top: 5px; }
                .image-result-item { border: 1px solid #ddd; border-radius: 4px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); transition: transform 0.2s; }
                .image-result-item:hover { transform: scale(1.03); }
                .image-result-item img { width: 100%; height: 100px; object-fit: cover; }
                .recognition-result-section { padding: 5px; border: 1px solid #e0e0e0; border-radius: 5px; background-color: #fff; max-height: 400px; overflow-y: auto; }
                .recognition-result-section h5 { margin-top: 0; color: #333; font-weight: bold; font-size: 16px; }
                .result-content { margin-top: 5px; line-height: 1.4; white-space: pre-wrap; font-size: 14px; }
                .product-result-section { margin-bottom: 5px; border: 1px solid #ccc; border-radius: 8px; background-color: #f0f0f0; overflow: hidden; }
                .product-result-section h5 { font-size: 15px; margin: 0; padding: 8px 10px; cursor: pointer; background-color: #e9ecef; }
                .product-result-section .collapsible-content { max-height: 1500px; overflow: hidden; transition: max-height 0.4s ease-in-out, padding 0.4s ease-in-out; padding: 5px 10px; }
                .product-result-section.collapsed .collapsible-content { max-height: 0; padding-top: 0; padding-bottom: 0; }
                .batch-log { max-height: 150px; overflow-y: auto; border: 1px solid #ddd; padding: 8px; margin-bottom: 10px; background-color: #f8f9fa; }
                .batch-log-item { font-size: 13px !important; } /* Ensure px and override Bootstrap fs-sm if it uses rem */
                .status-indicator .badge { font-size: 13px; padding: 5px 8px; }
            `;
            document.head.appendChild(style);
        }
    }
    addCustomStyles();

    // --- 状态变量 --- 
    let productsData = []; // Batch: [{ id, image_urls:[], image_files:[], resultText:'', status:'', error:null }]
    let currentProduct = { // Single: { id, image_urls:[], image_files:[], resultText:'', status:'', error:null }
        id: '', image_urls: [], image_files: [], resultText: '', status: '', error: null
    };
    let currentBatchProductIndex = 0;
    let isBatchProcessingMode = false;

    // --- API URLs ---
    const API_URL_FETCH_IMAGES = '/image/recognition/fetch_product_images';
    const API_URL_PROCESS_IMAGES = '/image/recognition/process';

    // --- 默认提示词 --- (保持不变)
    promptInput.value = `请识别图片中的所有文字，并按以下顺序生成三份输出，要求严格按照以下格式：

### RAW_START ###
(在这两行标记之间，放 OCR 识别到的完整原始文本)
### RAW_END ###

### GENERIC_JSON_START ###
(在这两行标记之间，放一个"通用"格式的 JSON 对象，键名可根据识别到的内容自定义，所有识别到的项都要包含)
### GENERIC_JSON_END ###

### STRUCTURED_JSON_START ###
(在这两行标记之间，放如下模板的 JSON 对象，字段顺序与模板一致，未识别到的值填 null，禁止多余字段或注释)
{
  "productName": null,             // 商品名称
  "brandName": null,               // 品牌名称
  "orderingInfo": null,            // 规格信息
  "validPeriod": null,             // 有效期
  "sterilizationValidity": null,   // 灭菌有效期
  "guaranteePeriod": null,         // 质保期
  "storage": null,                 // 储存方式
  "useLife": null,                 // 设备使用寿命
  "effectiveComponent": null,      // 有效成分及含量
  "instruction": null,             // 使用说明
  "clinicalOperation": null,       // 临床操作
  "mountingsDesc": null,           // 配件说明
  "maintaining": null,             // 维护/保养
  "productDesc": null,             // 商品介绍
  "sterilization": null            // 消毒灭菌
}
### STRUCTURED_JSON_END ###`;

    // --- UI辅助函数 --- 
    function createProductResultSection(productId) {
        const section = document.createElement('div');
        section.className = 'product-result-section collapsed';
        section.id = `product-result-${productId}`;
        section.innerHTML = `
            <h5 title="点击展开/收起" style="display: flex; justify-content: space-between; align-items: center;">
                <span>产品 ${productId} 处理结果</span>
                <button class="btn btn-sm btn-outline-secondary re-recognize-btn" data-product-id="${productId}" title="重新识别" style="padding: 0.1rem 0.3rem; font-size: 0.75rem; line-height: 1;"><i class="bi bi-arrow-clockwise"></i></button>
            </h5>
            <div class="collapsible-content">
                <div class="image-display-section mb-2">
                    <h6>产品图片：</h6>
                    <div class="image-grid"><p class="text-muted">等待加载图片...</p></div>
                </div>
                <div class="recognition-result-section">
                    <div class="result-content"><p class="text-muted">等待识别...</p></div>
                </div>
            </div>
        `;

        section.querySelector('h5').addEventListener('click', () => {
            section.classList.toggle('collapsed');
        });

        return section;
    }

    function updateStatusIndicator(statusKey, productId = null, message = '') {
        let fullMessage = '';
        switch (statusKey) {
            case 'idle': fullMessage = '<span class="badge bg-secondary">尚未开始</span>'; break;
            case 'fetching_urls': fullMessage = `<span class="badge bg-info">产品 ${productId}: 获取图片URL...</span>`; break;
            case 'converting_images': fullMessage = `<span class="badge bg-info">产品 ${productId}: 转换图片...</span>`; break;
            case 'processing_ocr': fullMessage = `<span class="badge bg-primary">产品 ${productId}: 识别中...${message}</span>`; break;
            case 'success_ocr': fullMessage = `<span class="badge bg-success">产品 ${productId}: 识别成功</span>`; break;
            case 'error': fullMessage = `<span class="badge bg-danger">产品 ${productId}: 错误 - ${message}</span>`; break;
            case 'batch_starting': fullMessage = `<span class="badge bg-primary">批量处理已启动...</span>`; break;
            case 'batch_item_processing': fullMessage = `<span class="badge bg-info">批量处理: 产品 ${productId} (${currentBatchProductIndex + 1}/${productsData.length})</span>`; break;
            case 'batch_complete': fullMessage = `<span class="badge bg-success">批量处理完成</span>`; break;
            default: fullMessage = `<span class="badge bg-secondary">${statusKey}</span>`;
        }
        statusIndicator.innerHTML = fullMessage;
        loadingIndicator.classList.toggle('d-none', !['fetching_urls', 'converting_images', 'processing_ocr', 'batch_starting', 'batch_item_processing'].includes(statusKey));
    }

    function addBatchLog(logMessage, type = 'info') {
        const batchLogContainer = resultContainer.querySelector('.batch-log');
        if (!batchLogContainer) return;
        const logItem = document.createElement('div');
        logItem.className = `batch-log-item alert alert-${type} p-2 mb-1 fs-sm`; // fs-sm for smaller font
        logItem.innerHTML = logMessage;
        batchLogContainer.appendChild(logItem);
        batchLogContainer.scrollTop = batchLogContainer.scrollHeight;
    }

    function displayProductImages(imageUrls, productId) {
        const productSection = document.getElementById(`product-result-${productId}`);
        if (!productSection) return;
        const imageGrid = productSection.querySelector('.image-display-section .image-grid');
        if (!imageGrid) return;

        imageGrid.innerHTML = ''; // Clear previous/loading message
        if (!imageUrls || imageUrls.length === 0) {
            imageGrid.innerHTML = '<p class="text-danger">未能加载到任何图片。</p>';
            return;
        }
        imageUrls.forEach((url, index) => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'image-result-item';
            const img = document.createElement('img');
            img.src = url;
            img.alt = `产品 ${productId} 图片 ${index + 1}`;
            img.classList.add('clickable-image'); // Added class for clickability
            img.dataset.imageUrl = url;          // Store URL for preview
            img.style.cursor = 'pointer';        // Visual cue for clickability
            img.onerror = () => { itemDiv.innerHTML = '<p class="text-danger sm">图片加载失败</p>'; };
            itemDiv.appendChild(img);
            imageGrid.appendChild(itemDiv);
        });
    }

    function updateProductRecognitionResult(productId, text, isStreaming = false) {
        const productSection = document.getElementById(`product-result-${productId}`);
        if (!productSection) return;
        const resultContentDiv = productSection.querySelector('.recognition-result-section .result-content');
        if (!resultContentDiv) return;
        resultContentDiv.innerHTML = formatResult(text); // formatResult handles markdown
    }
    
    function formatResult(text) {
        if (window.marked && typeof window.marked.parse === 'function') {
            return window.marked.parse(text || '');
        }
        return text ? text.replace(/\n/g, '<br>') : ''; // Basic fallback
    }

    function parseResultText(resultText) {
        if (!resultText) {
            return { raw: '', genericJson: '', structuredJson: '' };
        }
        const rawMatch = resultText.match(/### RAW_START ###\s*([\s\S]*?)\s*### RAW_END ###/);
        const genericJsonMatch = resultText.match(/### GENERIC_JSON_START ###\s*([\s\S]*?)\s*### GENERIC_JSON_END ###/);
        const structuredJsonMatch = resultText.match(/### STRUCTURED_JSON_START ###\s*([\s\S]*?)\s*### STRUCTURED_JSON_END ###/);

        return {
            raw: rawMatch ? rawMatch[1].trim() : '',
            genericJson: genericJsonMatch ? genericJsonMatch[1].trim() : '',
            structuredJson: structuredJsonMatch ? structuredJsonMatch[1].trim() : ''
        };
    }

    function expandAndScrollToProductPanel(productId) {
        const allSections = document.querySelectorAll('.product-result-section');
        allSections.forEach(sec => {
            const isTarget = sec.id === `product-result-${productId}`;
            sec.classList.toggle('collapsed', !isTarget);
        });

        const targetPanel = document.getElementById(`product-result-${productId}`);
        if (targetPanel) {
            targetPanel.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }

    // --- Core Logic --- 
    async function convertUrlToImageFile(imageUrl, filenamePrefix = 'image_') {
        try {
            const response = await fetch(imageUrl);
            if (!response.ok) throw new Error(`HTTP ${response.status} for ${imageUrl}`);
            const blob = await response.blob();
            let filename = imageUrl.substring(imageUrl.lastIndexOf('/') + 1).split(/[?#]/)[0]; // Get filename from URL
            if (!filename || !filename.includes('.')) { // Ensure filename has extension
                const extension = blob.type.split('/')[1] || 'jpg';
                filename = `${filenamePrefix}${Date.now()}.${extension}`;
            }
            return new File([blob], filename, { type: blob.type });
        } catch (error) {
            console.error(`Failed to convert URL to File (${imageUrl}):`, error);
            return null; // Indicate failure
        }
    }

    async function fetchAndPrepareImages(productId) {
        updateStatusIndicator('fetching_urls', productId);
        let targetProduct = isBatchProcessingMode ? productsData.find(p => p.id === productId) : currentProduct;
        if (!targetProduct) return null;

        targetProduct.image_urls = [];
        targetProduct.image_files = [];

        try {
            const response = await fetch(`${API_URL_FETCH_IMAGES}?product_id=${productId}`);
            if (!response.ok) throw new Error(`API error: ${response.statusText}`);
            const data = await response.json();

            if (!data.success || !data.images || data.images.length === 0) {
                throw new Error(data.message || '未找到产品图片。');
            }
            targetProduct.image_urls = data.images;
            displayProductImages(targetProduct.image_urls, productId);
            updateStatusIndicator('converting_images', productId);

            for (const url of targetProduct.image_urls) {
                const file = await convertUrlToImageFile(url, `${productId}_`);
                if (file) {
                    targetProduct.image_files.push(file);
                }
            }

            if (targetProduct.image_files.length === 0 && targetProduct.image_urls.length > 0) {
                 throw new Error('所有图片URL都无法转换为文件。');
            }
            return targetProduct.image_files;
        } catch (error) {
            console.error(`Error fetching/preparing images for ${productId}:`, error);
            targetProduct.error = error.message;
            targetProduct.status = 'error';
            updateStatusIndicator('error', productId, error.message);
            updateProductRecognitionResult(productId, `<p class="text-danger">图片处理失败: ${error.message}</p>`);
            return null;
        }
    }

    async function triggerImageRecognition(imageFiles, promptValue, productId) {
        if (!imageFiles || imageFiles.length === 0) {
            updateStatusIndicator('error', productId, '没有可识别的图片文件。');
            updateProductRecognitionResult(productId, '<p class="text-danger">错误: 没有可识别的图片文件。</p>');
            if (isBatchProcessingMode) {
                const product = productsData.find(p => p.id === productId);
                if (product) product.status = 'error';
            }
            return;
        }

        updateStatusIndicator('processing_ocr', productId);
        let targetProduct = isBatchProcessingMode ? productsData.find(p => p.id === productId) : currentProduct;
        if (targetProduct) {
            targetProduct.status = 'processing';
            targetProduct.resultText = ''; // Clear previous results
        }

        const formData = new FormData();
        formData.append('prompt', promptValue);
        imageFiles.forEach(file => formData.append('files', file));

        try {
            const response = await fetch(API_URL_PROCESS_IMAGES, { method: 'POST', body: formData });
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({ message: `HTTP ${response.status}` }));
                throw new Error(errorData.message || `API请求失败: ${response.statusText}`);
            }

            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '';

            while (true) {
                const { value, done } = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, { stream: true });
                let boundary = buffer.indexOf('\n\n');
                while (boundary !== -1) {
                    const chunk = buffer.substring(0, boundary);
                    buffer = buffer.substring(boundary + 2);
                    if (chunk.startsWith('data: ')) {
                        try {
                            const eventData = JSON.parse(chunk.substring(6));
                            if (eventData.text) {
                                if (targetProduct) targetProduct.resultText += eventData.text;
                                updateProductRecognitionResult(productId, targetProduct ? targetProduct.resultText : eventData.text, true);
                                updateStatusIndicator('processing_ocr', productId, '接收数据...');
                            } else if (eventData.done) {
                                if (targetProduct) targetProduct.status = 'success';
                                updateStatusIndicator('success_ocr', productId);
                                if (!isBatchProcessingMode) downloadResultsButton.style.display = 'inline-block';
                                return; // Stream finished successfully
                            }
                        } catch (e) { console.error('Error parsing stream data:', e, 'Chunk:', chunk); }
                    }
                    boundary = buffer.indexOf('\n\n');
                }
            }
            // Should be caught by eventData.done, but as a fallback:
            if (targetProduct) targetProduct.status = 'success';
            updateStatusIndicator('success_ocr', productId);
            if (!isBatchProcessingMode) downloadResultsButton.style.display = 'inline-block';

        } catch (error) {
            console.error(`Error during recognition for ${productId}:`, error);
            if (targetProduct) {
                targetProduct.error = error.message;
                targetProduct.status = 'error';
            }
            updateStatusIndicator('error', productId, error.message);
            updateProductRecognitionResult(productId, `<p class="text-danger">识别失败: ${error.message}</p>`);
        }
    }

    async function processSingleProduct(productId, promptValue) {
        isBatchProcessingMode = false;
        currentProduct.id = productId;
        currentProduct.status = 'pending';
        currentProduct.error = null;
        currentProduct.resultText = '';
        resultContainer.innerHTML = ''; // Clear global results for single product
        resultContainer.appendChild(createProductResultSection(productId));
        downloadResultsButton.style.display = 'none';

        const imageFiles = await fetchAndPrepareImages(productId);
        if (imageFiles && imageFiles.length > 0) {
            await triggerImageRecognition(imageFiles, promptValue, productId);
        } else if (!currentProduct.error) { // If no specific error from fetchAndPrepareImages but no files
            updateStatusIndicator('error', productId, '未能准备图片进行识别。');
        }
    }

    async function processBatch() {
        if (currentBatchProductIndex >= productsData.length) {
            updateStatusIndicator('batch_complete');
            addBatchLog('所有产品处理完成。', 'success');
            downloadResultsButton.style.display = 'inline-block';
            return;
        }

        const product = productsData[currentBatchProductIndex];
        expandAndScrollToProductPanel(product.id);
        updateStatusIndicator('batch_item_processing', product.id);
        addBatchLog(`开始处理产品: ${product.id} (${currentBatchProductIndex + 1}/${productsData.length})`, 'info');
        product.status = 'pending';
        product.error = null;
        product.resultText = '';

        const imageFiles = await fetchAndPrepareImages(product.id);
        if (imageFiles && imageFiles.length > 0) {
            await triggerImageRecognition(imageFiles, promptInput.value.trim(), product.id);
        } else if (!product.error) {
            updateStatusIndicator('error', product.id, '未能准备图片进行识别。');
            addBatchLog(`产品 ${product.id}: 未能准备图片。`, 'warning');
            product.status = 'error'; // Mark as error if no files and no prior error
        }
        
        // Log completion or error for the current item
        if (product.status === 'success') {
            addBatchLog(`产品 ${product.id} 处理成功。`, 'success');
        } else if (product.status === 'error') {
            addBatchLog(`产品 ${product.id} 处理失败: ${product.error || '未知错误'}`, 'danger');
        }

        currentBatchProductIndex++;
        setTimeout(processBatch, 100); // Process next item with a small delay
    }

    // --- Feature Implementations ---
    function handleReRecognize(productId) {
        console.log(`Attempting to re-recognize product: ${productId}`);
        const product = productsData.find(p => p.id === productId);

        if (!product || !product.image_files || product.image_files.length === 0) {
            alert('没有找到该产品的图片文件缓存，无法重新识别。请确保该产品已成功加载过一次。');
            return;
        }

        const promptValue = promptInput.value.trim();
        if (!promptValue) {
            alert('请输入提示词后再进行识别。');
            promptInput.focus();
            return;
        }

        // Clear previous results and show loading state
        const productSection = document.getElementById(`product-result-${productId}`);
        if (productSection) {
            const resultContent = productSection.querySelector('.result-content');
            const resultHeader = productSection.querySelector('.recognition-result-section h5');
            if (resultContent && resultHeader) {
                resultHeader.textContent = '识别结果 (重新识别中...)';
                resultContent.innerHTML = '<div class="d-flex justify-content-center"><div class="spinner-border spinner-border-sm" role="status"><span class="visually-hidden">Loading...</span></div></div>';
            }
            // Reset product data for re-recognition
            product.resultText = '';
            product.isProcessing = true;
        }
        
        updateStatusIndicator('processing', `正在重新识别产品 ${productId}...`);
        expandAndScrollToProductPanel(productId);

        // Reuse the existing triggerImageRecognition function
        triggerImageRecognition(product.image_files, promptValue, productId);
    }

    function showImagePreview(imageUrl) {
        const modalImage = document.getElementById('modalImage');
        if (modalImage) {
            modalImage.src = imageUrl;
            const modal = new bootstrap.Modal(document.getElementById('imagePreviewModal'));
            modal.show();
        } else {
            console.error('Image preview modal element not found!');
            alert('图片预览功能初始化失败。');
        }
    }

    // --- Event Listeners ---
    processButton.addEventListener('click', async () => {
        const productIdsRaw = productIdInput.value.trim();
        const promptValue = promptInput.value.trim();

        if (!promptValue) {
            alert('请输入提示词。');
            return;
        }
        if (!productIdsRaw) {
            alert('请输入产品ID。'); // For now, require product ID
            return;
        }

        const uniqueProductIds = [...new Set(productIdsRaw.split(',').map(id => id.trim()).filter(id => id))];

        if (uniqueProductIds.length === 0) {
            alert('未提供有效的产品ID。');
            return;
        }

        resultContainer.innerHTML = ''; // Clear previous results
        downloadResultsButton.style.display = 'none';

        if (uniqueProductIds.length > 1) {
            isBatchProcessingMode = true;
            productsData = uniqueProductIds.map(id => ({ 
                id: id, image_urls: [], image_files: [], resultText: '', status: 'pending', error: null 
            }));
            currentBatchProductIndex = 0;
            
            const batchLogContainer = document.createElement('div');
            batchLogContainer.className = 'batch-log card card-body mb-3';
            resultContainer.appendChild(batchLogContainer);
            addBatchLog(`启动批量处理，共 ${productsData.length} 个产品...`, 'primary');
            updateStatusIndicator('batch_starting');

            productsData.forEach(p => resultContainer.appendChild(createProductResultSection(p.id)));
            processBatch();
        } else {
            await processSingleProduct(uniqueProductIds[0], promptValue);
        }
    });

    downloadResultsButton.addEventListener('click', () => {
        if (isBatchProcessingMode && productsData.some(p => p.resultText)) {
            const zip = new JSZip();
            productsData.forEach(product => {
                if (product.resultText) {
                    const parsedResults = parseResultText(product.resultText);
                    const productFolder = zip.folder(product.id);
                    const imagesFolder = productFolder.folder('images');

                    let infoContent = `产品ID: ${product.id}\n\n`;
                    infoContent += `--- 包含图片文件 (位于 'images' 文件夹下) ---\n`;
                    if (product.image_files && product.image_files.length > 0) {
                        product.image_files.forEach(file => {
                            imagesFolder.file(file.name, file);
                            infoContent += `${file.name}\n`;
                        });
                    } else {
                        infoContent += `(无图片文件处理或包含在此次下载中)\n`;
                    }
                    productFolder.file(`info.txt`, infoContent);

                    if (parsedResults.raw) {
                        productFolder.file(`raw_ocr.txt`, parsedResults.raw);
                    }
                    if (parsedResults.genericJson) {
                        productFolder.file(`generic_output.json`, parsedResults.genericJson);
                    }
                    if (parsedResults.structuredJson) {
                        try {
                            const structuredData = JSON.parse(parsedResults.structuredJson);
                            structuredData.mmProductCode = product.id;
                            const updatedJson = JSON.stringify(structuredData, null, 2);
                            productFolder.file(`structured_output.json`, updatedJson);
                        } catch (e) {
                            console.error(`Error processing structured JSON for product ${product.id}:`, e);
                            productFolder.file(`structured_output.json`, parsedResults.structuredJson); // Fallback
                        }
                    }
                }
            });
            zip.generateAsync({ type: "blob" })
                .then(function(content) {
                    saveAs(content, "批量识别结果.zip");
                });
        } else if (!isBatchProcessingMode && currentProduct.id && currentProduct.resultText) {
            const zip = new JSZip();
            const parsedResults = parseResultText(currentProduct.resultText);
            const imagesFolder = zip.folder('images');

            let infoContent = `产品ID: ${currentProduct.id}\n\n`;
            infoContent += `--- 包含图片文件 (位于 'images' 文件夹下) ---\n`;
            if (currentProduct.image_files && currentProduct.image_files.length > 0) {
                currentProduct.image_files.forEach(file => {
                    imagesFolder.file(file.name, file);
                    infoContent += `${file.name}\n`;
                });
            } else {
                infoContent += `(无图片文件处理或包含在此次下载中)\n`;
            }
            zip.file(`info.txt`, infoContent);

            if (parsedResults.raw) {
                zip.file(`raw_ocr.txt`, parsedResults.raw);
            }
            if (parsedResults.genericJson) {
                zip.file(`generic_output.json`, parsedResults.genericJson);
            }
            if (parsedResults.structuredJson) {
                try {
                    const structuredData = JSON.parse(parsedResults.structuredJson);
                    structuredData.mmProductCode = currentProduct.id;
                    const updatedJson = JSON.stringify(structuredData, null, 2);
                    zip.file(`structured_output.json`, updatedJson);
                } catch (e) {
                    console.error(`Error processing structured JSON for product ${currentProduct.id}:`, e);
                    zip.file(`structured_output.json`, parsedResults.structuredJson); // Fallback
                }
            }

            zip.generateAsync({ type: "blob" })
                .then(function(blob) {
                    saveAs(blob, `${currentProduct.id}_识别结果.zip`);
                });
        } else {
            alert('没有可下载的结果。');
        }
    });

    // Delegated event listener for re-recognize and image preview
    resultContainer.addEventListener('click', function(event) {
        const target = event.target;
        // Check if the click is on a re-recognize button or its icon
        const reRecognizeButton = target.closest('.re-recognize-btn');
        if (reRecognizeButton) {
            const productId = reRecognizeButton.dataset.productId;
            handleReRecognize(productId);
            return; // Stop further processing if it's a re-recognize click
        }

        // Check if the click is on a clickable image
        if (target.classList.contains('clickable-image')) {
            const imageUrl = target.dataset.imageUrl;
            showImagePreview(imageUrl);
            return; // Stop further processing if it's an image click
        }
    });

    // Initial UI state
    updateStatusIndicator('idle');
});
