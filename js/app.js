/**
 * 主應用邏輯控制
 * 處理檔案拖放、按鈕選擇、解析排程、人機協同核對表單及 Excel 產生
 */

// 應用狀態
const state = {
  currentFile: null,
  currentConfigId: 'po_standard',
  extractedFields: [],
  fullText: ''
};

// DOM 元素快取
const elements = {
  dropzone: document.getElementById('dropzone'),
  fileInput: document.getElementById('fileInput'),
  btnBrowse: document.getElementById('btnBrowse'),
  btnLoadDemo: document.getElementById('btnLoadDemo'),
  fileBadge: document.getElementById('fileBadge'),
  fileNameSpan: document.getElementById('fileName'),
  
  previewPanel: document.getElementById('previewPanel'),
  configSelect: document.getElementById('configSelect'),
  fieldsContainer: document.getElementById('fieldsContainer'),
  btnExportExcel: document.getElementById('btnExportExcel'),
  btnReset: document.getElementById('btnReset'),
  templateInfo: document.getElementById('templateInfo'),
  toastContainer: document.getElementById('toastContainer')
};

// 初始化應用
function initApp() {
  populateConfigSelect();
  bindEvents();
}

// 填充表單範本選單
function populateConfigSelect() {
  elements.configSelect.innerHTML = '';
  for (const [id, config] of Object.entries(PDF_CONFIGS)) {
    const option = document.createElement('option');
    option.value = id;
    option.textContent = config.name;
    elements.configSelect.appendChild(option);
  }
}

// 事件綁定
function bindEvents() {
  // 1. 點擊「選擇檔案」按鈕觸發原生 input file
  elements.btnBrowse.addEventListener('click', (e) => {
    e.stopPropagation();
    elements.fileInput.click();
  });

  // 點擊拖曳框空白區亦可觸發選擇
  elements.dropzone.addEventListener('click', () => {
    elements.fileInput.click();
  });

  // 2. 原生檔案選擇變更
  elements.fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleSelectedFile(file);
  });

  // 3. 拖放 (Drag & Drop) 支援
  elements.dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.dropzone.classList.add('dragover');
  });

  elements.dropzone.addEventListener('dragleave', () => {
    elements.dropzone.classList.remove('dragover');
  });

  elements.dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
        handleSelectedFile(file);
      } else {
        showToast('請上傳 PDF 格式的檔案', 'error');
      }
    }
  });

  // 4. 一鍵載入示範 Demo 檔案
  elements.btnLoadDemo.addEventListener('click', async (e) => {
    e.stopPropagation();
    try {
      showToast('正在載入標準採購單示範檔案...', 'info');
      const response = await fetch('./samples/sample_po.pdf');
      if (!response.ok) throw new Error('找不到示範 PDF 檔案');
      const blob = await response.blob();
      const demoFile = new File([blob], 'sample_po.pdf', { type: 'application/pdf' });
      handleSelectedFile(demoFile);
    } catch (err) {
      console.error(err);
      showToast('載入示範檔案失敗：' + err.message, 'error');
    }
  });

  // 5. 手動切換表單模版選單（防呆覆寫）
  elements.configSelect.addEventListener('change', (e) => {
    state.currentConfigId = e.target.value;
    if (state.fullText) {
      reparseFieldsWithConfig(state.currentConfigId);
    }
  });

  // 6. 匯出 Excel 按鈕
  elements.btnExportExcel.addEventListener('click', handleExportExcel);

  // 7. 重新上傳按鈕
  elements.btnReset.addEventListener('click', resetApp);
}

// 處理選取的 PDF 檔案
async function handleSelectedFile(file) {
  state.currentFile = file;
  elements.fileNameSpan.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  elements.fileBadge.style.display = 'inline-flex';

  try {
    showToast('正在解析 PDF 全文內容...', 'info');
    
    // 呼叫 PDFParser 抽取文字
    const { fullText } = await PDFParser.extractText(file);
    state.fullText = fullText;

    // 自動比對指紋辨識表單格式
    const detectedId = PDFParser.detectConfigId(fullText, PDF_CONFIGS);
    if (detectedId) {
      state.currentConfigId = detectedId;
      elements.configSelect.value = detectedId;
      showToast(`已自動辨識表單類型：${PDF_CONFIGS[detectedId].name}`, 'success');
    } else {
      showToast('未命中特定指紋，使用預設標準模版', 'info');
    }

    // 依據選定的模版抽取各欄位
    reparseFieldsWithConfig(state.currentConfigId);

    // 展開核對預覽介面
    elements.previewPanel.classList.add('active');
    elements.previewPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (err) {
    console.error('PDF 解析失敗:', err);
    showToast('PDF 解析失敗：' + err.message, 'error');
  }
}

// 根據指定 Config 重新抽取欄位並渲染核對表單
function reparseFieldsWithConfig(configId) {
  const config = PDF_CONFIGS[configId];
  if (!config) return;

  state.extractedFields = PDFParser.extractFields(state.fullText, config);
  renderFields(state.extractedFields, config);
  updateTemplateInfo(config);
}

// 動態渲染核對與編輯欄位 UI
function renderFields(fields, config) {
  elements.fieldsContainer.innerHTML = '';

  fields.forEach((field, index) => {
    const card = document.createElement('div');
    card.className = `field-card ${field.isValid ? 'status-valid' : 'status-invalid'}`;

    card.innerHTML = `
      <div class="field-header">
        <label class="field-label" for="field_${field.key}">
          ${field.label} ${field.required ? '<span style="color:var(--danger)">*</span>' : ''}
        </label>
        <span class="field-target-cell" title="將寫入 Excel 儲存格">${field.targetCell}</span>
      </div>
      <input 
        type="text" 
        id="field_${field.key}" 
        class="field-input" 
        value="${escapeHtml(field.value)}" 
        placeholder="${escapeHtml(field.placeholder || '請輸入內容')}"
        data-index="${index}"
      />
      <div class="field-helper">
        ${field.value ? '✅ 已自動擷取' : (field.required ? '⚠️ 必填欄位，請手動補填' : '（非必填）')}
      </div>
    `;

    // 監聽使用者手動即時修改
    const input = card.querySelector('input');
    input.addEventListener('input', (e) => {
      const newVal = e.target.value.trim();
      state.extractedFields[index].value = newVal;
      
      const isValid = !(field.required && !newVal);
      state.extractedFields[index].isValid = isValid;

      if (isValid) {
        card.classList.add('status-valid');
        card.classList.remove('status-invalid');
        card.querySelector('.field-helper').innerHTML = '✏️ 已手動修改';
      } else {
        card.classList.remove('status-valid');
        card.classList.add('status-invalid');
        card.querySelector('.field-helper').innerHTML = '⚠️ 必填欄位不可為空';
      }
    });

    elements.fieldsContainer.appendChild(card);
  });
}

function updateTemplateInfo(config) {
  elements.templateInfo.innerHTML = `
    目前套用範本：<strong>${config.name}</strong> (${config.templateUrl})
  `;
}

// 執行產生與下載 Excel
async function handleExportExcel() {
  const config = PDF_CONFIGS[state.currentConfigId];
  if (!config) {
    showToast('找不到對應的表單設定檔', 'error');
    return;
  }

  // 驗證必填欄位
  const missingRequired = state.extractedFields.filter(f => f.required && !String(f.value).trim());
  if (missingRequired.length > 0) {
    const missingNames = missingRequired.map(f => f.label).join('、');
    showToast(`請填寫必填欄位：${missingNames}`, 'error');
    return;
  }

  try {
    elements.btnExportExcel.disabled = true;
    elements.btnExportExcel.textContent = '正在產生 Excel...';

    // 準備檔名
    let filename = '報表產出.xlsx';
    if (config.outputFilenamePattern) {
      const poNumField = state.extractedFields.find(f => f.key === 'poNumber');
      const poNumber = (poNumField && poNumField.value) ? poNumField.value : 'PO';
      filename = config.outputFilenamePattern.replace('{poNumber}', poNumber);
    }

    // 呼叫 ExcelWriter 載入範本並填值
    await ExcelWriter.generateAndDownload(config.templateUrl, state.extractedFields, filename);

    showToast(`✅ Excel 報表 [${filename}] 已成功產出並下載！`, 'success');
  } catch (err) {
    console.error('匯出失敗:', err);
    showToast('匯出失敗：' + err.message, 'error');
  } finally {
    elements.btnExportExcel.disabled = false;
    elements.btnExportExcel.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      確認無誤，產出並下載 Excel
    `;
  }
}

// 重設介面
function resetApp() {
  state.currentFile = null;
  state.extractedFields = [];
  state.fullText = '';
  elements.fileInput.value = '';
  elements.fileBadge.style.display = 'none';
  elements.previewPanel.classList.remove('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
  showToast('已重設，請重新選擇檔案', 'info');
}

// 簡易 Toast 提示訊息
function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  
  let icon = 'ℹ️';
  if (type === 'success') icon = '✅';
  if (type === 'error') icon = '❌';

  toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
  elements.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// DOM 加載完成後啟動
document.addEventListener('DOMContentLoaded', initApp);
