/**
 * app.js - 人事成本智慧計算與轉換工具主控制器
 */

document.addEventListener('DOMContentLoaded', () => {
  // 核心狀態
  const state = {
    currentMode: 'hr', // 'hr' | 'po'
    uploadedPdfFiles: [], // Array<{ name: string, arrayBuffer: ArrayBuffer }>
    customCostExcelBuffer: null, // ArrayBuffer or null
    customCostExcelName: null,
    parseResults: null,
    mergedEmployees: [],
    targetMonth: '2026-08-01',
    hrConfig: window.ConfigManager.getHRConfig()
  };

  // DOM 元素快取
  const dom = {
    btnOpenSettings: document.getElementById('btnOpenSettings'),

    // 上傳與按鈕
    dropzone: document.getElementById('dropzone'),
    pdfFileInput: document.getElementById('pdfFileInput'),
    excelFileInput: document.getElementById('excelFileInput'),
    btnBrowsePdf: document.getElementById('btnBrowsePdf'),
    btnBrowseExcel: document.getElementById('btnBrowseExcel'),
    fileListContainer: document.getElementById('fileListContainer'),

    // 核對面板
    previewPanel: document.getElementById('previewPanel'),
    monthSelect: document.getElementById('monthSelect'),
    templateInfo: document.getElementById('templateInfo'),
    btnDownloadTemplate: document.getElementById('btnDownloadTemplate'),
    btnReset: document.getElementById('btnReset'),
    btnExportCostExcel: document.getElementById('btnExportCostExcel'),

    // 勾稽卡片
    laborCompanyTotal: document.getElementById('laborCompanyTotal'),
    laborPersonalTotal: document.getElementById('laborPersonalTotal'),
    pensionEmployerTotal: document.getElementById('pensionEmployerTotal'),
    pensionExpectedTotal: document.getElementById('pensionExpectedTotal'),
    pensionAuditTag: document.getElementById('pensionAuditTag'),
    healthGrandTotal: document.getElementById('healthGrandTotal'),
    healthExpectedTotal: document.getElementById('healthExpectedTotal'),
    healthAuditTag: document.getElementById('healthAuditTag'),

    // 員工表格
    employeeTableBody: document.getElementById('employeeTableBody'),

    // 設定 Modal
    settingsModal: document.getElementById('settingsModal'),
    btnCloseSettings: document.getElementById('btnCloseSettings'),
    btnSaveSettings: document.getElementById('btnSaveSettings'),
    btnResetSettings: document.getElementById('btnResetSettings'),
    cfgHealthSheet: document.getElementById('cfgHealthSheet'),
    cfgLaborSheet: document.getElementById('cfgLaborSheet'),
    cfgPensionSheet: document.getElementById('cfgPensionSheet'),
    cfgEmpSheet: document.getElementById('cfgEmpSheet'),
    cfgDetailStartRow: document.getElementById('cfgDetailStartRow'),
    cfgDetailEndRow: document.getElementById('cfgDetailEndRow'),

    // Toast
    toastContainer: document.getElementById('toastContainer')
  };

  // 初始化
  initEventHandlers();
  loadSettingsToUI();

  /**
   * 註冊所有事件監聽器
   */
  function initEventHandlers() {
    // 檔案選取按鈕
    dom.btnBrowsePdf.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.pdfFileInput.click();
    });

    dom.btnBrowseExcel.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.excelFileInput.click();
    });

    dom.pdfFileInput.addEventListener('change', handlePdfFileSelect);
    dom.excelFileInput.addEventListener('change', handleExcelFileSelect);

    // 拖曳區事件
    dom.dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dom.dropzone.classList.add('drag-over');
    });

    dom.dropzone.addEventListener('dragleave', () => {
      dom.dropzone.classList.remove('drag-over');
    });

    dom.dropzone.addEventListener('drop', handleFileDrop);

    // 月份切換
    dom.monthSelect.addEventListener('change', (e) => {
      state.targetMonth = e.target.value;
      showToast(`已切換目標寫入月份至：${e.target.value.substring(0, 7)}`, 'info');
    });

    // 重設清空
    dom.btnReset.addEventListener('click', resetAll);

    // 下載空白範本
    dom.btnDownloadTemplate.addEventListener('click', () => {
      const a = document.createElement('a');
      a.href = './cost.xlsx';
      a.download = 'cost_範本.xlsx';
      a.click();
      showToast('開始下載空白年度 cost.xlsx 範本', 'success');
    });

    // 匯出人事成本 Excel
    dom.btnExportCostExcel.addEventListener('click', exportCostExcel);

    // 彈性設定 Modal
    dom.btnOpenSettings.addEventListener('click', openSettingsModal);
    dom.btnCloseSettings.addEventListener('click', closeSettingsModal);
    dom.btnSaveSettings.addEventListener('click', saveSettingsFromUI);
    dom.btnResetSettings.addEventListener('click', resetSettingsUI);
    dom.settingsModal.addEventListener('click', (e) => {
      if (e.target === dom.settingsModal) closeSettingsModal();
    });
  }

  /**
   * 處理 PDF 檔案選取 (可多選)
   */
  async function handlePdfFileSelect(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    await processIncomingFiles(files);
    dom.pdfFileInput.value = '';
  }

  /**
   * 處理 Excel 檔案選取 (支援任意中英文檔名)
   */
  async function handleExcelFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    await loadExcelFile(file);
    dom.excelFileInput.value = '';
  }

  /**
   * 載入並驗證 Excel 檔案
   */
  async function loadExcelFile(file) {
    try {
      const buffer = await file.arrayBuffer();
      // 驗證是否為可解析的 Excel 檔案
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      
      const sheetNames = wb.worksheets.map(w => w.name);
      const hasRequiredSheet = sheetNames.some(name => 
        name.includes('健保') || name.includes('勞保') || name.includes('勞退') || name.includes('員工')
      );

      state.customCostExcelBuffer = buffer;
      state.customCostExcelName = file.name;
      updateFileBadgesUI();

      if (hasRequiredSheet) {
        dom.templateInfo.innerHTML = `<span class="pulse-indicator"></span> 使用範本：${file.name} (自訂年度表)`;
        showToast(`已成功載入年度報表：${file.name}`, 'success');
      } else {
        dom.templateInfo.innerHTML = `<span class="pulse-indicator" style="background:#f59e0b;box-shadow:0 0 8px #f59e0b;"></span> 使用範本：${file.name} (未偵測到標準工作表名)`;
        showToast(`⚠️ 已載入 ${file.name}，但請確認是否包含「健保/勞保/勞退」工作表`, 'warn');
      }
    } catch (err) {
      console.error('Excel 載入失敗:', err);
      showToast(`載入 Excel 失敗: ${err.message}`, 'error');
    }
  }

  /**
   * 處理拖放事件 (同時支援多個 PDF 與 Excel，不限中英文檔名)
   */
  async function handleFileDrop(e) {
    e.preventDefault();
    dom.dropzone.classList.remove('drag-over');

    const files = Array.from(e.dataTransfer.files);
    if (!files.length) return;
    await processIncomingFiles(files);
  }

  /**
   * 處理傳入的檔案清單
   */
  async function processIncomingFiles(files) {
    let hasPdf = false;

    for (const file of files) {
      const lowerName = file.name.toLowerCase();
      if (lowerName.endsWith('.xlsx')) {
        await loadExcelFile(file);
      } else if (lowerName.endsWith('.pdf')) {
        const buffer = await file.arrayBuffer();
        if (!state.uploadedPdfFiles.some(f => f.name === file.name)) {
          state.uploadedPdfFiles.push({ name: file.name, arrayBuffer: buffer });
          hasPdf = true;
        }
      }
    }

    updateFileBadgesUI();

    if (hasPdf || state.uploadedPdfFiles.length > 0) {
      await runParsingPipeline();
    }
  }

  /**
   * 更新檔案徽章清單 UI (包含取消刪除按鈕)
   */
  function updateFileBadgesUI() {
    dom.fileListContainer.innerHTML = '';
    const items = [];

    // Excel
    if (state.customCostExcelBuffer) {
      items.push(`
        <div class="file-item-badge badge-excel">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
          </svg>
          <span>${state.customCostExcelName} (年度管理表)</span>
          <button type="button" class="btn-remove-badge" data-type="excel" title="取消此 Excel 檔案">&times;</button>
        </div>
      `);
    }

    // PDFs
    state.uploadedPdfFiles.forEach((f, idx) => {
      items.push(`
        <div class="file-item-badge badge-ok">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          </svg>
          <span>${f.name}</span>
          <button type="button" class="btn-remove-badge" data-type="pdf" data-idx="${idx}" title="取消此 PDF 檔案">&times;</button>
        </div>
      `);
    });

    if (items.length > 0) {
      dom.fileListContainer.innerHTML = items.join('');
      dom.fileListContainer.style.display = 'flex';

      // 綁定 X 刪除按鈕事件
      dom.fileListContainer.querySelectorAll('.btn-remove-badge').forEach(btn => {
        btn.addEventListener('click', async (e) => {
          e.stopPropagation();
          const type = btn.dataset.type;
          if (type === 'excel') {
            state.customCostExcelBuffer = null;
            state.customCostExcelName = null;
            dom.templateInfo.innerHTML = '<span class="pulse-indicator"></span> 使用範本：cost.xlsx (預設內建範本)';
            showToast('已取消選取自訂 Excel 表單，切換回預設範本', 'info');
          } else if (type === 'pdf') {
            const idx = parseInt(btn.dataset.idx, 10);
            const removed = state.uploadedPdfFiles.splice(idx, 1);
            showToast(`已移除檔案：${removed[0]?.name || ''}`, 'info');
          }
          updateFileBadgesUI();
          if (state.uploadedPdfFiles.length > 0) {
            await runParsingPipeline();
          } else {
            state.parseResults = null;
            state.mergedEmployees = [];
            dom.previewPanel.style.display = 'none';
          }
        });
      });
    } else {
      dom.fileListContainer.style.display = 'none';
    }
  }

  /**
   * 執行 PDF 解析管線
   */
  async function runParsingPipeline() {
    showToast('正在解析 PDF 清冊內容與眷屬關係...', 'info');

    try {
      const results = await PdfParser.parseAllFiles(state.uploadedPdfFiles);
      state.parseResults = results;

      // 自動設定目標月份
      if (results.detectedYearMonth) {
        state.targetMonth = results.detectedYearMonth;
        dom.monthSelect.value = results.detectedYearMonth;
      }

      // 彙整員工資料
      mergeEmployeeData();

      // 渲染勾稽卡片與表格
      renderAuditCards();
      renderEmployeeTable();

      dom.previewPanel.style.display = 'block';
      showToast('PDF 清冊解析完成！請核對下方勾稽與名冊', 'success');
    } catch (err) {
      console.error('解析失敗:', err);
      showToast(`解析失敗: ${err.message}`, 'error');
    }
  }

  /**
   * 整合多份報表的員工明細
   */
  function mergeEmployeeData() {
    const { laborInsurance, laborPension, healthInsurance } = state.parseResults;
    const empMap = new Map();

    // 1. 勞保
    if (laborInsurance && laborInsurance.records) {
      laborInsurance.records.forEach(r => {
        if (!empMap.has(r.name)) {
          empMap.set(r.name, { name: r.name, idNumber: r.idNumber || '', healthPersonal: 0, healthCompany: 0, laborPersonal: 0, laborCompany: 0, pensionEmployer: 0, dependentsInfo: '' });
        }
        const emp = empMap.get(r.name);
        emp.laborPersonal = r.personal;
        emp.laborCompany = r.company;
      });
    }

    // 2. 勞退
    if (laborPension && laborPension.records) {
      laborPension.records.forEach(r => {
        if (!empMap.has(r.name)) {
          empMap.set(r.name, { name: r.name, idNumber: '', healthPersonal: 0, healthCompany: 0, laborPersonal: 0, laborCompany: 0, pensionEmployer: 0, dependentsInfo: '' });
        }
        const emp = empMap.get(r.name);
        emp.pensionEmployer = r.employerAmount;
      });
    }

    // 3. 健保 (含眷屬歸戶)
    if (healthInsurance && healthInsurance.records) {
      healthInsurance.records.forEach(r => {
        if (!empMap.has(r.name)) {
          empMap.set(r.name, { name: r.name, idNumber: r.idNumber || '', healthPersonal: 0, healthCompany: 0, laborPersonal: 0, laborCompany: 0, pensionEmployer: 0, dependentsInfo: '' });
        }
        const emp = empMap.get(r.name);
        emp.healthPersonal = r.personal;
        emp.healthCompany = r.company;
        if (r.dependents && r.dependents.length > 0) {
          emp.dependentsInfo = `含 ${r.dependents.length} 眷屬 (${r.dependents.map(d => d.name).join(', ')})`;
        }
      });
    }

    // 對照預設代號 E01~E10
    const merged = Array.from(empMap.values());
    merged.forEach((emp, idx) => {
      // 找尋預設清冊中是否有同名
      const foundDefault = state.hrConfig.defaultEmployees.find(d => d.name === emp.name);
      if (foundDefault) {
        emp.code = foundDefault.code;
      } else {
        emp.code = `E${String(idx + 1).padStart(2, '0')}`;
      }
    });

    state.mergedEmployees = merged;
  }

  /**
   * 渲染官方勾稽指標卡片
   */
  function renderAuditCards() {
    const { laborInsurance, laborPension, healthInsurance } = state.parseResults;

    // 勞保
    if (laborInsurance) {
      dom.laborCompanyTotal.textContent = `$${laborInsurance.totalCompany.toLocaleString()}`;
      dom.laborPersonalTotal.textContent = `個人負擔加總: $${laborInsurance.totalPersonal.toLocaleString()}`;
    } else {
      dom.laborCompanyTotal.textContent = '未上傳';
      dom.laborPersonalTotal.textContent = '-';
    }

    // 勞退
    if (laborPension) {
      dom.pensionEmployerTotal.textContent = `$${laborPension.totalEmployer.toLocaleString()}`;
      dom.pensionExpectedTotal.textContent = `官方應繳總額: $${laborPension.expectedTotal.toLocaleString()}`;
      if (laborPension.isBalanced) {
        dom.pensionAuditTag.className = 'audit-tag tag-ok';
        dom.pensionAuditTag.textContent = '✅ 勾稽一致';
      } else {
        dom.pensionAuditTag.className = 'audit-tag tag-warn';
        dom.pensionAuditTag.textContent = '⚠️ 總額不符';
      }
    } else {
      dom.pensionEmployerTotal.textContent = '未上傳';
      dom.pensionExpectedTotal.textContent = '-';
      dom.pensionAuditTag.textContent = '未上傳';
    }

    // 健保
    if (healthInsurance) {
      dom.healthGrandTotal.textContent = `$${healthInsurance.actualTotal.toLocaleString()}`;
      dom.healthExpectedTotal.textContent = `官方應繳總額: $${healthInsurance.expectedTotal.toLocaleString()} (單位: $${healthInsurance.totalCompany.toLocaleString()})`;
      if (healthInsurance.isBalanced) {
        dom.healthAuditTag.className = 'audit-tag tag-ok';
        dom.healthAuditTag.textContent = '✅ 勾稽一致';
      } else {
        dom.healthAuditTag.className = 'audit-tag tag-warn';
        dom.healthAuditTag.textContent = '⚠️ 總額不符';
      }
    } else {
      dom.healthGrandTotal.textContent = '未上傳';
      dom.healthExpectedTotal.textContent = '-';
      dom.healthAuditTag.textContent = '未上傳';
    }
  }

  /**
   * 渲染員工資料明細表格
   */
  function renderEmployeeTable() {
    dom.employeeTableBody.innerHTML = '';

    state.mergedEmployees.forEach((emp, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${emp.name}</strong></td>
        <td>
          <select class="table-select" data-idx="${idx}" data-field="code">
            ${state.hrConfig.defaultEmployees.map(d => `
              <option value="${d.code}" ${emp.code === d.code ? 'selected' : ''}>${d.code} (${d.name})</option>
            `).join('')}
          </select>
        </td>
        <td><input type="number" class="table-num-input" data-idx="${idx}" data-field="healthPersonal" value="${emp.healthPersonal}" /></td>
        <td><input type="number" class="table-num-input" data-idx="${idx}" data-field="healthCompany" value="${emp.healthCompany}" /></td>
        <td><input type="number" class="table-num-input" data-idx="${idx}" data-field="laborPersonal" value="${emp.laborPersonal}" /></td>
        <td><input type="number" class="table-num-input" data-idx="${idx}" data-field="laborCompany" value="${emp.laborCompany}" /></td>
        <td><input type="number" class="table-num-input" data-idx="${idx}" data-field="pensionEmployer" value="${emp.pensionEmployer}" /></td>
        <td><span style="font-size: 0.8rem; color: var(--text-muted);">${emp.dependentsInfo || '-'}</span></td>
        <td style="text-align: center; white-space: nowrap;"><span class="audit-tag tag-ok" style="position: static; font-size: 0.76rem; white-space: nowrap; display: inline-block;">已核對</span></td>
      `;

      // 監聽數值微調事件
      tr.querySelectorAll('input, select').forEach(input => {
        input.addEventListener('change', (e) => {
          const field = e.target.dataset.field;
          const index = parseInt(e.target.dataset.idx, 10);
          state.mergedEmployees[index][field] = field === 'code' ? e.target.value : Number(e.target.value);
        });
      });

      dom.employeeTableBody.appendChild(tr);
    });
  }

  /**
   * 寫入並匯出 cost.xlsx (基於原位 XML 安全注入，徹底避免損毀)
   */
  async function exportCostExcel() {
    if (!state.parseResults) {
      showToast('請先上傳 PDF 清冊進行解析', 'error');
      return;
    }

    showToast('正在將資料安全寫入年度 EXCEL 表單...', 'info');

    try {
      // 載入 ZIP 容器 (優先使用使用者上傳的，否則載入 ./cost.xlsx 預設範本)
      const zipSource = state.customCostExcelBuffer || './cost.xlsx';
      let zip;
      try {
        zip = await ExcelWriter.loadZip(zipSource);
      } catch (loadErr) {
        if (!state.customCostExcelBuffer && window.location.protocol === 'file:') {
          throw new Error('偵測到您直接以檔案路徑開啟 (file://)，瀏覽器安全機制限制自動讀取本地預設範本。請點選「選取現有 EXCEL 表單」選入 cost.xlsx，或使用本機伺服器開啟。');
        }
        throw loadErr;
      }

      // 準備員工名冊對照 (依照 E01~E10 排序)
      const empMappings = state.hrConfig.defaultEmployees.map(d => {
        const found = state.mergedEmployees.find(m => m.code === d.code);
        return {
          code: d.code,
          name: found ? found.name : d.name
        };
      });

      // 整理各項險種寫入紀錄
      const laborRecords = state.mergedEmployees.map(e => ({
        name: e.name,
        personal: e.laborPersonal,
        company: e.laborCompany
      }));

      const pensionRecords = state.mergedEmployees.map(e => ({
        name: e.name,
        employerAmount: e.pensionEmployer,
        personalAmount: 0
      }));

      const healthRecords = state.mergedEmployees.map(e => ({
        name: e.name,
        personal: e.healthPersonal,
        company: e.healthCompany
      }));

      // 執行非破壞性安全更新
      await ExcelWriter.updateCostWorkbook(zip, {
        targetMonth: state.targetMonth,
        employeeMappings: empMappings,
        laborData: {
          totalCompany: laborRecords.reduce((s, r) => s + r.company, 0),
          records: laborRecords
        },
        pensionData: {
          expectedTotal: pensionRecords.reduce((s, r) => s + r.employerAmount, 0),
          records: pensionRecords
        },
        healthData: {
          expectedTotal: healthRecords.reduce((s, r) => s + r.personal + r.company, 0),
          records: healthRecords
        },
        hrConfig: state.hrConfig
      });

      // 智慧產出檔名：若使用者有上傳自訂 Excel，沿用其原檔名作為前綴
      const yyyymm = state.targetMonth.substring(0, 7).replace('-', '_');
      let outName = `cost_${yyyymm}.xlsx`;
      if (state.customCostExcelName) {
        const baseName = state.customCostExcelName.replace(/\.xlsx$/i, '');
        outName = `${baseName}_${yyyymm}.xlsx`;
      }

      await ExcelWriter.downloadZip(zip, outName);

      showToast(`🎉 成功更新！已產出並下載：${outName}`, 'success');
    } catch (err) {
      console.error('匯出失敗:', err);
      showToast(`匯出失敗: ${err.message}`, 'error');
    }
  }


  /**
   * 重設清空
   */
  function resetAll() {
    state.uploadedPdfFiles = [];
    state.customCostExcelBuffer = null;
    state.customCostExcelName = null;
    state.parseResults = null;
    state.mergedEmployees = [];
    dom.previewPanel.style.display = 'none';
    dom.templateInfo.innerHTML = '<span class="pulse-indicator"></span> 使用範本：cost.xlsx (預設內建範本)';
    updateFileBadgesUI();
    showToast('已重設所有檔案與資料', 'info');
  }

  /**
   * 彈性設定 Modal 相關
   */
  function openSettingsModal() {
    loadSettingsToUI();
    dom.settingsModal.style.display = 'flex';
  }

  function closeSettingsModal() {
    dom.settingsModal.style.display = 'none';
  }

  function loadSettingsToUI() {
    const s = state.hrConfig.sheets;
    dom.cfgHealthSheet.value = s.healthInsurance.sheetName;
    dom.cfgLaborSheet.value = s.laborInsurance.sheetName;
    dom.cfgPensionSheet.value = s.laborPension.sheetName;
    dom.cfgEmpSheet.value = s.employeeSettings.sheetName;
    dom.cfgDetailStartRow.value = s.healthInsurance.detailStartRow;
    dom.cfgDetailEndRow.value = s.healthInsurance.detailEndRow;
  }

  function saveSettingsFromUI() {
    state.hrConfig.sheets.healthInsurance.sheetName = dom.cfgHealthSheet.value;
    state.hrConfig.sheets.laborInsurance.sheetName = dom.cfgLaborSheet.value;
    state.hrConfig.sheets.laborPension.sheetName = dom.cfgPensionSheet.value;
    state.hrConfig.sheets.employeeSettings.sheetName = dom.cfgEmpSheet.value;
    const startRow = parseInt(dom.cfgDetailStartRow.value, 10);
    const endRow = parseInt(dom.cfgDetailEndRow.value, 10);

    state.hrConfig.sheets.healthInsurance.detailStartRow = startRow;
    state.hrConfig.sheets.healthInsurance.detailEndRow = endRow;
    state.hrConfig.sheets.laborInsurance.detailStartRow = startRow;
    state.hrConfig.sheets.laborInsurance.detailEndRow = endRow;
    state.hrConfig.sheets.laborPension.detailStartRow = startRow;
    state.hrConfig.sheets.laborPension.detailEndRow = endRow;

    window.ConfigManager.saveHRConfig(state.hrConfig);
    closeSettingsModal();
    showToast('設定已成功儲存至瀏覽器！', 'success');
  }

  function resetSettingsUI() {
    state.hrConfig = window.ConfigManager.resetHRConfig();
    loadSettingsToUI();
    showToast('設定已還原為標準預設值', 'info');
  }

  /**
   * Toast 訊息提示
   */
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    dom.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }
});
