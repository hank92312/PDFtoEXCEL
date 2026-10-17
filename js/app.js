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
    detectedMonths: [], // 所有解析出的月份清單，例如 ['2026-08-01', '2026-09-01']
    monthlyData: {}, // key 為 ym，值為 { mergedEmployees, laborInsurance, laborPension, healthInsurance }
    activeMonth: '2026-08-01', // 目前檢視/編輯中的月份
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
    multiMonthBadge: document.getElementById('multiMonthBadge'),
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

    // 月份切換事件
    dom.monthSelect.addEventListener('change', (e) => {
      switchActiveMonth(e.target.value);
      showToast(`已切換檢視月份至：${e.target.value.substring(0, 7)}`, 'info');
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
   * 執行 PDF 解析管線 (支援多月份批次聚合)
   */
  async function runParsingPipeline() {
    showToast('正在解析 PDF 清冊內容、保費年月與眷屬關係...', 'info');

    try {
      const results = await PdfParser.parseAllFiles(state.uploadedPdfFiles);
      state.parseResults = results;
      state.detectedMonths = results.detectedMonths || [];
      state.monthlyData = {};

      // 針對每個偵測到的月份分別彙整員工名冊
      state.detectedMonths.forEach(ym => {
        const mObj = results.byMonth[ym];
        const merged = mergeEmployeeDataForMonth(mObj);
        state.monthlyData[ym] = {
          yearMonth: ym,
          label: mObj.label,
          laborInsurance: mObj.laborInsurance,
          laborPension: mObj.laborPension,
          healthInsurance: mObj.healthInsurance,
          mergedEmployees: merged
        };
      });

      // 檢查是否有多個月份，動態更新 UI 提示與按鈕
      if (state.detectedMonths.length > 1) {
        const monthNames = state.detectedMonths.map(m => m.substring(5, 7) + '月').join('、');
        dom.multiMonthBadge.style.display = 'inline-flex';
        dom.multiMonthBadge.innerHTML = `<span class="pulse-dot"></span> 🌟 批次模式：已自動識別 ${state.detectedMonths.length} 個月份 (${monthNames})，匯出時將一次全數寫入！`;
        dom.btnExportCostExcel.innerHTML = `確認無誤，批次寫入 (${state.detectedMonths.length} 個月份) 並下載 EXCEL`;
      } else {
        dom.multiMonthBadge.style.display = 'none';
        dom.btnExportCostExcel.innerHTML = '確認無誤，寫入並下載 EXCEL';
      }

      // 預設切換至第一個偵測到的月份
      state.activeMonth = state.detectedMonths[0] || '2026-08-01';
      switchActiveMonth(state.activeMonth);

      dom.previewPanel.style.display = 'block';

      if (state.detectedMonths.length > 1) {
        showToast(`🎉 成功識別 ${state.detectedMonths.length} 個月份共 ${state.uploadedPdfFiles.length} 份檔案！請切換月份核對`, 'success');
      } else {
        showToast('PDF 清冊解析完成！請核對下方勾稽與名冊', 'success');
      }
    } catch (err) {
      console.error('解析失敗:', err);
      showToast(`解析失敗: ${err.message}`, 'error');
    }
  }

  /**
   * 切換當前正在檢視與編輯的月份
   */
  function switchActiveMonth(ym) {
    state.activeMonth = ym;
    dom.monthSelect.value = ym;
    renderAuditCardsForMonth(ym);
    renderEmployeeTableForMonth(ym);
  }

  /**
   * 針對單一月份整合多份報表的員工明細
   */
  function mergeEmployeeDataForMonth(monthObj) {
    const { laborInsurance, laborPension, healthInsurance } = monthObj;
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
      const foundDefault = state.hrConfig.defaultEmployees.find(d => d.name === emp.name);
      if (foundDefault) {
        emp.code = foundDefault.code;
      } else {
        emp.code = `E${String(idx + 1).padStart(2, '0')}`;
      }
    });

    return merged;
  }

  /**
   * 渲染指定月份的官方勾稽指標卡片
   */
  function renderAuditCardsForMonth(ym) {
    const monthData = state.monthlyData[ym] || (state.parseResults && state.parseResults.byMonth && state.parseResults.byMonth[ym]) || null;
    const labor = monthData ? monthData.laborInsurance : null;
    const pension = monthData ? monthData.laborPension : null;
    const health = monthData ? monthData.healthInsurance : null;

    // 勞保
    if (labor) {
      dom.laborCompanyTotal.textContent = `$${labor.totalCompany.toLocaleString()}`;
      dom.laborPersonalTotal.textContent = `個人負擔加總: $${labor.totalPersonal.toLocaleString()}`;
    } else {
      dom.laborCompanyTotal.textContent = '此月未上傳';
      dom.laborPersonalTotal.textContent = '-';
    }

    // 勞退
    if (pension) {
      dom.pensionEmployerTotal.textContent = `$${pension.totalEmployer.toLocaleString()}`;
      dom.pensionExpectedTotal.textContent = `官方應繳總額: $${pension.expectedTotal.toLocaleString()}`;
      if (pension.isBalanced) {
        dom.pensionAuditTag.className = 'audit-tag tag-ok';
        dom.pensionAuditTag.textContent = '✅ 勾稽一致';
      } else {
        dom.pensionAuditTag.className = 'audit-tag tag-warn';
        dom.pensionAuditTag.textContent = '⚠️ 總額不符';
      }
    } else {
      dom.pensionEmployerTotal.textContent = '此月未上傳';
      dom.pensionExpectedTotal.textContent = '-';
      dom.pensionAuditTag.textContent = '未上傳';
    }

    // 健保
    if (health) {
      dom.healthGrandTotal.textContent = `$${health.totalCompany.toLocaleString()}`;
      dom.healthExpectedTotal.textContent = `個人負擔加總: $${health.totalPersonal.toLocaleString()} | 官方應繳總額: $${health.expectedTotal.toLocaleString()}`;
      if (health.isBalanced) {
        dom.healthAuditTag.className = 'audit-tag tag-ok';
        dom.healthAuditTag.textContent = '✅ 勾稽一致';
      } else {
        dom.healthAuditTag.className = 'audit-tag tag-warn';
        dom.healthAuditTag.textContent = '⚠️ 總額不符';
      }
    } else {
      dom.healthGrandTotal.textContent = '此月未上傳';
      dom.healthExpectedTotal.textContent = '-';
      dom.healthAuditTag.textContent = '未上傳';
    }
  }

  /**
   * 渲染指定月份的員工明細表格
   */
  function renderEmployeeTableForMonth(ym) {
    dom.employeeTableBody.innerHTML = '';
    const mData = state.monthlyData[ym];
    const employees = mData ? mData.mergedEmployees : [];

    if (employees.length === 0) {
      const emptyTr = document.createElement('tr');
      emptyTr.innerHTML = `<td colspan="9" style="text-align: center; color: var(--text-muted); padding: 30px;">該月份 (${ym.substring(0, 7)}) 尚未上傳 PDF 清冊資料</td>`;
      dom.employeeTableBody.appendChild(emptyTr);
      return;
    }

    employees.forEach((emp, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${emp.name}</strong></td>
        <td>
          <select class="table-select" data-ym="${ym}" data-idx="${idx}" data-field="code">
            ${state.hrConfig.defaultEmployees.map(d => `
              <option value="${d.code}" ${emp.code === d.code ? 'selected' : ''}>${d.code} (${d.name})</option>
            `).join('')}
          </select>
        </td>
        <td><input type="number" class="table-num-input" data-ym="${ym}" data-idx="${idx}" data-field="healthPersonal" value="${emp.healthPersonal}" /></td>
        <td><input type="number" class="table-num-input" data-ym="${ym}" data-idx="${idx}" data-field="healthCompany" value="${emp.healthCompany}" /></td>
        <td><input type="number" class="table-num-input" data-ym="${ym}" data-idx="${idx}" data-field="laborPersonal" value="${emp.laborPersonal}" /></td>
        <td><input type="number" class="table-num-input" data-ym="${ym}" data-idx="${idx}" data-field="laborCompany" value="${emp.laborCompany}" /></td>
        <td><input type="number" class="table-num-input" data-ym="${ym}" data-idx="${idx}" data-field="pensionEmployer" value="${emp.pensionEmployer}" /></td>
        <td><span style="font-size: 0.8rem; color: var(--text-muted);">${emp.dependentsInfo || '-'}</span></td>
        <td style="text-align: center; white-space: nowrap;"><span class="audit-tag tag-ok" style="position: static; font-size: 0.76rem; white-space: nowrap; display: inline-block;">已核對</span></td>
      `;

      // 監聽數值微調事件 (支援多月份獨立更新)
      tr.querySelectorAll('input, select').forEach(input => {
        input.addEventListener('change', (e) => {
          const field = e.target.dataset.field;
          const index = parseInt(e.target.dataset.idx, 10);
          const targetYm = e.target.dataset.ym;
          if (state.monthlyData[targetYm] && state.monthlyData[targetYm].mergedEmployees[index]) {
            state.monthlyData[targetYm].mergedEmployees[index][field] = field === 'code' ? e.target.value : Number(e.target.value);
          }
        });
      });

      dom.employeeTableBody.appendChild(tr);
    });
  }

  /**
   * 寫入並匯出 cost.xlsx (支援多月份批次一次性注入)
   */
  async function exportCostExcel() {
    if (!state.parseResults) {
      showToast('請先上傳 PDF 清冊進行解析', 'error');
      return;
    }

    const monthsToExport = (state.detectedMonths && state.detectedMonths.length > 0)
      ? state.detectedMonths
      : [state.activeMonth];

    showToast(`正在將 ${monthsToExport.length} 個月份安全寫入年度 EXCEL 表單...`, 'info');

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

      // 依序寫入所有偵測到的月份
      for (const ym of monthsToExport) {
        const mData = state.monthlyData[ym];
        if (!mData || !mData.mergedEmployees) continue;

        // 準備員工名冊對照 (依照 E01~E10 排序)
        const empMappings = state.hrConfig.defaultEmployees.map(d => {
          const found = mData.mergedEmployees.find(m => m.code === d.code);
          return {
            code: d.code,
            name: found ? found.name : d.name
          };
        });

        // 整理各項險種寫入紀錄
        const laborRecords = mData.mergedEmployees.map(e => ({
          name: e.name,
          personal: e.laborPersonal || 0,
          company: e.laborCompany || 0
        }));

        const pensionRecords = mData.mergedEmployees.map(e => ({
          name: e.name,
          employerAmount: e.pensionEmployer || 0,
          personalAmount: 0
        }));

        const healthRecords = mData.mergedEmployees.map(e => ({
          name: e.name,
          personal: e.healthPersonal || 0,
          company: e.healthCompany || 0
        }));

        // 執行該月份的原位安全注入
        await ExcelWriter.updateCostWorkbook(zip, {
          targetMonth: ym,
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
            totalCompany: healthRecords.reduce((s, r) => s + r.company, 0),
            expectedTotal: healthRecords.reduce((s, r) => s + r.personal + r.company, 0),
            records: healthRecords
          },
          hrConfig: state.hrConfig
        });
      }

      // 智慧產出檔名：單月份與多月份批次智慧命名
      let prefix = 'cost';
      if (state.customCostExcelName) {
        prefix = state.customCostExcelName.replace(/\.xlsx$/i, '')
          .replace(/_\d{4}[_-]\d{2}(?:_至_\d{2})?$/, '')
          .replace(/_\d{4}\d{2}$/, '')
          .replace(/_\d{2}$/, '');
      }

      let outName;
      if (monthsToExport.length === 1) {
        const yyyymm = monthsToExport[0].substring(0, 7).replace('-', '_');
        outName = `${prefix}_${yyyymm}.xlsx`;
      } else {
        const firstYm = monthsToExport[0].substring(0, 7).replace('-', '_');
        const lastMonthStr = monthsToExport[monthsToExport.length - 1].substring(5, 7);
        outName = `${prefix}_${firstYm}_至_${lastMonthStr}.xlsx`;
      }

      await ExcelWriter.downloadZip(zip, outName);

      if (monthsToExport.length > 1) {
        showToast(`🎉 成功批次更新！已將 ${monthsToExport.length} 個月份寫入並下載：${outName}`, 'success');
      } else {
        showToast(`🎉 成功更新！已產出並下載：${outName}`, 'success');
      }
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
    state.detectedMonths = [];
    state.monthlyData = {};
    dom.previewPanel.style.display = 'none';
    dom.multiMonthBadge.style.display = 'none';
    dom.btnExportCostExcel.innerHTML = '確認無誤，寫入並下載 EXCEL';
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
