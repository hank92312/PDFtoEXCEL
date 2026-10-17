/**
 * excel-writer.js - 純客戶端 Excel 範本非破壞性資料注入與工作表動態管理
 * 採用 JSZip 直接進行 OpenXML 位元等級安全注入，徹底根除公式損毀問題
 * 支援 20 人彈性擴充、員工真實姓名動態工作表更名、全員工彙總自動連動
 */

class ExcelWriter {
  /**
   * 載入 Excel 範本為 JSZip 實例
   * @param {string | ArrayBuffer} source 範本檔案路徑或二進位陣列
   * @returns {Promise<JSZip>}
   */
  static async loadZip(source) {
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip 函式庫未載入，請檢查網路連線或 CDN 引用。');
    }

    let arrayBuffer = source;
    if (typeof source === 'string') {
      const response = await fetch(source);
      if (!response.ok) {
        throw new Error(`無法載入 Excel 範本 (${response.status}): ${source}`);
      }
      arrayBuffer = await response.arrayBuffer();
    }

    return await JSZip.loadAsync(arrayBuffer);
  }

  /**
   * 解析 workbook.xml 與 workbook.xml.rels，建立 工作表名稱 -> XML路徑 的動態映射
   * @param {JSZip} zip 
   * @returns {Promise<Map<string, string>>}
   */
  static async getSheetPathMap(zip) {
    const sheetMap = new Map();
    try {
      const wbXml = await zip.file('xl/workbook.xml').async('text');
      const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');

      // 1. 抓取所有 rId 與 Target
      const relsMap = new Map();
      const relRegex = /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
      let relMatch;
      while ((relMatch = relRegex.exec(relsXml)) !== null) {
        let target = relMatch[2];
        if (target.startsWith('/')) target = target.substring(1);
        if (!target.startsWith('xl/')) target = 'xl/' + target;
        relsMap.set(relMatch[1], target);
      }

      // 2. 抓取 sheet name 與 r:id
      const sheetRegex = /<sheet\s+[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g;
      let sheetMatch;
      while ((sheetMatch = sheetRegex.exec(wbXml)) !== null) {
        const name = sheetMatch[1];
        const rId = sheetMatch[2];
        if (relsMap.has(rId)) {
          sheetMap.set(name, relsMap.get(rId));
        }
      }
    } catch (e) {
      console.warn('動態解析工作表對應失敗，採用預設順序', e);
      sheetMap.set('使用說明', 'xl/worksheets/sheet1.xml');
      sheetMap.set('員工設定', 'xl/worksheets/sheet2.xml');
      sheetMap.set('健保', 'xl/worksheets/sheet3.xml');
      sheetMap.set('勞保', 'xl/worksheets/sheet4.xml');
      sheetMap.set('勞退', 'xl/worksheets/sheet5.xml');
      sheetMap.set('全員工彙總', 'xl/worksheets/sheet26.xml');
    }
    return sheetMap;
  }

  /**
   * 從既有 Excel 的「員工設定」工作表讀取目前已登記的員工名單 (以保留歷史員工)
   * @param {JSZip} zip
   * @param {Object} hrConfig
   * @returns {Promise<Array<{ code: string, name: string, isReal: boolean }>>}
   */
  static async readEmployeesFromExcel(zip, hrConfig = (window.DEFAULT_HR_CONFIG || {})) {
    const sheetPathMap = await this.getSheetPathMap(zip);
    const empSheetName = (hrConfig.sheets && hrConfig.sheets.employeeSettings && hrConfig.sheets.employeeSettings.sheetName) || '員工設定';
    const empSheetPath = sheetPathMap.get(empSheetName) || 'xl/worksheets/sheet2.xml';

    const result = [];
    const maxEmployees = 20;
    const startRow = (hrConfig.sheets && hrConfig.sheets.employeeSettings && hrConfig.sheets.employeeSettings.startRow) || 4;

    if (!zip.file(empSheetPath)) {
      return result;
    }

    try {
      const sheetXml = await zip.file(empSheetPath).async('text');
      for (let i = 0; i < maxEmployees; i++) {
        const row = startRow + i;
        const code = `E${String(i + 1).padStart(2, '0')}`;
        
        // 解析儲存格中的文字 (支援 inlineStr 或一般數值/文字格式)
        const cellRegex = new RegExp(`<c\\s+r="B${row}"[^>]*>(?:<is><t>([\\s\\S]*?)<\\/t><\\/is>|<v>([\\s\\S]*?)<\\/v>)?<\\/c>`);
        const match = sheetXml.match(cellRegex);
        const rawVal = match ? (match[1] || match[2] || '').trim() : '';

        // 判斷是否為真實員工姓名 (排除預設如 "員工01"、"員工1" 等)
        const isPlaceholder = !rawVal || /^(?:員工\s*\d+|[Ee]\d+)$/.test(rawVal);
        result.push({
          code,
          name: rawVal || `員工${String(i + 1).padStart(2, '0')}`,
          isReal: !isPlaceholder
        });
      }
    } catch (e) {
      console.warn('讀取既有員工設定失敗:', e);
    }

    return result;
  }

  /**
   * 淨化工作表名稱，符合 Excel 命名規範 (排除 \ / ? * : [ ]，長度 <= 31)
   */
  static sanitizeSheetName(name, fallbackName = '員工') {
    if (!name || typeof name !== 'string') return fallbackName;
    let clean = name.replace(/[\\/?*:[\]]/g, '_').trim();
    if (clean.length > 31) clean = clean.substring(0, 31);
    return clean || fallbackName;
  }

  /**
   * 動態將 workbook.xml 中的員工工作表更名為真實姓名
   * 排除系統工作表 (使用說明、員工設定、健保、勞保、勞退、全員工彙總)
   * @param {JSZip} zip 
   * @param {Array<{ code: string, name: string }>} employeeMappings 
   */
  static async updateWorkbookSheetNames(zip, employeeMappings) {
    if (!zip.file('xl/workbook.xml')) return;

    let wbXml = await zip.file('xl/workbook.xml').async('text');
    const systemSheets = new Set(['使用說明', '員工設定', '健保', '勞保', '勞退', '全員工彙總']);
    const usedNames = new Set(systemSheets);

    let empIndex = 0;

    wbXml = wbXml.replace(/(<sheet\s+[^>]*name=")([^"]+)("([^>]*>))/g, (full, p1, oldName, p3) => {
      // 若為系統保留工作表，絕對不改名
      if (systemSheets.has(oldName)) {
        return full;
      }

      // 對應第 empIndex 位員工
      const emp = employeeMappings[empIndex];
      empIndex++;

      let targetName = oldName;
      if (emp && emp.name && !/^(?:員工\s*\d+|[Ee]\d+)$/.test(emp.name)) {
        targetName = this.sanitizeSheetName(emp.name, oldName);
      }

      // 防止同名衝突
      let finalName = targetName;
      let count = 2;
      while (usedNames.has(finalName)) {
        finalName = `${targetName} (${count++})`;
      }
      usedNames.add(finalName);

      return `${p1}${finalName}${p3}`;
    });

    zip.file('xl/workbook.xml', wbXml);
  }

  /**
   * 在 XML 儲存格中安全寫入數值 (保持原始樣式 s，型別設為 t="n")
   */
  static setNumericCell(xml, cellRef, numVal) {
    const val = Number(numVal) || 0;
    const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*)(>)(.*?)(<\\/c>)|<c\\s+r="${cellRef}"([^>]*?)\\/>`);
    if (cellRegex.test(xml)) {
      return xml.replace(cellRegex, (match, openPrefix, openClose, inner, closeTag, selfCloseAttrs) => {
        if (openPrefix) {
          let prefix = openPrefix;
          if (!prefix.includes('t="n"')) {
            prefix = prefix.replace(/\s+t="[^"]*"/, '') + ' t="n"';
          }
          return `${prefix}${openClose}<v>${val}</v>${closeTag}`;
        } else {
          return `<c r="${cellRef}"${selfCloseAttrs} t="n"><v>${val}</v></c>`;
        }
      });
    }
    return xml;
  }

  /**
   * 在 XML 儲存格中安全寫入文字 (型別設為 t="inlineStr"，子標籤為 <is><t>...)
   */
  static setTextCell(xml, cellRef, textVal) {
    const escaped = String(textVal || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*)(>)(.*?)(<\\/c>)|<c\\s+r="${cellRef}"([^>]*?)\\/>`);
    if (cellRegex.test(xml)) {
      return xml.replace(cellRegex, (match, openPrefix, openClose, inner, closeTag, selfCloseAttrs) => {
        if (openPrefix) {
          let prefix = openPrefix.replace(/\s+t="[^"]*"/, '') + ' t="inlineStr"';
          return `${prefix}${openClose}<is><t>${escaped}</t></is>${closeTag}`;
        } else {
          return `<c r="${cellRef}"${selfCloseAttrs} t="inlineStr"><is><t>${escaped}</t></is></c>`;
        }
      });
    }
    return xml;
  }

  /**
   * 非破壞性更新 cost.xlsx 年度管理表 (以 JSZip 原位修改 XML)
   * 100% 保留所有個人分頁 SUMIFS 連動公式、全員工彙總表、樣式與格式
   */
  static async updateCostWorkbook(zip, options) {
    const {
      targetMonth = '2026-08-01',
      employeeMappings = [],
      laborData = null,
      pensionData = null,
      healthData = null,
      hrConfig = window.DEFAULT_HR_CONFIG
    } = options;

    const sheetPathMap = await this.getSheetPathMap(zip);

    // 解析目標月份 (1~12)
    const monthMatch = targetMonth.match(/\d{4}-(\d{2})/);
    const month = monthMatch ? parseInt(monthMatch[1], 10) : 8;

    // 1. 更新「員工設定」工作表姓名 (B4~B23，共 20 人)
    const empSheetName = hrConfig.sheets.employeeSettings.sheetName || '員工設定';
    const empSheetPath = sheetPathMap.get(empSheetName) || 'xl/worksheets/sheet2.xml';
    if (zip.file(empSheetPath) && employeeMappings.length > 0) {
      let empXml = await zip.file(empSheetPath).async('text');
      employeeMappings.forEach((emp, idx) => {
        const row = hrConfig.sheets.employeeSettings.startRow + idx;
        if (emp.name) {
          empXml = this.setTextCell(empXml, `B${row}`, emp.name);
        }
      });
      zip.file(empSheetPath, empXml);
    }

    // 計算目標月份在各保費工作表的行號 (每個月預留 20 人空間)
    const employeesPerMonth = 20;
    const summaryRow = hrConfig.sheets.laborInsurance.summaryStartRow + (month - 1);
    const detailMonthStartRow = hrConfig.sheets.laborInsurance.detailStartRow + (month - 1) * employeesPerMonth;

    // 2. 更新勞保 (sheet4.xml)
    const laborSheetName = hrConfig.sheets.laborInsurance.sheetName || '勞保';
    const laborSheetPath = sheetPathMap.get(laborSheetName) || 'xl/worksheets/sheet4.xml';
    if (zip.file(laborSheetPath) && laborData) {
      let laborXml = await zip.file(laborSheetPath).async('text');
      // 摘要
      laborXml = this.setNumericCell(laborXml, `B${summaryRow}`, laborData.totalCompany || 0);
      laborXml = this.setTextCell(laborXml, `D${summaryRow}`, '已核對');
      // 明細 20 人 (C:個人, D:公司)
      employeeMappings.forEach((emp, idx) => {
        const row = detailMonthStartRow + idx;
        const record = laborData.records.find(r => r.name === emp.name);
        const p = record ? (record.personal || 0) : 0;
        const c = record ? (record.company || 0) : 0;
        laborXml = this.setNumericCell(laborXml, `C${row}`, p);
        laborXml = this.setNumericCell(laborXml, `D${row}`, c);
      });
      zip.file(laborSheetPath, laborXml);
    }

    // 3. 更新勞退 (sheet5.xml)
    const pensionSheetName = hrConfig.sheets.laborPension.sheetName || '勞退';
    const pensionSheetPath = sheetPathMap.get(pensionSheetName) || 'xl/worksheets/sheet5.xml';
    if (zip.file(pensionSheetPath) && pensionData) {
      let pensionXml = await zip.file(pensionSheetPath).async('text');
      // 摘要
      pensionXml = this.setNumericCell(pensionXml, `B${summaryRow}`, pensionData.expectedTotal || 0);
      pensionXml = this.setTextCell(pensionXml, `D${summaryRow}`, '已核對');
      // 明細 20 人 (C:雇主提繳, D:個人自提)
      employeeMappings.forEach((emp, idx) => {
        const row = detailMonthStartRow + idx;
        const record = pensionData.records.find(r => r.name === emp.name);
        const e = record ? (record.employerAmount || 0) : 0;
        const p = record ? (record.personalAmount || 0) : 0;
        pensionXml = this.setNumericCell(pensionXml, `C${row}`, e);
        pensionXml = this.setNumericCell(pensionXml, `D${row}`, p);
      });
      zip.file(pensionSheetPath, pensionXml);
    }

    // 4. 更新健保 (sheet3.xml)
    const healthSheetName = hrConfig.sheets.healthInsurance.sheetName || '健保';
    const healthSheetPath = sheetPathMap.get(healthSheetName) || 'xl/worksheets/sheet3.xml';
    if (zip.file(healthSheetPath) && healthData) {
      let healthXml = await zip.file(healthSheetPath).async('text');
      // 摘要：寫入公司實際支付健保費總額 (投保單位負擔)，避免計入員工個人自付額導致人事成本失真
      healthXml = this.setNumericCell(healthXml, `B${summaryRow}`, healthData.totalCompany || 0);
      healthXml = this.setTextCell(healthXml, `D${summaryRow}`, '已核對');
      // 明細 20 人 (C:個人負擔, D:公司負擔)
      employeeMappings.forEach((emp, idx) => {
        const row = detailMonthStartRow + idx;
        const record = healthData.records.find(r => r.name === emp.name);
        const p = record ? (record.personal || 0) : 0;
        const c = record ? (record.company || 0) : 0;
        healthXml = this.setNumericCell(healthXml, `C${row}`, p);
        healthXml = this.setNumericCell(healthXml, `D${row}`, c);
      });
      zip.file(healthSheetPath, healthXml);
    }

    // 5. 動態將員工分頁更名為真實姓名
    await this.updateWorkbookSheetNames(zip, employeeMappings);
  }

  /**
   * 產出更新後 Excel 並觸發瀏覽器下載
   */
  static async downloadZip(zip, fileName = 'cost_updated.xlsx') {
    const blob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    if (typeof saveAs !== 'undefined') {
      saveAs(blob, fileName);
    } else {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(link.href);
    }
  }
}

window.ExcelWriter = ExcelWriter;
