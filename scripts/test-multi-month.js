const fs = require('fs');
const JSZip = require('jszip');

// 模擬 ExcelWriter
class ExcelWriter {
  static async loadZip(source) {
    let arrayBuffer = source;
    if (typeof source === 'string') {
      arrayBuffer = fs.readFileSync(source);
    }
    return await JSZip.loadAsync(arrayBuffer);
  }

  static async getSheetPathMap(zip) {
    const sheetMap = new Map();
    try {
      const wbXml = await zip.file('xl/workbook.xml').async('text');
      const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');
      const relsMap = new Map();
      const relMatches = relsXml.matchAll(/<Relationship\s+Id="([^"]+)"\s+Type="[^"]*worksheet"\s+Target="([^"]+)"/g);
      for (const m of relMatches) {
        let target = m[2];
        if (target.startsWith('/')) target = target.substring(1);
        else if (!target.startsWith('xl/')) target = 'xl/' + target;
        relsMap.set(m[1], target);
      }
      const sheetMatches = wbXml.matchAll(/<sheet\s+name="([^"]+)"\s+sheetId="([^"]+)"(?:\s+state="[^"]+")?\s+r:id="([^"]+)"/g);
      for (const m of sheetMatches) {
        const name = m[1];
        const rId = m[3];
        if (relsMap.has(rId)) {
          sheetMap.set(name, relsMap.get(rId));
        }
      }
    } catch (e) {
      console.warn('動態解析工作表對應失敗，採用預設順序', e);
      sheetMap.set('員工設定', 'xl/worksheets/sheet2.xml');
      sheetMap.set('健保', 'xl/worksheets/sheet3.xml');
      sheetMap.set('勞保', 'xl/worksheets/sheet4.xml');
      sheetMap.set('勞退', 'xl/worksheets/sheet5.xml');
    }
    return sheetMap;
  }

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

  static async updateCostWorkbook(zip, options) {
    const {
      targetMonth = '2026-08-01',
      employeeMappings = [],
      laborData = null,
      pensionData = null,
      healthData = null
    } = options;

    const sheetPathMap = await this.getSheetPathMap(zip);
    const monthMatch = targetMonth.match(/\d{4}-(\d{2})/);
    const month = monthMatch ? parseInt(monthMatch[1], 10) : 8;

    // 1. 員工設定
    const empSheetPath = sheetPathMap.get('員工設定') || 'xl/worksheets/sheet2.xml';
    if (zip.file(empSheetPath) && employeeMappings.length > 0) {
      let empXml = await zip.file(empSheetPath).async('text');
      employeeMappings.forEach((emp, idx) => {
        const row = 4 + idx;
        if (emp.name) {
          empXml = this.setTextCell(empXml, `B${row}`, emp.name);
        }
      });
      zip.file(empSheetPath, empXml);
    }

    const summaryRow = 4 + (month - 1);
    const detailMonthStartRow = 21 + (month - 1) * 10;

    // 2. 勞保
    const laborSheetPath = sheetPathMap.get('勞保') || 'xl/worksheets/sheet4.xml';
    if (zip.file(laborSheetPath) && laborData) {
      let laborXml = await zip.file(laborSheetPath).async('text');
      laborXml = this.setNumericCell(laborXml, `B${summaryRow}`, laborData.totalCompany || 0);
      laborXml = this.setTextCell(laborXml, `D${summaryRow}`, '已核對');
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

    // 3. 勞退
    const pensionSheetPath = sheetPathMap.get('勞退') || 'xl/worksheets/sheet5.xml';
    if (zip.file(pensionSheetPath) && pensionData) {
      let pensionXml = await zip.file(pensionSheetPath).async('text');
      pensionXml = this.setNumericCell(pensionXml, `B${summaryRow}`, pensionData.expectedTotal || 0);
      pensionXml = this.setTextCell(pensionXml, `D${summaryRow}`, '已核對');
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

    // 4. 健保
    const healthSheetPath = sheetPathMap.get('健保') || 'xl/worksheets/sheet3.xml';
    if (zip.file(healthSheetPath) && healthData) {
      let healthXml = await zip.file(healthSheetPath).async('text');
      healthXml = this.setNumericCell(healthXml, `B${summaryRow}`, healthData.expectedTotal || 0);
      healthXml = this.setTextCell(healthXml, `D${summaryRow}`, '已核對');
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
  }
}

async function runMultiMonthTest() {
  console.log('--- 測試開始：模擬連續兩個月 (8月 -> 9月) 寫入同一 Excel ---');

  const empMappings = [
    { code: 'E01', name: '李大同' },
    { code: 'E02', name: '王小明' },
    { code: 'E03', name: '張美麗' },
    { code: 'E04', name: '陳志強' },
    { code: 'E05', name: '林小華' }
  ];

  // 1. 第一次寫入：8 月
  console.log('步驟 1: 寫入 2026-08 (8月)...');
  const zip1 = await ExcelWriter.loadZip('cost.xlsx');
  await ExcelWriter.updateCostWorkbook(zip1, {
    targetMonth: '2026-08-01',
    employeeMappings: empMappings,
    laborData: {
      totalCompany: 21915,
      records: [
        { name: '李大同', personal: 1100, company: 3850 },
        { name: '王小明', personal: 900, company: 3150 },
        { name: '張美麗', personal: 800, company: 2800 },
        { name: '陳志強', personal: 700, company: 2450 },
        { name: '林小華', personal: 600, company: 2100 }
      ]
    },
    pensionData: {
      expectedTotal: 15336,
      records: [
        { name: '李大同', employerAmount: 3000, personalAmount: 0 },
        { name: '王小明', employerAmount: 2500, personalAmount: 0 },
        { name: '張美麗', employerAmount: 2200, personalAmount: 0 },
        { name: '陳志強', employerAmount: 1900, personalAmount: 0 },
        { name: '林小華', employerAmount: 1600, personalAmount: 0 }
      ]
    },
    healthData: {
      expectedTotal: 34168,
      records: [
        { name: '李大同', personal: 1200, company: 4200 },
        { name: '王小明', personal: 1000, company: 3500 },
        { name: '張美麗', personal: 900, company: 3150 },
        { name: '陳志強', personal: 800, company: 2800 },
        { name: '林小華', personal: 700, company: 2450 }
      ]
    }
  });

  const bufMonth8 = await zip1.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('scripts/test_cost_month8.xlsx', bufMonth8);
  console.log('✅ 8月寫入成功：scripts/test_cost_month8.xlsx');

  // 2. 第二次寫入：以 8 月產出的 Excel 為基底，寫入 2026-09 (9月)
  console.log('步驟 2: 以 8 月檔案為基底，寫入 2026-09 (9月)...');
  const zip2 = await ExcelWriter.loadZip(bufMonth8); // 直接以 8 月的 buffer 載入
  await ExcelWriter.updateCostWorkbook(zip2, {
    targetMonth: '2026-09-01',
    employeeMappings: empMappings,
    laborData: {
      totalCompany: 22500, // 9月金額有微調
      records: [
        { name: '李大同', personal: 1150, company: 3900 },
        { name: '王小明', personal: 920, company: 3200 },
        { name: '張美麗', personal: 850, company: 2900 },
        { name: '陳志強', personal: 720, company: 2500 },
        { name: '林小華', personal: 610, company: 2150 }
      ]
    },
    pensionData: {
      expectedTotal: 15800,
      records: [
        { name: '李大同', employerAmount: 3100, personalAmount: 0 },
        { name: '王小明', employerAmount: 2550, personalAmount: 0 },
        { name: '張美麗', employerAmount: 2250, personalAmount: 0 },
        { name: '陳志強', employerAmount: 1950, personalAmount: 0 },
        { name: '林小華', employerAmount: 1650, personalAmount: 0 }
      ]
    },
    healthData: {
      expectedTotal: 35000,
      records: [
        { name: '李大同', personal: 1250, company: 4300 },
        { name: '王小明', personal: 1050, company: 3600 },
        { name: '張美麗', personal: 920, company: 3200 },
        { name: '陳志強', personal: 820, company: 2850 },
        { name: '林小華', personal: 710, company: 2500 }
      ]
    }
  });

  const bufMonth9 = await zip2.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('scripts/test_cost_month8_and_9.xlsx', bufMonth9);
  console.log('✅ 9月寫入成功：scripts/test_cost_month8_and_9.xlsx');

  // 3. 驗證儲存格內容
  const laborXml = await zip2.file('xl/worksheets/sheet4.xml').async('text');
  
  // 檢查 8月 摘要 (B11) 與 明細 (C91, D91)
  const b11 = laborXml.match(/<c r="B11"[^>]*><v>(\d+)<\/v><\/c>/);
  const c91 = laborXml.match(/<c r="C91"[^>]*><v>(\d+)<\/v><\/c>/);
  const d91 = laborXml.match(/<c r="D91"[^>]*><v>(\d+)<\/v><\/c>/);
  console.log('8月摘要 B11:', b11 ? b11[1] : 'null', '(預期: 21915)');
  console.log('8月李大同勞保 C91:', c91 ? c91[1] : 'null', '(預期: 1100)');
  console.log('8月李大同公司 D91:', d91 ? d91[1] : 'null', '(預期: 3850)');

  // 檢查 9月 摘要 (B12) 與 明細 (C101, D101)
  const b12 = laborXml.match(/<c r="B12"[^>]*><v>(\d+)<\/v><\/c>/);
  const c101 = laborXml.match(/<c r="C101"[^>]*><v>(\d+)<\/v><\/c>/);
  const d101 = laborXml.match(/<c r="D101"[^>]*><v>(\d+)<\/v><\/c>/);
  console.log('9月摘要 B12:', b12 ? b12[1] : 'null', '(預期: 22500)');
  console.log('9月李大同勞保 C101:', c101 ? c101[1] : 'null', '(預期: 1150)');
  console.log('9月李大同公司 D101:', d101 ? d101[1] : 'null', '(預期: 3900)');

  const isOk = (b11 && b11[1] === '21915') &&
               (c91 && c91[1] === '1100') &&
               (d91 && d91[1] === '3850') &&
               (b12 && b12[1] === '22500') &&
               (c101 && c101[1] === '1150') &&
               (d101 && d101[1] === '3900');

  console.log('\n==== 檢驗結果 ====\n', isOk ? '🎉 雙月份完美並存，彼此互不干擾！' : '❌ 資料有缺漏或覆蓋！');
}

runMultiMonthTest();
