const fs = require('fs');
const JSZip = require('jszip');
global.JSZip = JSZip;
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
pdfjsLib.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

global.pdfjsLib = pdfjsLib;
global.window = {
  DEFAULT_HR_CONFIG: {
    defaultYear: 2026,
    sheets: {
      employeeSettings: { sheetName: '員工設定', startRow: 4 },
      laborInsurance: { sheetName: '勞保', summaryStartRow: 4, detailStartRow: 21 },
      laborPension: { sheetName: '勞退', summaryStartRow: 4, detailStartRow: 21 },
      healthInsurance: { sheetName: '健保', summaryStartRow: 4, detailStartRow: 21 }
    },
    defaultEmployees: [
      { code: 'E01', name: '李大同' },
      { code: 'E02', name: '王小明' },
      { code: 'E03', name: '張美麗' },
      { code: 'E04', name: '陳志強' },
      { code: 'E05', name: '陳萱蘋' },
      { code: 'E06', name: '歐明安' },
      { code: 'E07', name: '林家羽' },
      { code: 'E08', name: '員工08' },
      { code: 'E09', name: '員工09' },
      { code: 'E10', name: '員工10' }
    ]
  }
};

eval(fs.readFileSync('js/pdf-parser.js', 'utf-8'));
eval(fs.readFileSync('js/excel-writer.js', 'utf-8'));

async function test() {
  const buf = fs.readFileSync('samples/cost_機構健保明細11508.pdf');
  const res = await global.window.PdfParser.parseAllFiles([
    { name: 'cost_機構健保明細11508.pdf', arrayBuffer: buf.buffer }
  ]);
  const healthData = res.byMonth['2026-08-01'].healthInsurance;
  console.log('解析官方應繳總額 expectedTotal:', healthData.expectedTotal);
  console.log('解析官方投保單位負擔 officialCompanyBurden:', healthData.officialCompanyBurden);
  console.log('計算明細單位負擔加總 totalCompany:', healthData.totalCompany);

  const costBuf = fs.readFileSync('cost.xlsx');
  const zip = await global.window.ExcelWriter.loadZip(costBuf);

  await global.window.ExcelWriter.updateCostWorkbook(zip, {
    targetMonth: '2026-08-01',
    employeeMappings: global.window.DEFAULT_HR_CONFIG.defaultEmployees,
    healthData: {
      totalCompany: healthData.totalCompany,
      expectedTotal: healthData.expectedTotal,
      records: healthData.records
    },
    hrConfig: global.window.DEFAULT_HR_CONFIG
  });

  const healthXml = await zip.file('xl/worksheets/sheet3.xml').async('text');
  const b11 = healthXml.match(/<c r="B11"[^>]*><v>(\d+)<\/v><\/c>/);
  console.log('寫入 Excel 健保工作表 B11 (公司實際支付健保費總額):', b11 ? b11[1] : 'null');
  
  if (b11 && b11[1] === '4650') {
    console.log('🎉 驗證完全正確！寫入的是投保單位負擔 4650，而非包含個人自付的 7630！');
  } else {
    console.error('❌ 驗證失敗，B11 數值為:', b11 ? b11[1] : 'null');
  }
}

test().catch(console.error);
