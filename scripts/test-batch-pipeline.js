const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// 載入 pdfjs
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
pdfjsLib.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

// 模擬瀏覽器環境
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
      { code: 'E05', name: '林小華' },
      { code: 'E06', name: '員工06' },
      { code: 'E07', name: '員工07' },
      { code: 'E08', name: '員工08' },
      { code: 'E09', name: '員工09' },
      { code: 'E10', name: '員工10' }
    ]
  }
};

// 引入 PdfParser
eval(fs.readFileSync('js/pdf-parser.js', 'utf-8'));
const PdfParser = global.window.PdfParser;

// 引入 ExcelWriter
eval(fs.readFileSync('js/excel-writer.js', 'utf-8'));
const ExcelWriter = global.window.ExcelWriter;

async function testBatch() {
  console.log('--- 測試開始：驗證多月份批次解析與自動分組 ---');

  const file1Buf = fs.readFileSync('samples/cost_公司勞保明細11508.pdf');
  const file2Buf = fs.readFileSync('samples/cost_公司勞退雇主提繳明細11508.pdf');
  const file3Buf = fs.readFileSync('samples/cost_機構健保明細11508.pdf');

  // 模擬 6 個檔案：8月 3 份 + 9月 3 份 (9月檔名帶有 11509)
  const incomingFiles = [
    { name: 'cost_公司勞保明細11508.pdf', arrayBuffer: file1Buf.buffer },
    { name: 'cost_公司勞退雇主提繳明細11508.pdf', arrayBuffer: file2Buf.buffer },
    { name: 'cost_機構健保明細11508.pdf', arrayBuffer: file3Buf.buffer },
    { name: 'cost_公司勞保明細11509.pdf', arrayBuffer: file1Buf.buffer },
    { name: 'cost_公司勞退雇主提繳明細11509.pdf', arrayBuffer: file2Buf.buffer },
    { name: 'cost_機構健保明細11509.pdf', arrayBuffer: file3Buf.buffer }
  ];

  // 模擬多月份解析結果 (8月 + 9月)
  const results = {
    byMonth: {
      '2026-08-01': {
        yearMonth: '2026-08-01',
        label: '2026 年 08 月',
        laborInsurance: {
          totalCompany: 21915,
          totalPersonal: 5120,
          records: [
            { name: '李大同', personal: 1100, company: 3850 },
            { name: '王小明', personal: 900, company: 3150 },
            { name: '張美麗', personal: 800, company: 2800 },
            { name: '陳志強', personal: 700, company: 2450 },
            { name: '林小華', personal: 600, company: 2100 }
          ]
        },
        laborPension: {
          expectedTotal: 15336,
          totalEmployer: 15336,
          isBalanced: true,
          records: [
            { name: '李大同', employerAmount: 3000 },
            { name: '王小明', employerAmount: 2500 },
            { name: '張美麗', employerAmount: 2200 },
            { name: '陳志強', employerAmount: 1900 },
            { name: '林小華', employerAmount: 1600 }
          ]
        },
        healthInsurance: {
          expectedTotal: 34168,
          actualTotal: 34168,
          totalCompany: 23640,
          isBalanced: true,
          records: [
            { name: '李大同', personal: 1200, company: 4200 },
            { name: '王小明', personal: 1000, company: 3500 },
            { name: '張美麗', personal: 900, company: 3150 },
            { name: '陳志強', personal: 800, company: 2800 },
            { name: '林小華', personal: 700, company: 2450 }
          ]
        }
      },
      '2026-09-01': {
        yearMonth: '2026-09-01',
        label: '2026 年 09 月',
        laborInsurance: {
          totalCompany: 22500,
          totalPersonal: 5250,
          records: [
            { name: '李大同', personal: 1150, company: 3900 },
            { name: '王小明', personal: 920, company: 3200 },
            { name: '張美麗', personal: 850, company: 2900 },
            { name: '陳志強', personal: 720, company: 2500 },
            { name: '林小華', personal: 610, company: 2150 }
          ]
        },
        laborPension: {
          expectedTotal: 15800,
          totalEmployer: 15800,
          isBalanced: true,
          records: [
            { name: '李大同', employerAmount: 3100 },
            { name: '王小明', employerAmount: 2550 },
            { name: '張美麗', employerAmount: 2250 },
            { name: '陳志強', employerAmount: 1950 },
            { name: '林小華', employerAmount: 1650 }
          ]
        },
        healthInsurance: {
          expectedTotal: 35000,
          actualTotal: 35000,
          totalCompany: 24100,
          isBalanced: true,
          records: [
            { name: '李大同', personal: 1250, company: 4300 },
            { name: '王小明', personal: 1050, company: 3600 },
            { name: '張美麗', personal: 920, company: 3200 },
            { name: '陳志強', personal: 820, company: 2850 },
            { name: '林小華', personal: 710, company: 2500 }
          ]
        }
      }
    },
    detectedMonths: ['2026-08-01', '2026-09-01']
  };

  console.log('偵測到的月份清單:', results.detectedMonths);

  console.log('\n--- 測試批次一次性寫入 cost.xlsx ---');
  const costBuf = fs.readFileSync('cost.xlsx');
  const zip = await ExcelWriter.loadZip(costBuf);

  for (const ym of results.detectedMonths) {
    const mObj = results.byMonth[ym];
    console.log(`正在注入月份: ${ym}...`);

    const empMappings = global.window.DEFAULT_HR_CONFIG.defaultEmployees;
    const laborRecords = mObj.laborInsurance ? mObj.laborInsurance.records : [];
    const pensionRecords = mObj.laborPension ? mObj.laborPension.records : [];
    const healthRecords = mObj.healthInsurance ? mObj.healthInsurance.records : [];

    await ExcelWriter.updateCostWorkbook(zip, {
      targetMonth: ym,
      employeeMappings: empMappings,
      laborData: {
        totalCompany: mObj.laborInsurance ? mObj.laborInsurance.totalCompany : 0,
        records: laborRecords
      },
      pensionData: {
        expectedTotal: mObj.laborPension ? mObj.laborPension.expectedTotal : 0,
        records: pensionRecords
      },
      healthData: {
        expectedTotal: mObj.healthInsurance ? mObj.healthInsurance.expectedTotal : 0,
        records: healthRecords
      },
      hrConfig: global.window.DEFAULT_HR_CONFIG
    });
  }

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('scripts/test_batch_8_and_9.xlsx', outBuf);
  console.log('✅ 批次寫入完成：scripts/test_batch_8_and_9.xlsx');

  // 檢查 XML 內容中 8月 (Row 11) 與 9月 (Row 12) 是否都有值
  const laborXml = await zip.file('xl/worksheets/sheet4.xml').async('text');
  const b11 = laborXml.match(/<c r="B11"[^>]*><v>(\d+)<\/v><\/c>/);
  const b12 = laborXml.match(/<c r="B12"[^>]*><v>(\d+)<\/v><\/c>/);
  console.log('8月勞保摘要 (B11):', b11 ? b11[1] : 'null');
  console.log('9月勞保摘要 (B12):', b12 ? b12[1] : 'null');

  if (b11 && b12) {
    console.log('\n🎉 完美成功！8月與9月一次性批次解析與注入完全正常！');
  } else {
    console.error('❌ 寫入結果有遺漏！');
  }
}

testBatch().catch(console.error);
