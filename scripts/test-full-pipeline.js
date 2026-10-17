const fs = require('fs');
const JSZip = require('jszip');

// 引入 ExcelWriter 邏輯模擬
const excelWriterCode = fs.readFileSync('js/excel-writer.js', 'utf8');
const vm = require('vm');
const windowObj = {
  DEFAULT_HR_CONFIG: {
    sheets: {
      employeeSettings: { sheetName: '員工設定', startRow: 4, endRow: 23 },
      healthInsurance: { sheetName: '健保', summaryStartRow: 4, detailStartRow: 21, detailEndRow: 260 },
      laborInsurance: { sheetName: '勞保', summaryStartRow: 4, detailStartRow: 21, detailEndRow: 260 },
      laborPension: { sheetName: '勞退', summaryStartRow: 4, detailStartRow: 21, detailEndRow: 260 }
    }
  }
};
const ctx = vm.createContext({
  window: windowObj,
  JSZip: JSZip,
  console: console
});
vm.runInContext(excelWriterCode, ctx);
const ExcelWriter = ctx.window.ExcelWriter;

async function runTest() {
  console.log('=== 階段 1：8 月份初次寫入 (7 位員工) ===');
  const baseZip = await ExcelWriter.loadZip(fs.readFileSync('cost.xlsx'));

  // 模擬 8 月份名冊
  const empAugust = [
    { code: 'E01', name: '陳萱蘋' },
    { code: 'E02', name: '林小美' },
    { code: 'E03', name: '張志豪' },
    { code: 'E04', name: '李佳穎' },
    { code: 'E05', name: '王大同' },
    { code: 'E06', name: '周杰倫' },
    { code: 'E07', name: '蔡依林' }
  ];

  // 建立 20 人的 mappings
  const mappingsAugust = [];
  for (let i = 1; i <= 20; i++) {
    const code = `E${String(i).padStart(2, '0')}`;
    const found = empAugust.find(e => e.code === code);
    mappingsAugust.push({
      code,
      name: found ? found.name : `員工${String(i).padStart(2, '0')}`
    });
  }

  // 8 月資料
  const laborDataAug = {
    totalCompany: 25000,
    records: [
      { name: '陳萱蘋', personal: 1100, company: 3500 },
      { name: '林小美', personal: 900, company: 3000 },
      { name: '張志豪', personal: 1200, company: 4000 }
    ]
  };
  const pensionDataAug = {
    expectedTotal: 15000,
    records: [
      { name: '陳萱蘋', employerAmount: 2200, personalAmount: 0 },
      { name: '林小美', employerAmount: 1800, personalAmount: 0 },
      { name: '張志豪', employerAmount: 2400, personalAmount: 0 }
    ]
  };
  const healthDataAug = {
    totalCompany: 4650, // 投保單位負擔
    records: [
      { name: '陳萱蘋', personal: 1200, company: 1500 },
      { name: '林小美', personal: 1000, company: 1300 },
      { name: '張志豪', personal: 1400, company: 1850 }
    ]
  };

  await ExcelWriter.updateCostWorkbook(baseZip, {
    targetMonth: '2026-08-01',
    employeeMappings: mappingsAugust,
    laborData: laborDataAug,
    pensionData: pensionDataAug,
    healthData: healthDataAug,
    hrConfig: windowObj.DEFAULT_HR_CONFIG
  });

  const augBuf = await baseZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('cost_august_result.xlsx', augBuf);
  console.log('8 月份 Excel 已產出：cost_august_result.xlsx');

  console.log('\n=== 階段 2：下個月 (9 月) 輸入既有 Excel，並加入新員工 ===');
  // 模擬載入上個月產出的 Excel
  const step2Zip = await ExcelWriter.loadZip(fs.readFileSync('cost_august_result.xlsx'));
  const existingEmps = await ExcelWriter.readEmployeesFromExcel(step2Zip, windowObj.DEFAULT_HR_CONFIG);
  console.log('從 Excel 讀取到的既有員工數量:', existingEmps.filter(e => e.isReal).length);
  existingEmps.filter(e => e.isReal).forEach(e => console.log(`- ${e.code}: ${e.name}`));

  // 模擬 9 月份：林小美 (E02) 離職/無出勤，新增 2 位兼職新員工「郭台銘」、「張忠謀」
  const pdfEmpsSept = ['陳萱蘋', '張志豪', '李佳穎', '王大同', '周杰倫', '蔡依林', '郭台銘', '張忠謀'];

  // 合併名冊池
  const pool = [...existingEmps];
  pdfEmpsSept.forEach(name => {
    if (!pool.some(p => p.name === name && p.isReal)) {
      // 找第一個空白位置
      const emptySlot = pool.find(p => !p.isReal);
      if (emptySlot) {
        emptySlot.name = name;
        emptySlot.isReal = true;
      }
    }
  });

  console.log('\n合併後 9 月的名冊池分配:');
  pool.filter(p => p.isReal).forEach(p => console.log(`- ${p.code}: ${p.name}`));

  // 寫入 9 月份資料
  const mappingsSept = pool.map(p => ({ code: p.code, name: p.name }));
  const laborDataSept = {
    totalCompany: 28000,
    records: [
      { name: '陳萱蘋', personal: 1100, company: 3500 },
      { name: '郭台銘', personal: 1500, company: 4500 },
      { name: '張忠謀', personal: 1800, company: 5500 }
    ]
  };
  const pensionDataSept = {
    expectedTotal: 18000,
    records: [
      { name: '陳萱蘋', employerAmount: 2200, personalAmount: 0 },
      { name: '郭台銘', employerAmount: 3000, personalAmount: 0 },
      { name: '張忠謀', employerAmount: 3500, personalAmount: 0 }
    ]
  };
  const healthDataSept = {
    totalCompany: 5200,
    records: [
      { name: '陳萱蘋', personal: 1200, company: 1500 },
      { name: '郭台銘', personal: 1600, company: 2000 },
      { name: '張忠謀', personal: 1900, company: 2400 }
    ]
  };

  await ExcelWriter.updateCostWorkbook(step2Zip, {
    targetMonth: '2026-09-01',
    employeeMappings: mappingsSept,
    laborData: laborDataSept,
    pensionData: pensionDataSept,
    healthData: healthDataSept,
    hrConfig: windowObj.DEFAULT_HR_CONFIG
  });

  const septBuf = await step2Zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('cost_september_result.xlsx', septBuf);
  console.log('9 月份累加 Excel 已產出：cost_september_result.xlsx');
}

runTest().catch(console.error);
