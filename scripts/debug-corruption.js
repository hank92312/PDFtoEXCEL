const ExcelJS = require('exceljs');
const path = require('path');

async function testStep() {
  const originalPath = path.join(__dirname, '..', 'cost.xlsx');

  // 測試 1: 僅讀取並直接重新存檔
  const wb1 = new ExcelJS.Workbook();
  await wb1.xlsx.readFile(originalPath);
  await wb1.xlsx.writeFile(path.join(__dirname, 'test_1_pure_resave.xlsx'));
  console.log('Test 1 pure resave generated');

  // 測試 2: 只修改員工設定中的一個儲存格 (B4)
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(originalPath);
  wb2.getWorksheet('員工設定').getCell('B4').value = '吳金燕';
  await wb2.xlsx.writeFile(path.join(__dirname, 'test_2_only_emp.xlsx'));
  console.log('Test 2 only emp generated');

  // 測試 3: 只修改勞保摘要 (B11)
  const wb3 = new ExcelJS.Workbook();
  await wb3.xlsx.readFile(originalPath);
  wb3.getWorksheet('勞保').getCell('B11').value = 6453;
  await wb3.xlsx.writeFile(path.join(__dirname, 'test_3_only_summary.xlsx'));
  console.log('Test 3 only summary generated');

  // 測試 4: 只修改勞保明細 C91, D91
  const wb4 = new ExcelJS.Workbook();
  await wb4.xlsx.readFile(originalPath);
  wb4.getWorksheet('勞保').getCell('C91').value = 870;
  wb4.getWorksheet('勞保').getCell('D91').value = 3087;
  await wb4.xlsx.writeFile(path.join(__dirname, 'test_4_only_detail.xlsx'));
  console.log('Test 4 only detail generated');
}

testStep().catch(console.error);
