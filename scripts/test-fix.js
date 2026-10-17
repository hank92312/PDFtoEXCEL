const ExcelJS = require('exceljs');
const path = require('path');

async function testFixFilter() {
  const originalPath = path.join(__dirname, '..', 'cost.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(originalPath);

  // 檢查 definedNames
  console.log('DefinedNames before:', wb.definedNames.model);

  // 清除 _xlnm._FilterDatabase 或所有 definedNames
  // 或者只保留正常的自訂名稱
  wb.definedNames.model = wb.definedNames.model.filter(d => !d.name.includes('_xlnm._FilterDatabase'));

  console.log('DefinedNames after filter:', wb.definedNames.model);

  const outPath = path.join(__dirname, 'test_fix_filter.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log('Saved test_fix_filter.xlsx');
}

testFixFilter().catch(console.error);
