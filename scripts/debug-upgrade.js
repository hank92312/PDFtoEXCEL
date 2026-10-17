const fs = require('fs');
const JSZip = require('jszip');

async function debugDiff() {
  const origBuf = fs.readFileSync('cost_backup_10emp.xlsx');
  const newBuf = fs.readFileSync('cost.xlsx');
  const origZip = await JSZip.loadAsync(origBuf);
  const newZip = await JSZip.loadAsync(newBuf);

  console.log('--- 檢查 xl/workbook.xml ---');
  const newWb = await newZip.file('xl/workbook.xml').async('text');
  console.log('newWb length:', newWb.length);

  console.log('--- 檢查 xl/_rels/workbook.xml.rels ---');
  const newRels = await newZip.file('xl/_rels/workbook.xml.rels').async('text');
  console.log('newRels:', newRels);

  console.log('--- 檢查 [Content_Types].xml ---');
  const newCt = await newZip.file('[Content_Types].xml').async('text');
  console.log('newCt:', newCt);

  console.log('--- 檢查 sheet26.xml (前 500 字) ---');
  const s26 = await newZip.file('xl/worksheets/sheet26.xml').async('text');
  console.log(s26.substring(0, 500));
}

debugDiff().catch(console.error);
