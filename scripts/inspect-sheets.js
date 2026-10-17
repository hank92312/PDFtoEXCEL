const fs = require('fs');
const JSZip = require('jszip');

async function inspect() {
  const buf = fs.readFileSync('cost.xlsx');
  const zip = await JSZip.loadAsync(buf);
  
  // 1. 查看 workbook.xml 中所有工作表清單
  const wbXml = await zip.file('xl/workbook.xml').async('text');
  const sheets = [...wbXml.matchAll(/<sheet\s+name="([^"]+)"\s+sheetId="([^"]+)"[^>]*r:id="([^"]+)"/g)].map(m => m[1]);
  console.log('現有工作表清單:', sheets);

  // 2. 查看 員工設定 工作表
  const empSetXml = await zip.file('xl/worksheets/sheet2.xml').async('text');
  console.log('\n--- 員工設定 工作表 (前 15 列) ---');
  const empRows = [...empSetXml.matchAll(/<row\s+r="(\d+)"[^>]*>(.*?)<\/row>/g)];
  for (const r of empRows.slice(0, 16)) {
    console.log('Row ' + r[1] + ':', r[2].replace(/<[^>]+>/g, ' ').trim());
  }

  // 3. 查看 員工01 工作表 (sheet6.xml)
  const emp1Xml = await zip.file('xl/worksheets/sheet6.xml').async('text');
  console.log('\n--- 員工01 工作表公式與結構 ---');
  const formulas = [...emp1Xml.matchAll(/<c\s+r="([^"]+)"[^>]*><f[^>]*>([^<]+)<\/f>/g)];
  console.log('前 15 個公式儲存格:');
  for (const f of formulas.slice(0, 15)) {
    console.log('  ' + f[1] + ': =' + f[2]);
  }

  // 4. 查看 健保 / 勞保 / 勞退 工作表中的公式或範圍
  const healthXml = await zip.file('xl/worksheets/sheet3.xml').async('text');
  const hFormulas = [...healthXml.matchAll(/<c\s+r="([^"]+)"[^>]*><f[^>]*>([^<]+)<\/f>/g)];
  console.log('\n健保工作表的公式:');
  for (const f of hFormulas.slice(0, 10)) {
    console.log('  ' + f[1] + ': =' + f[2]);
  }
}

inspect().catch(console.error);
