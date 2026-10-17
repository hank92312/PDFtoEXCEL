const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost_september_result.xlsx'));
  
  console.log('=== 1. workbook.xml sheets ===');
  const wbXml = await zip.file('xl/workbook.xml').async('text');
  const sheetMatches = wbXml.match(/<sheet\s+[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"/g);
  sheetMatches.forEach(m => console.log(' ', m));

  console.log('\n=== 2. 全員工彙總 (sheet26.xml) B2, Row 12 (8月), Row 13 (9月) ===');
  const sheet26 = await zip.file('xl/worksheets/sheet26.xml').async('text');
  const b2 = sheet26.match(/<c r="B2"[^>]*>[\s\S]*?<\/c>/);
  console.log('B2:', b2 ? b2[0] : 'not found');

  const r12 = sheet26.match(/<row r="12"[^>]*>[\s\S]*?<\/row>/);
  console.log('Row 12 (8月):', r12 ? r12[0] : 'not found');

  const r13 = sheet26.match(/<row r="13"[^>]*>[\s\S]*?<\/row>/);
  console.log('Row 13 (9月):', r13 ? r13[0] : 'not found');

  console.log('\n=== 3. 員工設定 (sheet2.xml) B4~B13 ===');
  const sheet2 = await zip.file('xl/worksheets/sheet2.xml').async('text');
  for (let r = 4; r <= 13; r++) {
    const c = sheet2.match(new RegExp(`<c r="B${r}"[^>]*>[\\s\\S]*?<\\/c>`));
    console.log(`B${r}:`, c ? c[0] : 'not found');
  }

  console.log('\n=== 4. 健保 (sheet3.xml) B11, B12 (公司實付摘要) 與 Row 161, 181 (明細) ===');
  const sheet3 = await zip.file('xl/worksheets/sheet3.xml').async('text');
  console.log('B11 (8月實付健保摘要):', sheet3.match(/<c r="B11"[^>]*>[\s\S]*?<\/c>/)?.[0]);
  console.log('B12 (9月實付健保摘要):', sheet3.match(/<c r="B12"[^>]*>[\s\S]*?<\/c>/)?.[0]);
  console.log('Row 161 (8月 E01 陳萱蘋 明細):', sheet3.match(/<row r="161"[^>]*>[\s\S]*?<\/row>/)?.[0]);
  console.log('Row 181 (9月 E01 陳萱蘋 明細):', sheet3.match(/<row r="181"[^>]*>[\s\S]*?<\/row>/)?.[0]);
  console.log('Row 188 (9月 E08 郭台銘 明細):', sheet3.match(/<row r="188"[^>]*>[\s\S]*?<\/row>/)?.[0]);
}

main().catch(console.error);
