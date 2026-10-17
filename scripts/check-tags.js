const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function checkSheetDiff() {
  const origBuffer = fs.readFileSync(path.join(__dirname, '..', 'cost.xlsx'));
  const origZip = await JSZip.loadAsync(origBuffer);

  const cleanBuffer = fs.readFileSync(path.join(__dirname, 'test_clean_output.xlsx'));
  const cleanZip = await JSZip.loadAsync(cleanBuffer);

  for (const s of ['sheet3.xml', 'sheet4.xml', 'sheet5.xml', 'sheet2.xml']) {
    const o = await origZip.file(`xl/worksheets/${s}`).async('text');
    const c = await cleanZip.file(`xl/worksheets/${s}`).async('text');
    console.log(`${s} orig length: ${o.length}, clean length: ${c.length}`);
    
    // 檢查是否有非法 xml (例如 <c ...><v>... 或重複標籤)
    // 測試用正規表示式找奇怪的標籤
    const doubleC = c.match(/<c[^>]*<c/g);
    if (doubleC) console.log(`ALERT: ${s} has double <c tag!`);
    
    // 找孤立的 </c>
    const openCount = (c.match(/<c\b/g) || []).length;
    const closeCount = (c.match(/<\/c>/g) || []).length;
    const selfCloseCount = (c.match(/<c\b[^>]*\/>/g) || []).length;
    console.log(`${s} open: ${openCount}, close: ${closeCount}, selfClose: ${selfCloseCount}, sum: ${closeCount + selfCloseCount}`);
  }
}

checkSheetDiff().catch(console.error);
