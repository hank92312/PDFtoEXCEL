const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  const s6 = await zip.file('xl/worksheets/sheet6.xml').async('text');
  console.log('--- sheet6 Row 4 (Header) ---');
  console.log(s6.match(/<row r="4"[^>]*>[\s\S]*?<\/row>/)?.[0]);
  console.log('--- sheet6 Row 12 (8月) ---');
  console.log(s6.match(/<row r="12"[^>]*>[\s\S]*?<\/row>/)?.[0]);
}

main().catch(console.error);
