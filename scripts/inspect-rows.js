const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  const sheet4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
  
  const m21 = sheet4.match(/<row r="21"[^>]*>[\s\S]*?<\/row>/);
  console.log('Row 21:', m21 ? m21[0] : 'not found');
  
  const m161 = sheet4.match(/<row r="161"[^>]*>[\s\S]*?<\/row>/);
  console.log('Row 161:', m161 ? m161[0] : 'not found');

  const m162 = sheet4.match(/<row r="162"[^>]*>[\s\S]*?<\/row>/);
  console.log('Row 162:', m162 ? m162[0] : 'not found');
}

main().catch(console.error);
