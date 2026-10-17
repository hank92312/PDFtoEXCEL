const fs = require('fs');
const JSZip = require('jszip');

async function testRead() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  const sheet2 = await zip.file('xl/worksheets/sheet2.xml').async('text');
  
  console.log('--- Reading B4~B23 from sheet2.xml ---');
  for (let row = 4; row <= 23; row++) {
    const regex = new RegExp('<c\\s+r="B' + row + '"[^>]*>(?:<is><t>([\\s\\S]*?)<\\/t><\\/is>|<v>([\\s\\S]*?)<\\/v>)?<\\/c>');
    const match = sheet2.match(regex);
    const val = match ? (match[1] || match[2] || '').trim() : '';
    console.log(`Row ${row} (E${String(row - 3).padStart(2, '0')}): "${val}"`);
  }
}

testRead().catch(console.error);
