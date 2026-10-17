const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function testJsZip() {
  const originalBuffer = fs.readFileSync(path.join(__dirname, '..', 'cost.xlsx'));
  const zip = await JSZip.loadAsync(originalBuffer);

  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });

  const outPath = path.join(__dirname, 'test_jszip_pure.xlsx');
  fs.writeFileSync(outPath, outBuffer);
  console.log('Saved test_jszip_pure.xlsx');
}

testJsZip().catch(console.error);
