const fs = require('fs');
const JSZip = require('jszip');

async function fixRanges() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  
  let totalFixed = 0;
  for (let i = 6; i <= 26; i++) {
    const path = `xl/worksheets/sheet${i}.xml`;
    if (!zip.file(path)) continue;

    let xml = await zip.file(path).async('text');
    const beforeCount = (xml.match(/140/g) || []).length;
    if (beforeCount > 0) {
      // 替換所有 $21:$...$140 或 $140
      xml = xml.replace(/(\$[A-Z]+\$\d+:\$[A-Z]+)\$140/g, '$1$260');
      xml = xml.replace(/(\$[A-Z]+)\$140/g, '$1$260');
      xml = xml.replace(/:([A-Z]+)140/g, ':$1260');
      xml = xml.replace(/\$140\b/g, '$260');
      
      const afterCount = (xml.match(/140/g) || []).length;
      console.log(`sheet${i}: 140 before = ${beforeCount}, after = ${afterCount}`);
      zip.file(path, xml);
      totalFixed += (beforeCount - afterCount);
    }
  }

  console.log(`Total replaced: ${totalFixed}`);
  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('cost.xlsx', outBuf);
  console.log('Successfully updated cost.xlsx with 260 range!');
}

fixRanges().catch(console.error);
