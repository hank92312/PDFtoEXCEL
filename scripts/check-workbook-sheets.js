const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  const wbXml = await zip.file('xl/workbook.xml').async('text');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');

  console.log('--- Sheets in workbook.xml ---');
  const sheetRegex = /<sheet\s+[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*r:id="([^"]+)"/g;
  let match;
  while ((match = sheetRegex.exec(wbXml)) !== null) {
    console.log(`Sheet: name="${match[1]}" | sheetId="${match[2]}" | r:id="${match[3]}"`);
  }
}

main().catch(console.error);
