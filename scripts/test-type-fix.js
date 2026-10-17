const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

function setNumericCell(xml, cellRef, numVal) {
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*)(>)(.*?)(<\\/c>)`);
  return xml.replace(cellRegex, (match, openPrefix, openClose, inner, closeTag) => {
    // 確保 t="n"
    let prefix = openPrefix;
    if (!prefix.includes('t="n"')) {
      prefix = prefix.replace(/\s+t="[^"]*"/, '') + ' t="n"';
    }
    return `${prefix}${openClose}<v>${numVal}</v>${closeTag}`;
  });
}

function setTextCell(xml, cellRef, textVal) {
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*)(>)(.*?)(<\\/c>)`);
  const escapedText = textVal.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return xml.replace(cellRegex, (match, openPrefix, openClose, inner, closeTag) => {
    // 關鍵！移除舊的 t="n"，換成 t="inlineStr"！
    let prefix = openPrefix.replace(/\s+t="[^"]*"/, '') + ' t="inlineStr"';
    return `${prefix}${openClose}<is><t>${escapedText}</t></is>${closeTag}`;
  });
}

async function runTestFix() {
  const templatePath = path.join(__dirname, '..', 'cost.xlsx');
  const buffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buffer);

  // 1. 員工設定
  let s2 = await zip.file('xl/worksheets/sheet2.xml').async('text');
  const empNames = ['吳金燕', '易金香', '古力予', '鄭百翔', '陳萱蘋', '歐明安', '林家羽', '員工08', '員工09', '員工10'];
  empNames.forEach((name, idx) => {
    s2 = setTextCell(s2, `B${4 + idx}`, name);
  });
  zip.file('xl/worksheets/sheet2.xml', s2);

  // 2. 勞保
  let s4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
  s4 = setNumericCell(s4, 'B11', 6453);
  s4 = setTextCell(s4, 'D11', '已核對');
  const laborData = [
    { p: 870, c: 3087 }, { p: 0, c: 40 }, { p: 0, c: 50 }, { p: 922, c: 3276 },
    { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }
  ];
  laborData.forEach((d, idx) => {
    s4 = setNumericCell(s4, `C${91 + idx}`, d.p);
    s4 = setNumericCell(s4, `D${91 + idx}`, d.c);
  });
  zip.file('xl/worksheets/sheet4.xml', s4);

  // 3. 勞退
  let s5 = await zip.file('xl/worksheets/sheet5.xml').async('text');
  s5 = setNumericCell(s5, 'B11', 6606);
  s5 = setTextCell(s5, 'D11', '已核對');
  const pensionData = [
    { e: 2088, p: 0 }, { e: 1998, p: 0 }, { e: 2520, p: 0 }, { e: 0, p: 0 },
    { e: 0, p: 0 }, { e: 0, p: 0 }, { e: 0, p: 0 }, { e: 0, p: 0 }, { e: 0, p: 0 }, { e: 0, p: 0 }
  ];
  pensionData.forEach((d, idx) => {
    s5 = setNumericCell(s5, `C${91 + idx}`, d.e);
    s5 = setNumericCell(s5, `D${91 + idx}`, d.p);
  });
  zip.file('xl/worksheets/sheet5.xml', s5);

  // 4. 健保
  let s3 = await zip.file('xl/worksheets/sheet3.xml').async('text');
  s3 = setNumericCell(s3, 'B11', 7630);
  s3 = setTextCell(s3, 'D11', '已核對');
  const healthData = [
    { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 },
    { p: 916, c: 1428 }, { p: 516, c: 1611 }, { p: 1548, c: 1611 },
    { p: 0, c: 0 }, { p: 0, c: 0 }, { p: 0, c: 0 }
  ];
  healthData.forEach((d, idx) => {
    s3 = setNumericCell(s3, `C${91 + idx}`, d.p);
    s3 = setNumericCell(s3, `D${91 + idx}`, d.c);
  });
  zip.file('xl/worksheets/sheet3.xml', s3);

  const out = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(path.join(__dirname, 'test_all_fixed.xlsx'), out);
  console.log('Saved test_all_fixed.xlsx');
}

runTestFix().catch(console.error);
