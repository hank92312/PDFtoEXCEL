const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

function setNumericCell(xml, cellRef, numVal) {
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*>)(.*?)(<\\/c>)`);
  return xml.replace(cellRegex, (match, openTag, inner, closeTag) => {
    return `${openTag}<v>${numVal}</v>${closeTag}`;
  });
}

function setTextCell(xml, cellRef, textVal) {
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*>)(.*?)(<\\/c>)`);
  const escapedText = textVal.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return xml.replace(cellRegex, (match, openTag, inner, closeTag) => {
    return `${openTag}<is><t>${escapedText}</t></is>${closeTag}`;
  });
}

async function runGradual() {
  const templatePath = path.join(__dirname, '..', 'cost.xlsx');
  const buffer = fs.readFileSync(templatePath);

  // 步驟 1: 員工設定全改
  const zip1 = await JSZip.loadAsync(buffer);
  let s2 = await zip1.file('xl/worksheets/sheet2.xml').async('text');
  const empNames = ['吳金燕', '易金香', '古力予', '鄭百翔', '陳萱蘋', '歐明安', '林家羽', '員工08', '員工09', '員工10'];
  empNames.forEach((name, idx) => {
    s2 = setTextCell(s2, `B${4 + idx}`, name);
  });
  zip1.file('xl/worksheets/sheet2.xml', s2);
  fs.writeFileSync(path.join(__dirname, 'step1_all_emp.xlsx'), await zip1.generateAsync({ type: 'nodebuffer' }));

  // 步驟 2: 步驟 1 + 勞保 (sheet4.xml)
  const zip2 = await JSZip.loadAsync(await zip1.generateAsync({ type: 'nodebuffer' }));
  let s4 = await zip2.file('xl/worksheets/sheet4.xml').async('text');
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
  zip2.file('xl/worksheets/sheet4.xml', s4);
  fs.writeFileSync(path.join(__dirname, 'step2_labor.xlsx'), await zip2.generateAsync({ type: 'nodebuffer' }));

  // 步驟 3: 步驟 2 + 勞退 (sheet5.xml)
  const zip3 = await JSZip.loadAsync(await zip2.generateAsync({ type: 'nodebuffer' }));
  let s5 = await zip3.file('xl/worksheets/sheet5.xml').async('text');
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
  zip3.file('xl/worksheets/sheet5.xml', s5);
  fs.writeFileSync(path.join(__dirname, 'step3_pension.xlsx'), await zip3.generateAsync({ type: 'nodebuffer' }));

  // 步驟 4: 步驟 3 + 健保 (sheet3.xml)
  const zip4 = await JSZip.loadAsync(await zip3.generateAsync({ type: 'nodebuffer' }));
  let s3 = await zip4.file('xl/worksheets/sheet3.xml').async('text');
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
  zip4.file('xl/worksheets/sheet3.xml', s3);
  fs.writeFileSync(path.join(__dirname, 'step4_health.xlsx'), await zip4.generateAsync({ type: 'nodebuffer' }));

  console.log('Generated step 1, 2, 3, 4');
}

runGradual().catch(console.error);
