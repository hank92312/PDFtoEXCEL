const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function testOneByOne() {
  const templatePath = path.join(__dirname, '..', 'cost.xlsx');
  const buffer = fs.readFileSync(templatePath);

  // 測試 A: 只更新 sheet2.xml (員工設定)
  {
    const zip = await JSZip.loadAsync(buffer);
    let s2 = await zip.file('xl/worksheets/sheet2.xml').async('text');
    // 把 <t>&#21729;&#24037;01</t> 改成 <t>吳金燕</t>
    s2 = s2.replace(/<t>&#21729;&#24037;01<\/t>/, '<t>吳金燕</t>');
    zip.file('xl/worksheets/sheet2.xml', s2);
    const out = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(path.join(__dirname, 'test_A_emp.xlsx'), out);
  }

  // 測試 B: 只更新 sheet4.xml (勞保) 的一個數字儲存格 B11
  {
    const zip = await JSZip.loadAsync(buffer);
    let s4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
    // B11 原本是 <c r="B11" s="11" t="n"></c>
    s4 = s4.replace('<c r="B11" s="11" t="n"></c>', '<c r="B11" s="11" t="n"><v>6453</v></c>');
    zip.file('xl/worksheets/sheet4.xml', s4);
    const out = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(path.join(__dirname, 'test_B_labor_summary.xlsx'), out);
  }

  // 測試 C: 只更新 sheet4.xml (勞保) 的 C91
  {
    const zip = await JSZip.loadAsync(buffer);
    let s4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
    // C91 原本是 <c r="C91" s="11" t="n"></c>
    s4 = s4.replace('<c r="C91" s="11" t="n"></c>', '<c r="C91" s="11" t="n"><v>870</v></c>');
    zip.file('xl/worksheets/sheet4.xml', s4);
    const out = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(path.join(__dirname, 'test_C_labor_c91.xlsx'), out);
  }

  // 測試 D: 只更新 D11 (文字 '已核對')
  {
    const zip = await JSZip.loadAsync(buffer);
    let s4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
    // D11 原本是 <c r="D11" s="6" t="n"></c>
    s4 = s4.replace('<c r="D11" s="6" t="n"></c>', '<c r="D11" s="6" t="inlineStr"><is><t>已核對</t></is></c>');
    zip.file('xl/worksheets/sheet4.xml', s4);
    const out = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(path.join(__dirname, 'test_D_labor_d11.xlsx'), out);
  }

  console.log('Generated test A, B, C, D');
}

testOneByOne().catch(console.error);
