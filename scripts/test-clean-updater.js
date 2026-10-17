const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

/**
 * 輔助函式：在 XML 的特定儲存格設定數值 (t="n")
 */
function setNumericCell(xml, cellRef, numVal) {
  // 如果儲存格已存在: <c r="C91" ...>...</c> 或 <c r="C91" .../>
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*>)(.*?)(<\\/c>)|<c\\s+r="${cellRef}"([^>]*?)\\/>`);
  if (cellRegex.test(xml)) {
    return xml.replace(cellRegex, (match, openTag, inner, closeTag, selfCloseAttrs) => {
      if (openTag) {
        // 保留 s 樣式屬性
        return `${openTag}<v>${numVal}</v>${closeTag}`;
      } else {
        return `<c r="${cellRef}"${selfCloseAttrs}><v>${numVal}</v></c>`;
      }
    });
  } else {
    // 若儲存格不存在，找到該 row 插入 (本表儲存格均已預先建立)
    return xml;
  }
}

/**
 * 輔助函式：在 XML 的特定儲存格設定文字 (t="inlineStr")
 */
function setTextCell(xml, cellRef, textVal) {
  const cellRegex = new RegExp(`(<c\\s+r="${cellRef}"[^>]*>)(.*?)(<\\/c>)|<c\\s+r="${cellRef}"([^>]*?)\\/>`);
  const escapedText = textVal.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return xml.replace(cellRegex, (match, openTag, inner, closeTag, selfCloseAttrs) => {
    if (openTag) {
      return `${openTag}<is><t>${escapedText}</t></is>${closeTag}`;
    } else {
      return `<c r="${cellRef}"${selfCloseAttrs} t="inlineStr"><is><t>${escapedText}</t></is></c>`;
    }
  });
}

async function runCleanUpdaterTest() {
  const templatePath = path.join(__dirname, '..', 'cost.xlsx');
  const buffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buffer);

  // 1. 更新員工設定 (sheet2.xml)
  let sheet2Xml = await zip.file('xl/worksheets/sheet2.xml').async('text');
  const empNames = ['吳金燕', '易金香', '古力予', '鄭百翔', '陳萱蘋', '歐明安', '林家羽', '員工08', '員工09', '員工10'];
  empNames.forEach((name, idx) => {
    const row = 4 + idx;
    sheet2Xml = setTextCell(sheet2Xml, `B${row}`, name);
  });
  zip.file('xl/worksheets/sheet2.xml', sheet2Xml);
  console.log('Updated sheet2.xml (員工設定)');

  // 2. 8 月份在 cost.xlsx 中的行號：
  // 摘要：Row 11 (8月)
  // 明細：Row 91 ~ 100 (8月 10 位員工)
  // 勞保 (sheet4.xml)
  let sheet4Xml = await zip.file('xl/worksheets/sheet4.xml').async('text');
  sheet4Xml = setNumericCell(sheet4Xml, 'B11', 6453); // 摘要總額
  sheet4Xml = setTextCell(sheet4Xml, 'D11', '已核對'); // 摘要狀態
  // 寫入明細
  const laborData = [
    { p: 870, c: 3087 }, // 吳金燕
    { p: 0, c: 40 },     // 易金香
    { p: 0, c: 50 },     // 古力予
    { p: 922, c: 3276 }, // 鄭百翔
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
  ];
  laborData.forEach((d, idx) => {
    const row = 91 + idx;
    sheet4Xml = setNumericCell(sheet4Xml, `C${row}`, d.p);
    sheet4Xml = setNumericCell(sheet4Xml, `D${row}`, d.c);
  });
  zip.file('xl/worksheets/sheet4.xml', sheet4Xml);
  console.log('Updated sheet4.xml (勞保)');

  // 3. 勞退 (sheet5.xml)
  let sheet5Xml = await zip.file('xl/worksheets/sheet5.xml').async('text');
  sheet5Xml = setNumericCell(sheet5Xml, 'B11', 6606);
  sheet5Xml = setTextCell(sheet5Xml, 'D11', '已核對');
  const pensionData = [
    { employer: 2088, personal: 0 }, // 吳金燕
    { employer: 1998, personal: 0 }, // 易金香
    { employer: 2520, personal: 0 }, // 古力予
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
    { employer: 0, personal: 0 },
  ];
  pensionData.forEach((d, idx) => {
    const row = 91 + idx;
    sheet5Xml = setNumericCell(sheet5Xml, `C${row}`, d.employer); // 雇主提繳
    sheet5Xml = setNumericCell(sheet5Xml, `D${row}`, d.personal); // 個人自提
  });
  zip.file('xl/worksheets/sheet5.xml', sheet5Xml);
  console.log('Updated sheet5.xml (勞退)');

  // 4. 健保 (sheet3.xml)
  let sheet3Xml = await zip.file('xl/worksheets/sheet3.xml').async('text');
  sheet3Xml = setNumericCell(sheet3Xml, 'B11', 7630);
  sheet3Xml = setTextCell(sheet3Xml, 'D11', '已核對');
  const healthData = [
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 916, c: 1428 },  // 陳萱蘋
    { p: 516, c: 1611 },  // 歐明安
    { p: 1548, c: 1611 }, // 林家羽
    { p: 0, c: 0 },
    { p: 0, c: 0 },
    { p: 0, c: 0 },
  ];
  healthData.forEach((d, idx) => {
    const row = 91 + idx;
    sheet3Xml = setNumericCell(sheet3Xml, `C${row}`, d.p);
    sheet3Xml = setNumericCell(sheet3Xml, `D${row}`, d.c);
  });
  zip.file('xl/worksheets/sheet3.xml', sheet3Xml);
  console.log('Updated sheet3.xml (健保)');

  // 產出檔案
  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });

  const outPath = path.join(__dirname, 'test_clean_output.xlsx');
  fs.writeFileSync(outPath, outBuffer);
  console.log('Successfully saved to:', outPath);
}

runCleanUpdaterTest().catch(console.error);
