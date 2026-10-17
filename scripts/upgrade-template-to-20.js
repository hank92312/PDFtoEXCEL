const fs = require('fs');
const JSZip = require('jszip');

async function upgradeCostTemplate() {
  console.log('--- 開始精確升級 cost.xlsx 範本 ---');

  // 從乾淨的備份載入
  const buf = fs.readFileSync('cost_backup_10emp.xlsx');
  const zip = await JSZip.loadAsync(buf);

  // 1. 升級「員工設定」 (sheet2.xml) 至 E01~E20 (Row 4 ~ Row 23)
  let sheet2Xml = await zip.file('xl/worksheets/sheet2.xml').async('text');
  let empRowsXml = '';
  for (let i = 1; i <= 20; i++) {
    const row = 3 + i;
    const code = `E${String(i).padStart(2, '0')}`;
    const defaultName = `員工${String(i).padStart(2, '0')}`;
    empRowsXml += `<row r="${row}"><c r="A${row}" s="5" t="inlineStr"><is><t>${code}</t></is></c><c r="B${row}" s="6" t="inlineStr"><is><t>${defaultName}</t></is></c><c r="C${row}" s="5" t="inlineStr"><is><t>請將預設名稱改成實際姓名</t></is></c></row>`;
  }

  sheet2Xml = sheet2Xml.replace(/<row r="4"[\s\S]*<\/row>/, empRowsXml + `<row r="25"><c r="A25" s="7" t="inlineStr"><is><t>注意</t></is></c><c r="B25" s="8" t="inlineStr"><is><t>更改姓名後，其他分頁會自動連動；請勿刪除員工列或更改員工代號。</t></is></c></row>`);
  sheet2Xml = sheet2Xml.replace(/<dimension ref="[^"]+"/, '<dimension ref="A1:C25"');
  zip.file('xl/worksheets/sheet2.xml', sheet2Xml);
  console.log('✅ 員工設定已擴充至 E01~E20');

  // 2. 升級「健保」、「勞保」、「勞退」明細列至 21~260
  const origSheet4 = await zip.file('xl/worksheets/sheet4.xml').async('text');
  const monthSerials = [];
  for (let m = 1; m <= 12; m++) {
    const summaryRow = 3 + m;
    const match = origSheet4.match(new RegExp(`<c r="A${summaryRow}"[^>]*><v>(\\d+)<\\/v><\\/c>`));
    monthSerials.push(match ? match[1] : '46235');
  }

  const upgradeDetailSheet = (xml) => {
    const row20Index = xml.indexOf('<row r="21"');
    const headerPart = xml.substring(0, row20Index);
    const endPart = xml.substring(xml.lastIndexOf('</row>') + 6);

    let detailRowsXml = '';
    let curRow = 21;
    for (let m = 0; m < 12; m++) {
      const serial = monthSerials[m];
      for (let empIdx = 1; empIdx <= 20; empIdx++) {
        const empRow = 3 + empIdx;
        detailRowsXml += `<row r="${curRow}"><c r="A${curRow}" s="10" t="n"><v>${serial}</v></c><c r="B${curRow}" s="5"><f>'&#21729;&#24037;&#35373;&#23450;'!$B$${empRow}</f><v></v></c><c r="C${curRow}" s="11" t="n"></c><c r="D${curRow}" s="11" t="n"></c></row>`;
        curRow++;
      }
    }

    let updated = headerPart + detailRowsXml + endPart;
    updated = updated.replace(/<dimension ref="[^"]+"/, '<dimension ref="A1:E260"');
    updated = updated.replace(/ref="A20:D140"/g, 'ref="A20:D260"');
    return updated;
  };

  zip.file('xl/worksheets/sheet3.xml', upgradeDetailSheet(await zip.file('xl/worksheets/sheet3.xml').async('text')));
  zip.file('xl/worksheets/sheet4.xml', upgradeDetailSheet(await zip.file('xl/worksheets/sheet4.xml').async('text')));
  zip.file('xl/worksheets/sheet5.xml', upgradeDetailSheet(await zip.file('xl/worksheets/sheet5.xml').async('text')));
  console.log('✅ 健保、勞保、勞退明細列擴充至 260 列');

  // 3. 更新現有員工工作表公式範圍 (sheet6 ~ sheet15)
  for (let s = 6; s <= 15; s++) {
    const p = `xl/worksheets/sheet${s}.xml`;
    let sXml = await zip.file(p).async('text');
    sXml = sXml.replace(/\$21:\$140/g, '$21:$260');
    zip.file(p, sXml);
  }

  // 4. 新增 員工11 ~ 員工20 (sheet16.xml ~ sheet25.xml)
  const templateEmpXml = await zip.file('xl/worksheets/sheet6.xml').async('text');
  for (let empIdx = 11; empIdx <= 20; empIdx++) {
    const sheetNum = 5 + empIdx;
    const path = `xl/worksheets/sheet${sheetNum}.xml`;
    const empRow = 3 + empIdx;
    let newEmpXml = templateEmpXml;
    newEmpXml = newEmpXml.replace(/<c r="B2"[^>]*><f[^>]*>[^<]+<\/f>/, `<c r="B2" s="3"><f>'&#21729;&#24037;&#35373;&#23450;'!B${empRow}</f>`);
    newEmpXml = newEmpXml.replace(/<c r="E2"[^>]*><f[^>]*>[^<]+<\/f>/, `<c r="E2" s="4"><f>'&#21729;&#24037;&#35373;&#23450;'!A${empRow}</f>`);
    zip.file(path, newEmpXml);
  }
  console.log('✅ 員工11 ~ 員工20 工作表已生成');

  // 5. 建立「全員工彙總」工作表 (sheet26.xml)
  let summarySheetXml = templateEmpXml;
  summarySheetXml = summarySheetXml.replace(/<c r="A1"[^>]*>[\s\S]*?<\/c>/, '<c r="A1" s="2" t="inlineStr"><is><t>全員工薪資與人事成本彙總</t></is></c>');
  summarySheetXml = summarySheetXml.replace(/<c r="B2"[^>]*>[\s\S]*?<\/c>/, '<c r="B2" s="3" t="inlineStr"><is><t>公司全體員工</t></is></c>');
  summarySheetXml = summarySheetXml.replace(/<c r="E2"[^>]*>[\s\S]*?<\/c>/, '<c r="E2" s="4" t="inlineStr"><is><t>ALL</t></is></c>');

  for (let m = 1; m <= 12; m++) {
    const row = 4 + m;
    const sumRow = 3 + m;
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="E${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="E${row}" s="16"><f>&#20581;&#20445;!$B$${sumRow}</f><v></v></c>`);
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="F${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="F${row}" s="16"><f>&#21214;&#20445;!$B$${sumRow}</f><v></v></c>`);
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="G${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="G${row}" s="16"><f>&#21214;&#36864;!$B$${sumRow}</f><v></v></c>`);
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="H${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="H${row}" s="16"><f>SUMIFS(&#20581;&#20445;!$C$21:$C$260,&#20581;&#20445;!$A$21:$A$260,$A${row})</f><v></v></c>`);
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="I${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="I${row}" s="16"><f>SUMIFS(&#21214;&#20445;!$C$21:$C$260,&#21214;&#20445;!$A$21:$A$260,$A${row})</f><v></v></c>`);
    summarySheetXml = summarySheetXml.replace(new RegExp(`<c r="J${row}"[^>]*><f[^>]*>[^<]+<\\/f><v><\\/v><\\/c>`), `<c r="J${row}" s="16"><f>SUMIFS(&#21214;&#36864;!$D$21:$D$260,&#21214;&#36864;!$A$21:$A$260,$A${row})</f><v></v></c>`);
  }
  zip.file('xl/worksheets/sheet26.xml', summarySheetXml);
  console.log('✅ 全員工彙總工作表 (sheet26) 已生成');

  // 6. 正確配置 relationships 與 workbook.xml (避免 rId 衝突)
  // 原有: rId1~rId15 (sheet1~15), rId16 (styles), rId17 (theme)
  // 新增: rId18~rId27 (sheet16~25), rId28 (sheet26:全員工彙總)
  let relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');
  let relsEntries = '';
  for (let empIdx = 11; empIdx <= 20; empIdx++) {
    const sId = 5 + empIdx; // 16 ~ 25
    const rId = 7 + empIdx; // 18 ~ 27
    relsEntries += `<Relationship Id="rId${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet${sId}.xml"/>`;
  }
  // 全員工彙總: sheet26, rId28
  relsEntries += `<Relationship Id="rId28" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet26.xml"/>`;
  relsXml = relsXml.replace('</Relationships>', relsEntries + '</Relationships>');
  zip.file('xl/_rels/workbook.xml.rels', relsXml);

  // 構造 workbook.xml
  let wbXml = await zip.file('xl/workbook.xml').async('text');
  let newSheetsXml = '<sheets>';
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="使用說明" sheetId="1" state="visible" r:id="rId1"/>';
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="員工設定" sheetId="2" state="visible" r:id="rId2"/>';
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="健保" sheetId="3" state="visible" r:id="rId3"/>';
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="勞保" sheetId="4" state="visible" r:id="rId4"/>';
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="勞退" sheetId="5" state="visible" r:id="rId5"/>';
  // 插入「全員工彙總」
  newSheetsXml += '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="全員工彙總" sheetId="26" state="visible" r:id="rId28"/>';
  
  // 員工01 ~ 員工10 (rId6 ~ rId15)
  for (let i = 1; i <= 10; i++) {
    const sId = 5 + i;
    const name = `員工${String(i).padStart(2, '0')}`;
    newSheetsXml += `<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${name}" sheetId="${sId}" state="visible" r:id="rId${sId}"/>`;
  }
  // 員工11 ~ 員工20 (rId18 ~ rId27)
  for (let i = 11; i <= 20; i++) {
    const sId = 5 + i;
    const rId = 7 + i;
    const name = `員工${String(i).padStart(2, '0')}`;
    newSheetsXml += `<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${name}" sheetId="${sId}" state="visible" r:id="rId${rId}"/>`;
  }
  newSheetsXml += '</sheets>';

  wbXml = wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/, newSheetsXml);
  // 清理有問題的 definedNames（避免與 sheet 順序與名稱不符）
  wbXml = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, '');
  zip.file('xl/workbook.xml', wbXml);

  // [Content_Types].xml
  let ctXml = await zip.file('[Content_Types].xml').async('text');
  let ctEntries = '';
  for (let i = 16; i <= 26; i++) {
    const part = `/xl/worksheets/sheet${i}.xml`;
    if (!ctXml.includes(part)) {
      ctEntries += `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }
  }
  ctXml = ctXml.replace('</Types>', ctEntries + '</Types>');
  zip.file('[Content_Types].xml', ctXml);

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('cost.xlsx', outBuf);
  console.log('🎉 cost.xlsx 範本升級完成！');
}

upgradeCostTemplate().catch(console.error);
