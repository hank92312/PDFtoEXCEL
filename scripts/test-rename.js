const fs = require('fs');
const JSZip = require('jszip');

function sanitizeSheetName(name, fallback) {
  if (!name || typeof name !== 'string') return fallback;
  // Excel 禁止字元: \ / ? * : [ ]
  let clean = name.replace(/[\\/?*:[\]]/g, '_').trim();
  if (clean.length > 31) clean = clean.substring(0, 31);
  return clean || fallback;
}

async function testRename() {
  const zip = await JSZip.loadAsync(fs.readFileSync('cost.xlsx'));
  let wbXml = await zip.file('xl/workbook.xml').async('text');

  const systemSheets = new Set(['使用說明', '員工設定', '健保', '勞保', '勞退', '全員工彙總']);
  
  // 模擬 7 位真實員工
  const employees = [
    { code: 'E01', name: '陳萱蘋' },
    { code: 'E02', name: '林小美' },
    { code: 'E03', name: '張志豪' },
    { code: 'E04', name: '李佳穎' },
    { code: 'E05', name: '王大同' },
    { code: 'E06', name: '周杰倫' },
    { code: 'E07', name: '蔡依林' },
    // E08~E20 保持空白
  ];

  let empIdx = 0;
  const usedNames = new Set(systemSheets);

  wbXml = wbXml.replace(/(<sheet\s+[^>]*name=")([^"]+)("([^>]*>))/g, (full, p1, oldName, p3) => {
    if (systemSheets.has(oldName)) {
      return full;
    }

    const currentEmp = employees[empIdx];
    empIdx++;

    let newName = oldName;
    if (currentEmp && currentEmp.name && !currentEmp.name.startsWith('員工')) {
      newName = sanitizeSheetName(currentEmp.name, oldName);
    }

    // 防止同名衝突
    let finalName = newName;
    let count = 2;
    while (usedNames.has(finalName)) {
      finalName = `${newName} (${count++})`;
    }
    usedNames.add(finalName);

    return `${p1}${finalName}${p3}`;
  });

  zip.file('xl/workbook.xml', wbXml);

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync('cost_test_renamed.xlsx', outBuf);
  console.log('Successfully created cost_test_renamed.xlsx');
}

testRename().catch(console.error);
