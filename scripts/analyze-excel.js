const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function analyzeExcel() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(__dirname, '..', 'cost.xlsx'));

  let output = '';
  const log = (s) => { output += s + '\n'; };

  log('=== EXCEL WORKBOOK 基本資訊 ===');
  log('工作表數量: ' + workbook.worksheets.length);
  log('工作表名稱: ' + workbook.worksheets.map(s => s.name).join(', '));
  
  // 只分析關鍵的工作表
  const sheetsToAnalyze = ['使用說明', '員工設定', '健保', '勞保', '勞退', '員工01'];
  
  for (const ws of workbook.worksheets) {
    if (!sheetsToAnalyze.includes(ws.name)) continue;
    
    log(`\n${'='.repeat(60)}`);
    log(`工作表: "${ws.name}" | 行數: ${ws.rowCount} | 列數: ${ws.columnCount}`);
    log('='.repeat(60));
    
    // 合併儲存格
    if (ws._merges && Object.keys(ws._merges).length > 0) {
      log('\n合併儲存格:');
      for (const key of Object.keys(ws._merges)) {
        log('  ' + key);
      }
    }
    
    // 儲存格內容
    log('\n--- 儲存格內容 ---');
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const cells = [];
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colLetter = getColumnLetter(colNumber);
        let val = cell.value;
        
        if (val && typeof val === 'object') {
          if (val.formula) {
            val = `[公式: =${val.formula}] (結果: ${val.result})`;
          } else if (val.sharedFormula) {
            val = `[共用公式: ref ${val.sharedFormula}] (結果: ${val.result})`;
          } else if (val.richText) {
            val = val.richText.map(r => r.text).join('');
          } else if (val instanceof Date) {
            val = `[日期: ${val.getFullYear()}-${String(val.getMonth()+1).padStart(2,'0')}-${String(val.getDate()).padStart(2,'0')}]`;
          }
        }

        cells.push(`${colLetter}${rowNumber}=${JSON.stringify(val)}`);
      });
      
      if (cells.length > 0) {
        log(`Row ${String(rowNumber).padStart(3)}: ${cells.join(' | ')}`);
      }
    });
    
    // 欄寬
    log('\n--- 欄位寬度 ---');
    for (let c = 1; c <= ws.columnCount; c++) {
      const col = ws.getColumn(c);
      if (col.width) {
        log(`  ${getColumnLetter(c)}: width=${col.width}`);
      }
    }
    
    // 填滿顏色 (for first 5 rows with data)
    log('\n--- 儲存格填滿色彩 (前20行) ---');
    for (let r = 1; r <= Math.min(20, ws.rowCount); r++) {
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.fill && cell.fill.fgColor) {
          log(`  ${getColumnLetter(colNumber)}${r}: fill=${JSON.stringify(cell.fill.fgColor)}`);
        }
      });
    }
  }

  const outPath = path.join(__dirname, '..', 'scripts', 'excel-analysis.txt');
  fs.writeFileSync(outPath, output, 'utf-8');
  console.log('分析完成，已寫入:', outPath);
}

function getColumnLetter(colNumber) {
  let letter = '';
  while (colNumber > 0) {
    const remainder = (colNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    colNumber = Math.floor((colNumber - 1) / 26);
  }
  return letter;
}

analyzeExcel().catch(err => {
  console.error('分析失敗:', err);
  process.exit(1);
});
