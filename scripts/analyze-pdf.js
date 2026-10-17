const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const fs = require('fs');
const path = require('path');

async function extractPdfText(filePath) {
  const data = fs.readFileSync(filePath);
  const uint8 = new Uint8Array(data);
  const loadingTask = pdfjsLib.getDocument({ data: uint8 });
  const pdf = await loadingTask.promise;
  
  let output = `\n${'='.repeat(70)}\n`;
  output += `PDF檔案: ${path.basename(filePath)}\n`;
  output += `頁數: ${pdf.numPages}\n`;
  output += `${'='.repeat(70)}\n`;
  
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    
    output += `\n--- 第 ${i} 頁 ---\n`;
    
    // 取得每個 text item 的座標與內容
    const items = textContent.items;
    
    // 依Y座標分群（同一行）再依X排序
    const lines = {};
    items.forEach(item => {
      const y = Math.round(item.transform[5]); // Y座標
      const x = item.transform[4]; // X座標
      if (!lines[y]) lines[y] = [];
      lines[y].push({ x, str: item.str });
    });
    
    // 按Y座標從大到小（PDF座標系Y軸朝上）
    const sortedYs = Object.keys(lines).map(Number).sort((a, b) => b - a);
    
    for (const y of sortedYs) {
      const lineItems = lines[y].sort((a, b) => a.x - b.x);
      const lineText = lineItems.map(i => i.str).join(' ').trim();
      if (lineText) {
        output += `[Y=${y}] ${lineText}\n`;
      }
    }
  }
  
  return output;
}

async function main() {
  const samplesDir = path.join(__dirname, '..', 'samples');
  const files = fs.readdirSync(samplesDir).filter(f => f.startsWith('cost_') && f.endsWith('.pdf'));
  
  let allOutput = '=== PDF 樣本文字內容分析 ===\n';
  allOutput += `分析時間: ${new Date().toISOString()}\n`;
  allOutput += `檔案數量: ${files.length}\n`;
  
  for (const file of files) {
    const filePath = path.join(samplesDir, file);
    allOutput += await extractPdfText(filePath);
  }
  
  const outPath = path.join(__dirname, 'pdf-analysis.txt');
  fs.writeFileSync(outPath, allOutput, 'utf-8');
  console.log('PDF 分析完成，已寫入:', outPath);
}

main().catch(err => {
  console.error('分析失敗:', err);
  process.exit(1);
});
