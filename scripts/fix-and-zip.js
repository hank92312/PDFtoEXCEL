const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// 複製一份 saved_zip 到 fixed_zip
const savedDir = path.join(__dirname, 'saved_zip');
const fixedDir = path.join(__dirname, 'fixed_zip');

if (fs.existsSync(fixedDir)) {
  fs.rmSync(fixedDir, { recursive: true, force: true });
}

// 遞迴複製
function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const item of fs.readdirSync(src)) {
    const s = path.join(src, item);
    const d = path.join(dest, item);
    if (fs.statSync(s).isDirectory()) copyDir(s, d);
    else fs.copyFileSync(s, d);
  }
}
copyDir(savedDir, fixedDir);

// 替換所有 xml 檔案中的 &apos; 為 '
// 特別是 <f>...</f> 標籤中，以及 workbook.xml 中的 definedNames
function fixXmlFiles(dir) {
  for (const item of fs.readdirSync(dir)) {
    const full = path.join(dir, item);
    if (fs.statSync(full).isDirectory()) {
      fixXmlFiles(full);
    } else if (item.endsWith('.xml')) {
      let content = fs.readFileSync(full, 'utf-8');
      if (content.includes('&apos;')) {
        content = content.replace(/&apos;/g, "'");
        fs.writeFileSync(full, content, 'utf-8');
        console.log('Fixed &apos; in:', item);
      }
    }
  }
}

fixXmlFiles(fixedDir);

// 同時修復 workbook.xml 中缺少 localSheetId 的 _xlnm._FilterDatabase (如果有的話)
const wbXmlPath = path.join(fixedDir, 'xl', 'workbook.xml');
let wbXml = fs.readFileSync(wbXmlPath, 'utf-8');
// 移除 _xlnm._FilterDatabase 避免衝突
wbXml = wbXml.replace(/<definedNames>.*?<\/definedNames>/s, '');
fs.writeFileSync(wbXmlPath, wbXml, 'utf-8');

console.log('Done preparing fixed_zip. Now compressing to test_fixed.xlsx...');

// 使用 PowerShell Compress-Archive 壓回 .xlsx
const zipPath = path.join(__dirname, 'test_fixed.zip');
const xlsxPath = path.join(__dirname, 'test_fixed.xlsx');
if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
if (fs.existsSync(xlsxPath)) fs.unlinkSync(xlsxPath);

execSync(`powershell -Command "Compress-Archive -Path D:\\PDFtoEXCEL\\scripts\\fixed_zip\\* -DestinationPath D:\\PDFtoEXCEL\\scripts\\test_fixed.zip -Force"`);
fs.renameSync(zipPath, xlsxPath);
console.log('Successfully created test_fixed.xlsx');
