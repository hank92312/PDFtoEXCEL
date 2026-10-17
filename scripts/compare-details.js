const fs = require('fs');
const path = require('path');

const origDir = path.join(__dirname, 'orig_zip', 'xl');
const savedDir = path.join(__dirname, 'saved_zip', 'xl');

const origSheets = fs.readdirSync(path.join(origDir, 'worksheets'));
console.log('Worksheet comparison:');
for (const s of origSheets) {
  const oSize = fs.statSync(path.join(origDir, 'worksheets', s)).size;
  const sPath = path.join(savedDir, 'worksheets', s);
  const sSize = fs.existsSync(sPath) ? fs.statSync(sPath).size : 'MISSING';
  console.log(`  ${s}: orig=${oSize}, saved=${sSize}`);
}

console.log('calcChain in orig:', fs.existsSync(path.join(origDir, 'calcChain.xml')));
console.log('calcChain in saved:', fs.existsSync(path.join(savedDir, 'calcChain.xml')));

// 檢查 [Content_Types].xml
console.log('\n--- 檢查 [Content_Types].xml ---');
const ctOrig = fs.readFileSync(path.join(__dirname, 'orig_zip', '[Content_Types].xml'), 'utf-8');
const ctSaved = fs.readFileSync(path.join(__dirname, 'saved_zip', '[Content_Types].xml'), 'utf-8');

// 檢查 calcChain 是否在 [Content_Types].xml 中
console.log('calcChain in orig [Content_Types]:', ctOrig.includes('calcChain'));
console.log('calcChain in saved [Content_Types]:', ctSaved.includes('calcChain'));

// 檢查 sharedStrings
console.log('sharedStrings in orig [Content_Types]:', ctOrig.includes('sharedStrings'));
console.log('sharedStrings in saved [Content_Types]:', ctSaved.includes('sharedStrings'));

// 檢查 sheet1.xml (使用說明)
const s1Orig = fs.readFileSync(path.join(origDir, 'worksheets', 'sheet1.xml'), 'utf-8');
const s1Saved = fs.readFileSync(path.join(savedDir, 'worksheets', 'sheet1.xml'), 'utf-8');
console.log('sheet1.xml orig header:', s1Orig.substring(0, 200));
console.log('sheet1.xml saved header:', s1Saved.substring(0, 200));
