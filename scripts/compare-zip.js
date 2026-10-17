const fs = require('fs');
const path = require('path');

function getFiles(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  list.forEach(file => {
    file = path.join(dir, file);
    const stat = fs.statSync(file);
    if (stat && stat.isDirectory()) {
      results = results.concat(getFiles(file));
    } else {
      results.push(file);
    }
  });
  return results;
}

const origDir = path.join(__dirname, 'orig_zip');
const savedDir = path.join(__dirname, 'saved_zip');

const origFiles = getFiles(origDir).map(f => path.relative(origDir, f));
const savedFiles = getFiles(savedDir).map(f => path.relative(savedDir, f));

console.log('--- 原始檔案存在但儲存後缺失的檔案 ---');
origFiles.forEach(f => {
  if (!savedFiles.includes(f)) console.log('-', f);
});

console.log('--- 儲存後多出的檔案 ---');
savedFiles.forEach(f => {
  if (!origFiles.includes(f)) console.log('+', f);
});

// 比對 [Content_Types].xml
console.log('\n--- [Content_Types].xml 比較 ---');
const ctOrig = fs.readFileSync(path.join(origDir, '[Content_Types].xml'), 'utf-8');
const ctSaved = fs.readFileSync(path.join(savedDir, '[Content_Types].xml'), 'utf-8');
console.log('Orig length:', ctOrig.length, 'Saved length:', ctSaved.length);

// 檢查 workbook.xml
const wbOrig = fs.readFileSync(path.join(origDir, 'xl', 'workbook.xml'), 'utf-8');
const wbSaved = fs.readFileSync(path.join(savedDir, 'xl', 'workbook.xml'), 'utf-8');
console.log('xl/workbook.xml Orig length:', wbOrig.length, 'Saved length:', wbSaved.length);
