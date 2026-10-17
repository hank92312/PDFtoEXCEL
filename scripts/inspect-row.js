const fs = require('fs');
const o = fs.readFileSync('scripts/orig_zip/xl/worksheets/sheet4.xml', 'utf-8');

const m1 = o.match(/<row r="11"[^>]*>.*?<\/row>/);
console.log('Row 11 in sheet4:', m1 ? m1[0] : 'not found');

const m2 = o.match(/<row r="91"[^>]*>.*?<\/row>/);
console.log('Row 91 in sheet4:', m2 ? m2[0] : 'not found');

const empXml = fs.readFileSync('scripts/orig_zip/xl/worksheets/sheet2.xml', 'utf-8');
const r4 = empXml.match(/<row r="4"[^>]*>.*?<\/row>/);
console.log('Row 4 in sheet2 (員工設定):', r4 ? r4[0] : 'not found');
