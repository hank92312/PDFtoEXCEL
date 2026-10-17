const fs = require('fs');
const o = fs.readFileSync('scripts/orig_zip/xl/worksheets/sheet6.xml', 'utf-8');
const s = fs.readFileSync('scripts/saved_zip/xl/worksheets/sheet6.xml', 'utf-8');

console.log('--- ORIG E5 ---');
const m1 = o.match(/<c r="E5"[^>]*>.*?<\/c>/);
console.log(m1 ? m1[0] : 'not found');

console.log('--- SAVED E5 ---');
const m2 = s.match(/<c r="E5"[^>]*>.*?<\/c>/);
console.log(m2 ? m2[0] : 'not found');

console.log('--- ORIG B2 ---');
const b1 = o.match(/<c r="B2"[^>]*>.*?<\/c>/);
console.log(b1 ? b1[0] : 'not found');

console.log('--- SAVED B2 ---');
const b2 = s.match(/<c r="B2"[^>]*>.*?<\/c>/);
console.log(b2 ? b2[0] : 'not found');
