const fs = require('fs');
const path = require('path');

const origDir = path.join(__dirname, 'orig_zip');
const savedDir = path.join(__dirname, 'saved_zip');

console.log('=== 1. _rels/.rels ===');
console.log('--- ORIG ---');
console.log(fs.readFileSync(path.join(origDir, '_rels', '.rels'), 'utf-8'));
console.log('--- SAVED ---');
console.log(fs.readFileSync(path.join(savedDir, '_rels', '.rels'), 'utf-8'));

console.log('\n=== 2. xl/_rels/workbook.xml.rels ===');
console.log('--- ORIG ---');
console.log(fs.readFileSync(path.join(origDir, 'xl', '_rels', 'workbook.xml.rels'), 'utf-8'));
console.log('--- SAVED ---');
console.log(fs.readFileSync(path.join(savedDir, 'xl', '_rels', 'workbook.xml.rels'), 'utf-8'));

console.log('\n=== 3. [Content_Types].xml ===');
console.log('--- ORIG ---');
console.log(fs.readFileSync(path.join(origDir, '[Content_Types].xml'), 'utf-8'));
console.log('--- SAVED ---');
console.log(fs.readFileSync(path.join(savedDir, '[Content_Types].xml'), 'utf-8'));
