const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, 'saved_zip', 'xl', 'worksheets');
let count = 0;
for (const f of fs.readdirSync(dir)) {
  const content = fs.readFileSync(path.join(dir, f), 'utf-8');
  const matches = content.match(/&apos;/g);
  if (matches) {
    console.log(f + ' has ' + matches.length + ' &apos;');
    count += matches.length;
  }
}
console.log('Total &apos;:', count);
