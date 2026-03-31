const fs = require('fs');
const path = require('path');

// 读取薪酬费用发票2026年的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\薪酬费用发票2026年';
const folders = fs.readdirSync(base);

console.log('=== 薪酬费用发票2026年详细 ===');
folders.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.statSync(fPath).isDirectory()) {
        const files = fs.readdirSync(fPath);
        console.log(f + '/ (' + files.length + '个文件)');
        files.forEach(file => console.log('  - ' + file));
    }
});
