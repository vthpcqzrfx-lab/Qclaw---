const fs = require('fs');
const path = require('path');

// 读取薪酬福利费报销的子文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\薪酬福利费报销';
const folders = fs.readdirSync(base);

console.log('=== 薪酬福利费报销详细 ===');
folders.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.statSync(fPath).isDirectory()) {
        const files = fs.readdirSync(fPath);
        console.log('\n' + f + '/ (' + files.length + '个文件)');
        files.slice(0, 10).forEach(file => console.log('  - ' + file));
        if (files.length > 10) console.log('  ... 还有' + (files.length - 10) + '个文件');
    }
});
