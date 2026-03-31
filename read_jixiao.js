const fs = require('fs');
const path = require('path');

// 读取绩效文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效';
const folders = fs.readdirSync(base);

console.log('=== 绩效文件夹详细 ===');
folders.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.statSync(fPath).isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.forEach(s => console.log('  - ' + s));
    } else {
        console.log(f);
    }
});
