const fs = require('fs');
const path = require('path');

// 读取个人文件夹的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\个人';
const files = fs.readdirSync(base);

console.log('=== 个人文件夹详细 ===');
files.forEach(f => {
    const fPath = path.join(base, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.slice(0, 3).forEach(s => console.log('  - ' + s));
    } else {
        console.log(f);
    }
});
