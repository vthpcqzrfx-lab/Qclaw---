const fs = require('fs');
const path = require('path');

// 读取我的公文包/数据文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据';
const files = fs.readdirSync(base);

console.log('=== 数据文件夹详细 ===');
files.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.statSync(fPath).isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
    } else {
        console.log(f);
    }
});
