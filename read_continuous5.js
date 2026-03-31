const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 连续读取 - 批次15（招聘文件夹）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\招聘';
const files = fs.readdirSync(base);

console.log('=== 连续读取 - 批次15（招聘文件夹）===\n');

files.forEach((f, i) => {
    const fPath = path.join(base, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const subFiles = fs.readdirSync(fPath);
        console.log('【' + (i+1) + '】' + f + '/ (' + subFiles.length + '个文件)');
        subFiles.forEach(s => console.log('  - ' + s));
    } else {
        console.log('【' + (i+1) + '】' + f);
    }
    console.log('');
});
