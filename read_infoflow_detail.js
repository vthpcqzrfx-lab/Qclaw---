const fs = require('fs');
const path = require('path');

// 读取11-12月素养信息流数据的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据';
const files = fs.readdirSync(base);

console.log('=== 11-12月素养信息流数据详细 ===');
files.forEach(f => {
    const fPath = path.join(base, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.slice(0, 5).forEach(s => console.log('  - ' + s));
        if (sub.length > 5) console.log('  ... 还有' + (sub.length - 5) + '个文件');
    } else {
        const size = (stat.size / 1024 / 1024).toFixed(2);
        console.log(f + ' (' + size + ' MB)');
    }
});

// 读取0323信息流分析
const infoBase = 'C:\\Users\\zhaoyang26\\Desktop\\0323信息流分析';
const infoFiles = fs.readdirSync(infoBase);
console.log('\n=== 0323信息流分析 ===');
infoFiles.forEach(f => {
    const fPath = path.join(infoBase, f);
    const stat = fs.statSync(fPath);
    const size = (stat.size / 1024 / 1024).toFixed(2);
    console.log(f + ' (' + size + ' MB)');
});
