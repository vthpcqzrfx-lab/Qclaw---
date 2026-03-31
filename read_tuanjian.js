const fs = require('fs');
const path = require('path');

// 读取团建文件夹的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\团建';
const files = fs.readdirSync(base);

console.log('=== 团建文件夹详细 ===');
files.forEach(f => {
    const fPath = path.join(base, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
    } else {
        console.log(f);
    }
});

// 统计文件类型
const types = {};
files.forEach(f => {
    const ext = path.extname(f).toLowerCase();
    types[ext] = (types[ext] || 0) + 1;
});

console.log('\n文件类型分布：');
Object.entries(types).forEach(([ext, count]) => {
    console.log(ext + ': ' + count);
});
