const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 连续读取 - 批次17（数据文件夹子文件夹）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据';
const folders = fs.readdirSync(base).filter(f => {
    const fPath = path.join(base, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 连续读取 - 批次17（数据子文件夹）===\n');

folders.forEach((f, i) => {
    const fPath = path.join(base, f);
    const files = fs.readdirSync(fPath);
    const excels = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'));
    console.log('【' + (i+1) + '】' + f + '/ (' + files.length + '个文件, ' + excels.length + '个Excel)');
    if (excels.length > 0) {
        excels.slice(0, 3).forEach(e => console.log('  - ' + e));
        if (excels.length > 3) console.log('  ... 还有' + (excels.length - 3) + '个');
    }
    console.log('');
});
