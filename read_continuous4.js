const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 连续读取 - 批次14（1月份素材）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\1月份素材';
const folders = fs.readdirSync(base).filter(f => {
    const fPath = path.join(base, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 连续读取 - 批次14（1月份素材子文件夹）===\n');

folders.forEach((f, i) => {
    const fPath = path.join(base, f);
    const files = fs.readdirSync(fPath);
    const excels = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'));
    console.log('【' + (i+1) + '】' + f + '/');
    console.log('  总文件数: ' + files.length);
    console.log('  Excel文件数: ' + excels.length);
    if (excels.length > 0) {
        excels.forEach(e => console.log('    - ' + e));
    }
    console.log('');
});
