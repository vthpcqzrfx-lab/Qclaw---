const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 连续读取 - 批次16（合作方）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\合作方';
const folders = fs.readdirSync(base);

console.log('=== 连续读取 - 批次16（合作方）===\n');

folders.forEach((f, i) => {
    const fPath = path.join(base, f);
    const files = fs.readdirSync(fPath);
    const excels = files.filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'));
    console.log('【' + (i+1) + '】' + f + '/ (' + files.length + '个文件)');
    if (excels.length > 0) {
        excels.forEach(e => {
            const ePath = path.join(fPath, e);
            try {
                const wb = XLSX.readFile(ePath);
                console.log('  - ' + e + ' (' + wb.SheetNames.length + ' Sheets)');
            } catch (err) {
                console.log('  - ' + e + ' (Error)');
            }
        });
    }
    console.log('');
});
