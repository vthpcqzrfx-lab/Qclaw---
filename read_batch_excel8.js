const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次8（高途体系培训）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途体系培训';
const files = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls')).slice(0, 10);

console.log('=== 批量读取Excel - 批次8（高途体系培训）===\n');

files.forEach((f, i) => {
    const fPath = path.join(base, f);
    try {
        const wb = XLSX.readFile(fPath);
        console.log('【' + (i+1) + '】' + f);
        console.log('  Sheets: ' + wb.SheetNames.slice(0, 5).join(', '));
        if (wb.SheetNames.length > 5) {
            console.log('  ... 还有' + (wb.SheetNames.length - 5) + '个Sheet');
        }
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('  行数: ' + data.length);
        console.log('');
    } catch (e) {
        console.log('  Error: ' + e.message + '\n');
    }
});
