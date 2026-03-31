const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次10（薪酬福利费报销）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\薪酬福利费报销\\其他文件';
const files = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

console.log('=== 批量读取Excel - 批次10（薪酬福利费报销）===\n');

files.forEach((f, i) => {
    const fPath = path.join(base, f);
    try {
        const wb = XLSX.readFile(fPath);
        console.log('【' + (i+1) + '】' + f);
        console.log('  Sheets: ' + wb.SheetNames.join(', '));
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('  行数: ' + data.length);
        console.log('');
    } catch (e) {
        console.log('  Error: ' + e.message + '\n');
    }
});
