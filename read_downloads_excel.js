const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取下载文件夹中的Excel文件
const downloadPath = 'C:\\Users\\zhaoyang26\\Downloads';
const files = fs.readdirSync(downloadPath).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls')).slice(0, 20);

console.log('=== 下载文件夹Excel读取（前20个）===\n');

let totalRows = 0;

files.forEach((f, i) => {
    const fPath = path.join(downloadPath, f);
    try {
        const wb = XLSX.readFile(fPath);
        let fileRows = 0;
        wb.SheetNames.forEach(sheetName => {
            const sheet = wb.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            fileRows += data.length;
        });
        totalRows += fileRows;
        console.log('【' + (i+1) + '】' + f);
        console.log('  Sheets: ' + wb.SheetNames.join(', '));
        console.log('  总行数: ' + fileRows);
        console.log('');
    } catch (e) {
        console.log('【' + (i+1) + '】' + f + ' - Error: ' + e.message + '\n');
    }
});

console.log('前20个Excel总行数：' + totalRows);
