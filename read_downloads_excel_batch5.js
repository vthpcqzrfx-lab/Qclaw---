const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取下载文件夹中的Excel文件 - 批次5（151-200）
const downloadPath = 'C:\\Users\\zhaoyang26\\Downloads';
const files = fs.readdirSync(downloadPath).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls')).slice(150, 200);

console.log('=== 下载文件夹Excel读取（151-200）===\n');

let totalRows = 0;
let successCount = 0;

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
        successCount++;
        
        if (fileRows > 1000) {
            console.log('【' + (i+151) + '】' + f + ' ⭐');
            console.log('  Sheets: ' + wb.SheetNames.slice(0, 5).join(', ') + (wb.SheetNames.length > 5 ? '...' : ''));
            console.log('  总行数: ' + fileRows.toLocaleString());
            console.log('');
        }
    } catch (e) {
        // 跳过错误文件
    }
});

console.log('批次5（151-200）成功读取：' + successCount + '个');
console.log('批次5总行数：' + totalRows.toLocaleString());
