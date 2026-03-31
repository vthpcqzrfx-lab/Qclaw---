const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取下载文件夹中的Excel文件 - 最终批次（301-370）
const downloadPath = 'C:\\Users\\zhaoyang26\\Downloads';
const allExcelFiles = fs.readdirSync(downloadPath).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
const files = allExcelFiles.slice(300);

console.log('=== 下载文件夹Excel读取 - 最终批次（301-' + allExcelFiles.length + '）===\n');

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
            console.log('【' + (i+301) + '】' + f + ' ⭐');
            console.log('  Sheets: ' + wb.SheetNames.slice(0, 5).join(', ') + (wb.SheetNames.length > 5 ? '...' : ''));
            console.log('  总行数: ' + fileRows.toLocaleString());
            console.log('');
        }
    } catch (e) {
        // 跳过错误文件
    }
});

console.log('最终批次成功读取：' + successCount + '个');
console.log('最终批次总行数：' + totalRows.toLocaleString());
console.log('');
console.log('下载文件夹Excel总计：' + allExcelFiles.length + '个');
