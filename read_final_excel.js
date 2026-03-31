const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 最终批次 - 读取剩余所有Excel
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
let excelFiles = [];

function findExcels(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findExcels(fPath, depth + 1);
        } else if ((f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~')) {
            excelFiles.push(fPath);
        }
    });
}

findExcels(desktop);

console.log('=== 最终Excel统计 ===');
console.log('总计Excel文件：' + excelFiles.length);

// 快速统计所有Excel
let totalSheets = 0;
let totalRows = 0;
let successCount = 0;

excelFiles.forEach(f => {
    try {
        const wb = XLSX.readFile(f);
        totalSheets += wb.SheetNames.length;
        wb.SheetNames.forEach(sheetName => {
            const sheet = wb.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            totalRows += data.length;
        });
        successCount++;
    } catch (e) {
        // 跳过有密码或损坏的文件
    }
});

console.log('成功读取：' + successCount + '个文件');
console.log('总Sheet数：' + totalSheets);
console.log('总数据行数：' + totalRows.toLocaleString());
