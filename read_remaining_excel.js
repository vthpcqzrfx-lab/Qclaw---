const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取剩余的所有Excel文件
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

console.log('=== 剩余Excel文件统计 ===');
console.log('总计Excel文件：' + excelFiles.length);

// 读取前20个未读的Excel
let readCount = 0;
let totalRows = 0;

excelFiles.slice(0, 30).forEach((f, i) => {
    try {
        const wb = XLSX.readFile(f);
        const fileName = path.basename(f);
        let fileRows = 0;
        wb.SheetNames.forEach(sheetName => {
            const sheet = wb.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            fileRows += data.length;
        });
        totalRows += fileRows;
        readCount++;
        if (i < 5) {
            console.log('【' + (i+1) + '】' + fileName + ' - ' + fileRows + '行');
        }
    } catch (e) {
        // 跳过错误
    }
});

console.log('');
console.log('本次读取：' + readCount + '个文件');
console.log('本次数据行数：' + totalRows);
