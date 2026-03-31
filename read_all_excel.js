const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 递归读取所有Excel文件
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
let excelCount = 0;
let totalRows = 0;

function findAndReadExcels(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findAndReadExcels(fPath, depth + 1);
        } else if ((f.endsWith('.xlsx') || f.endsWith('.xls')) && !f.startsWith('~')) {
            excelCount++;
            try {
                const wb = XLSX.readFile(fPath);
                wb.SheetNames.forEach(sheetName => {
                    const sheet = wb.Sheets[sheetName];
                    const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                    totalRows += data.length;
                });
            } catch (e) {
                // 跳过有密码或损坏的文件
            }
        }
    });
}

console.log('=== 开始统计所有Excel文件 ===');
findAndReadExcels(desktop);
console.log('Excel文件总数：' + excelCount);
console.log('总数据行数：' + totalRows);
