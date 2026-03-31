const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取更多数据文件夹
const dataBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据';
const dataFiles = fs.readdirSync(dataBase).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls')).slice(0, 10);

console.log('=== 数据文件夹Excel文件（前10个）===');
dataFiles.forEach(f => console.log(f));

// 读取前3个数据文件
const fullPaths = dataFiles.slice(0, 3).map(f => path.join(dataBase, f));

fullPaths.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        console.log('Columns:', data[0] ? data[0].slice(0, 8) : 'N/A');
        data.slice(1, 5).forEach((row, i) => {
            console.log('R' + (i+1) + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
