const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取KM青少学部经历的子文件夹
const kmBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KM青少学部经历';
const kmFiles = fs.readdirSync(kmBase);
console.log('=== KM青少学部经历（全部文件）===');
kmFiles.forEach(f => {
    const fPath = path.join(kmBase, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.slice(0, 3).forEach(s => console.log('  - ' + s));
    } else {
        console.log(f);
    }
});

// 读取KM青少学部的Excel文件
const kmExcels = kmFiles.filter(f => f.endsWith('.xlsx') || f.endsWith('.xls') || f.endsWith('.et'));
console.log('\n=== KM青少学部Excel文件 ===');
kmExcels.forEach(f => {
    try {
        const workbook = XLSX.readFile(path.join(kmBase, f));
        console.log('\n--- ' + f + ' ---');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        console.log('Columns:', data[0] ? data[0].slice(0, 6) : 'N/A');
        data.slice(1, 5).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                console.log('R' + (i+1) + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
            }
        });
    } catch (e) {
        console.log('Error reading ' + f + ':', e.message);
    }
});
