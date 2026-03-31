const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取信息流数据分享的Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\信息流数据分享';
const excels = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$'));

excels.forEach(f => {
    try {
        const workbook = XLSX.readFile(path.join(base, f));
        console.log('\n=== ' + f + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        data.slice(0, 6).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                console.log('R' + i + ':', row.slice(0, 6).map(c => String(c).substring(0, 25)));
            }
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});

// 读取20240910X业务线交接文件夹
const handoverBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接';
const handoverFiles = fs.readdirSync(handoverBase);
console.log('\n=== 20240910X业务线交接 ===');
handoverFiles.forEach(f => {
    const fPath = path.join(handoverBase, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
    } else {
        console.log(f);
    }
});
