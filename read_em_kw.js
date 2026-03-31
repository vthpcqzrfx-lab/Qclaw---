const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取EM业务线的Excel文件
const emBase = 'C:\\Users\\zhaoyang26\\Desktop\\EM业务线';
const emExcels = ['市场素材数量.xlsx'];

emExcels.forEach(f => {
    const fPath = path.join(emBase, f);
    if (fs.existsSync(fPath)) {
        try {
            const workbook = XLSX.readFile(fPath);
            console.log('\n=== ' + f + ' ===');
            console.log('Sheets:', workbook.SheetNames);
            workbook.SheetNames.slice(0, 2).forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                console.log('\n--- Sheet: ' + sheetName + ' (' + data.length + ' rows) ---');
                data.slice(0, 10).forEach((row, i) => {
                    if (row.some(c => c !== '')) {
                        console.log('R' + i + ':', row.slice(0, 6).map(c => String(c).substring(0, 25)));
                    }
                });
            });
        } catch (e) {
            console.log('Error:', e.message);
        }
    }
});

// 读取关键词的Excel文件
const kwBase = 'C:\\Users\\zhaoyang26\\Desktop\\关键词';
const kwExcels = fs.readdirSync(kwBase).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls') || f.endsWith('.csv'));

console.log('\n=== 关键词文件 ===');
kwExcels.forEach(f => {
    const fPath = path.join(kwBase, f);
    try {
        const workbook = XLSX.readFile(fPath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log(f + ': ' + data.length + '行');
    } catch (e) {
        console.log(f + ': Error - ' + e.message);
    }
});
