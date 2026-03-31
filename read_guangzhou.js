const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取高途-KA广州经历的关键Excel
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KA广州经历';
const excels = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls') || f.endsWith('.et'));

console.log('=== 高途-KA广州经历Excel文件 ===');
excels.forEach(f => {
    const fPath = path.join(base, f);
    try {
        const workbook = XLSX.readFile(fPath);
        console.log('\n--- ' + f + ' ---');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        data.slice(0, 5).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                console.log('R' + i + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
            }
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
