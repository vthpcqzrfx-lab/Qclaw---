const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取绩效Excel
const files = [
    { name: '2022伙伴民主互评表.xlsx', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效\\2022伙伴民主互评表.xlsx' },
    { name: '2022年年度绩效评定.xlsx', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效\\2022年年度绩效评定.xlsx' },
    { name: '创新增长组2025年Q1绩效打分.xls', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2025年绩效\\创新增长组2025年Q1绩效打分.xls' }
];

files.forEach(f => {
    if (fs.existsSync(f.path)) {
        try {
            const workbook = XLSX.readFile(f.path);
            console.log('\n=== ' + f.name + ' ===');
            console.log('Sheets:', workbook.SheetNames);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('Total rows:', data.length);
            data.slice(0, 8).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 6).map(c => String(c).substring(0, 20)));
                }
            });
        } catch (e) {
            console.log('Error:', e.message);
        }
    }
});
