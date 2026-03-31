const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取我的公文包中的Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包';
const excels = [
    '【供应商信息底表】20240102.xlsx',
    '上班的值不值测算版.xlsx',
    '创新增长组绩效-2023Q3.xlsx',
    '南二层工位图.xlsx',
    '增长数据总览.xlsx',
    '市场部工位安排.xlsx',
    '资产明细20220909.xlsx'
];

excels.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.existsSync(fPath)) {
        try {
            const workbook = XLSX.readFile(fPath);
            console.log('\n=== ' + f + ' ===');
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
    }
});
