const XLSX = require('xlsx');

const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\图书产品部-24年目标拆解-私域V3版.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\星火-素养 历史数据.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\EM业务线\\EM小学预算.xlsx'
];

files.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('\n--- Sheet: ' + sheetName + ' ---');
            data.slice(0, 8).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 8).map(c => String(c).substring(0, 20)));
                }
            });
        });
        console.log('');
    } catch (e) {
        console.log('Error:', e.message);
    }
});
