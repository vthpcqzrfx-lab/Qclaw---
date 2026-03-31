const XLSX = require('xlsx');

// 读取核心文件
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\私域运营组架构.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\增长数据总览.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\【供应商信息底表】20240102.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\上班的值不值测算版.xlsx'
];

files.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        
        workbook.SheetNames.slice(0, 2).forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('\n--- Sheet: ' + sheetName + ' (' + data.length + ' rows) ---');
            data.slice(0, 10).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 6).map(c => String(c).substring(0, 30)));
                }
            });
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
