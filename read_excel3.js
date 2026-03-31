const XLSX = require('xlsx');

const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\2024年信息流-思维杨易.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\初中信息流过程结果数据.xlsx'
];

files.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('\n--- Sheet: ' + workbook.SheetNames[0] + ' ---');
        data.slice(0, 15).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                const cells = row.slice(0, 12).map(c => String(c).substring(0, 25));
                console.log('R' + i + ':', cells);
            }
        });
        console.log('\nTotal rows:', data.length);
        console.log('');
    } catch (e) {
        console.log('Error:', e.message);
    }
});
