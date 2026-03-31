const XLSX = require('xlsx');

const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\信息流分析-头条&广点通-1月.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\小学信息流用户信息.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\线索明细数据 0205使用.xlsx'
];

files.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('First 3 rows:');
        data.slice(0, 3).forEach((row, i) => console.log('Row' + i + ':', row.slice(0, 10)));
        console.log('');
    } catch (e) {
        console.log('Error reading ' + f + ':', e.message);
    }
});
