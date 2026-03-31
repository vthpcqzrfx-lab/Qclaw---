const XLSX = require('xlsx');

// 读取关键词Excel
const keywordFiles = [
    'C:\\Users\\zhaoyang26\\Desktop\\关键词\\小学网红老师词.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\关键词\\小学英语竞品词.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\关键词\\高途竞品词.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\关键词\\【春二轮&春暑秋】-城市P8等级表.xlsx'
];

keywordFiles.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n=== ' + f.split('\\').pop() + ' ===');
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        console.log('Columns:', data[0] ? data[0].slice(0, 10) : 'N/A');
        data.slice(1, 8).forEach((row, i) => {
            console.log('R' + (i+1) + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
