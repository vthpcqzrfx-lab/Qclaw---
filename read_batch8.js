const XLSX = require('xlsx');

const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KA中价课经历\\KA年课创新项目组需要-跨学部线索推荐-数据汇总（截至25.12.4）.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KM青少学部经历\\KM-青少学部跨学部线索推荐成单明细（截止25.6.11）.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-集团市场&X业务线经历\\【k12】爆款文案素材库.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-集团市场&X业务线经历\\星火-素养.xlsx'
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
