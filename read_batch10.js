const XLSX = require('xlsx');

const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\私域gmv_数分新口径_2023-02-01.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\私域gmv_经分口径_2023-01-30.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\微信集群统计-市场部私域.xlsx'
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
            console.log('Columns:', data[0] ? data[0].slice(0, 8) : 'N/A');
            data.slice(1, 6).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + (i+1) + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
                }
            });
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
