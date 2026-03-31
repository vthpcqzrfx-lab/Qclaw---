const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取11-12月素养信息流数据的Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据';
const excels = [
    '信息流分析-头条&广点通-11&12月.xlsx',
    '信息流分析-头条&广点通-1月.xlsx',
    '初中信息流过程结果数据.xlsx'
];

excels.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.existsSync(fPath)) {
        try {
            const workbook = XLSX.readFile(fPath);
            console.log('\n=== ' + f + ' ===');
            console.log('Sheets:', workbook.SheetNames);
            
            workbook.SheetNames.slice(0, 3).forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                console.log('\n--- Sheet: ' + sheetName + ' (' + data.length + ' rows) ---');
                if (data.length > 0) {
                    console.log('Columns:', data[0].slice(0, 8));
                    data.slice(1, 6).forEach((row, i) => {
                        if (row.some(c => c !== '')) {
                            console.log('R' + (i+1) + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
                        }
                    });
                }
            });
        } catch (e) {
            console.log('Error reading ' + f + ':', e.message);
        }
    }
});
