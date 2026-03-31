const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取X业务线交接的关键Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接';
const excels = [
    'X业务线-图书产品部经营数据_8月_-_副2.xlsx',
    'X业务线-图书产品部经营数据_8月_-_副本.xlsx',
    '学科课结算金额（截止8.27）.xlsx',
    '李静定标逻辑.xls',
    '高途代理.xls',
    '工业化种草拆解V1.xlsx'
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
            data.slice(0, 6).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
                }
            });
        } catch (e) {
            console.log('Error reading ' + f + ':', e.message);
        }
    } else {
        console.log('File not found: ' + f);
    }
});
