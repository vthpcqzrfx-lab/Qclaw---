const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取图书部数据汇总的关键Excel
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\图书部数据汇总';
const excels = [
    '图书1-3月收款收入-20240416.xlsx',
    '图书KPI-销售部2024.xlsx',
    '业绩前20名单.xlsx',
    '分销员行为数据信息表_20240218141259-0.xls'
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
            console.log('Error:', e.message);
        }
    }
});
