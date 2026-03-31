const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取X业务线文档的Excel
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档';
const excels = ['7日分销员特训营.xlsx', '9.1-11.30发货相关客诉统计.xlsx', '星火-素养 历史数据.xlsx'];

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
            data.slice(0, 5).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
                }
            });
        } catch (e) {
            console.log('Error:', e.message);
        }
    }
});
