const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次1
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\私域gmv_数分新口径_2023-02-01.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\微信集群统计-市场部私域.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\集团私域3月份数据-20230427.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\学科课结算金额（截止8.27）.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\X业务线-图书产品部经营数据_8月_-_副本.xlsx'
];

console.log('=== 批量读取Excel - 批次1 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            const fileName = path.basename(f);
            console.log('【' + (i+1) + '】' + fileName);
            console.log('  Sheets: ' + wb.SheetNames.join(', '));
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('  行数: ' + data.length);
            if (data.length > 0) {
                console.log('  列名: ' + data[0].slice(0, 5).join(', '));
            }
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
