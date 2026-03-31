const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取更多Excel - 批次22
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\订单留痕-语言1.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\订单留痕-高中.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\订单线索留痕-语言2.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\私域统计顾问-2301-2302-订单信息1.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\私域统计顾问-2301-2302-订单信息2.xlsx'
];

console.log('=== 批次22 - 订单数据 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            console.log('【' + (i+1) + '】' + path.basename(f));
            console.log('  Sheets: ' + wb.SheetNames.join(', '));
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('  行数: ' + data.length);
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
