const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 最终批次 - 读取剩余所有Excel的详细内容
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\微信私域1月业务与crm明细比对-20230223-1.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\微信私域1月业务与crm明细比对-20230223-第二版.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\微信私域1月业务与crm明细比对-20230224.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\给江珊-公司市场部私域素养订单2月.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\X业务线-图书产品部经营数据_8月_-_副本.xlsx'
];

console.log('=== 最终批次 - 详细读取Excel ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            const fileName = path.basename(f);
            console.log('【' + (i+1) + '】' + fileName);
            console.log('  Sheets: ' + wb.SheetNames.join(', '));
            
            let totalRows = 0;
            wb.SheetNames.forEach(sheetName => {
                const sheet = wb.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                totalRows += data.length;
                if (data.length > 0 && i === 0) {
                    console.log('  Sheet: ' + sheetName + ' - ' + data.length + '行');
                    console.log('  列名: ' + data[0].slice(0, 5).join(', '));
                }
            });
            console.log('  总行数: ' + totalRows);
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
