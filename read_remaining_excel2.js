const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取剩余Excel - 批次2
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\0323信息流分析\\线索明细数据 (14).xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\2024年信息流-思维杨易.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\线索明细数据 0205使用.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\小学信息流用户信息.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\EM业务线\\信息流分析-头条&广点通-11&12月.xlsx'
];

console.log('=== 读取剩余Excel - 批次2 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            const fileName = path.basename(f);
            console.log('【' + (i+1) + '】' + fileName);
            console.log('  Sheets: ' + wb.SheetNames.slice(0, 5).join(', '));
            if (wb.SheetNames.length > 5) {
                console.log('  ... 还有' + (wb.SheetNames.length - 5) + '个Sheet');
            }
            let totalRows = 0;
            wb.SheetNames.forEach(sheetName => {
                const sheet = wb.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                totalRows += data.length;
            });
            console.log('  总行数: ' + totalRows);
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
