const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次5（信息流数据）
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\信息流分析-头条&广点通-11&12月.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\信息流分析-头条&广点通-1月.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\小学信息流用户信息.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\EM业务线\\市场素材数量.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\EM业务线\\EM小学预算.xlsx'
];

console.log('=== 批量读取Excel - 批次5（信息流）===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            const fileName = path.basename(f);
            console.log('【' + (i+1) + '】' + fileName);
            console.log('  Sheets: ' + wb.SheetNames.slice(0, 8).join(', '));
            if (wb.SheetNames.length > 8) {
                console.log('  ... 还有' + (wb.SheetNames.length - 8) + '个Sheet');
            }
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('  行数: ' + data.length);
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
