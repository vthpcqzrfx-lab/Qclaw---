const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次3（绩效相关）
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效\\2022伙伴民主互评表.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效\\2022年年度绩效评定.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2025年绩效\\创新增长组2025年Q1绩效打分.xls',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\【主管级（含储备）管理能力打分表】-to赵杨-完成.xlsx'
];

console.log('=== 批量读取Excel - 批次3（绩效）===\n');

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
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
