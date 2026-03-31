const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 最终批次 - 读取剩余的重要Excel
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\图书产品部-24年目标拆解-私域V3版.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\图书产品部-24年目标拆解.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\1205-图书产品部-24年目标拆解.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\1205-图书产品部-24年目标拆解-私域.xlsx'
];

console.log('=== 最终批次 - 图书产品部目标文件 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            console.log('【' + (i+1) + '】' + path.basename(f));
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
    } else {
        console.log('【' + (i+1) + '】' + path.basename(f) + ' - 不存在\n');
    }
});
