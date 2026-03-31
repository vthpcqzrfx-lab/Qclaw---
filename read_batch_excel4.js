const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次4（关键词库）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\关键词';
const files = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

console.log('=== 批量读取Excel - 批次4（关键词库）===\n');
console.log('关键词文件总数：' + files.length + '\n');

files.slice(0, 10).forEach((f, i) => {
    const fPath = path.join(base, f);
    try {
        const wb = XLSX.readFile(fPath);
        console.log('【' + (i+1) + '】' + f);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('  行数: ' + data.length);
    } catch (e) {
        console.log('  Error: ' + e.message);
    }
});

if (files.length > 10) {
    console.log('\n... 还有' + (files.length - 10) + '个关键词文件');
}
