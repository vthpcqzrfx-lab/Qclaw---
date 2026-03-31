const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取更多Excel - 批次24（X业务线交接-图书部数据汇总）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\图书部数据汇总';
const files = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls')).slice(0, 8);

console.log('=== 批次24 - 图书部数据汇总 ===\n');

files.forEach((f, i) => {
    const fPath = path.join(base, f);
    try {
        const wb = XLSX.readFile(fPath);
        console.log('【' + (i+1) + '】' + f);
        console.log('  Sheets: ' + wb.SheetNames.slice(0, 4).join(', '));
        if (wb.SheetNames.length > 4) {
            console.log('  ... 还有' + (wb.SheetNames.length - 4) + '个Sheet');
        }
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('  行数: ' + data.length);
        console.log('');
    } catch (e) {
        console.log('  Error: ' + e.message + '\n');
    }
});
