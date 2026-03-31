const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 连续读取 - 批次18（KA广州经历）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KA广州经历';
const files = fs.readdirSync(base).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls') || f.endsWith('.et'));

console.log('=== 连续读取 - 批次18（KA广州经历）===\n');
console.log('Excel文件数：' + files.length + '\n');

files.slice(0, 8).forEach((f, i) => {
    const fPath = path.join(base, f);
    try {
        const wb = XLSX.readFile(fPath);
        console.log('【' + (i+1) + '】' + f);
        console.log('  Sheets: ' + wb.SheetNames.slice(0, 5).join(', '));
        if (wb.SheetNames.length > 5) {
            console.log('  ... 还有' + (wb.SheetNames.length - 5) + '个Sheet');
        }
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('  行数: ' + data.length);
        console.log('');
    } catch (e) {
        console.log('  Error: ' + e.message + '\n');
    }
});

if (files.length > 8) {
    console.log('... 还有' + (files.length - 8) + '个Excel文件');
}
