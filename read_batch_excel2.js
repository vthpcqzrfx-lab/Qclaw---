const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 批量读取Excel文件 - 批次2
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KM青少学部经历\\爆款文案素材库.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-集团市场&X业务线经历\\【k12】爆款文案素材库.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-集团市场&X业务线经历\\朋友圈节奏运营.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-集团市场&X业务线经历\\星火-素养.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KA广州经历\\心田花开官方号_直播记录_20250101-20251110.et'
];

console.log('=== 批量读取Excel - 批次2 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const wb = XLSX.readFile(f);
            const fileName = path.basename(f);
            console.log('【' + (i+1) + '】' + fileName);
            console.log('  Sheets: ' + wb.SheetNames.slice(0, 10).join(', '));
            if (wb.SheetNames.length > 10) {
                console.log('  ... 还有' + (wb.SheetNames.length - 10) + '个Sheet');
            }
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('  行数: ' + data.length);
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    } else {
        console.log('【' + (i+1) + '】' + path.basename(f) + ' - 文件不存在\n');
    }
});
