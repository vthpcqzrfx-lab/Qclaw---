const XLSX = require('xlsx');
const fs = require('fs');

// 读取绩效文件
const perfFiles = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\【主管级（含储备）管理能力打分表】-to赵杨-完成.xlsx'
];

// 读取X业务线文档
const xFiles = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\7日分销员特训营.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\给销售伙伴.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\星火-素养 历史数据.xlsx'
];

// 读取0323信息流分析
const infoFiles = [
    'C:\\Users\\zhaoyang26\\Desktop\\0323信息流分析\\线索明细数据 (14).xlsx'
];

const allFiles = [...perfFiles, ...xFiles, ...infoFiles];

allFiles.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        
        workbook.SheetNames.slice(0, 2).forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('\n--- Sheet: ' + sheetName + ' (' + data.length + ' rows) ---');
            data.slice(0, 10).forEach((row, i) => {
                if (row.some(c => c !== '')) {
                    console.log('R' + i + ':', row.slice(0, 8).map(c => String(c).substring(0, 30)));
                }
            });
        });
    } catch (e) {
        console.log('Error reading ' + f.split('\\').pop() + ':', e.message);
    }
});
