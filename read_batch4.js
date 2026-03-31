const XLSX = require('xlsx');
const fs = require('fs');

// 读取关键绩效Excel
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效\\2022年年度绩效评定.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2025年绩效\\创新增长组2025年Q1绩效打分.xls',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\会议\\20221010Q3晋升会\\附件2-晋升答辩汇总名单-【私域运营】.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\会议\\20230905三六零人员盘点\\3【战略规划&组织架构调整计划】.xlsx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\会议\\书链合作\\天猫日常销售情况.xlsx'
];

files.forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n=== ' + f.split('\\').pop() + ' ===');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        data.slice(0, 12).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                console.log('R' + i + ':', row.slice(0, 8).map(c => String(c).substring(0, 30)));
            }
        });
    } catch (e) {
        console.log('Error reading ' + f.split('\\').pop() + ':', e.message);
    }
});
