const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

console.log('=== 开始读取重要Excel文件 ===\n');

// 读取增长数据总览
const file1 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\增长数据总览.xlsx';
console.log('【文件1】增长数据总览.xlsx');
try {
    const wb1 = XLSX.readFile(file1);
    const sheet1 = wb1.Sheets[wb1.SheetNames[0]];
    const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1, defval: ''});
    console.log('  - 数据行数：' + data1.length);
    console.log('  - 关键数据：');
    data1.slice(1, 10).forEach((row, i) => {
        if (row[0] && row[3]) {
            console.log('    ' + row[0] + ' - ' + row[3] + ': ' + row[4]);
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取私域运营组架构
const file2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\私域运营组架构.xlsx';
console.log('\n【文件2】私域运营组架构.xlsx');
try {
    const wb2 = XLSX.readFile(file2);
    console.log('  - Sheets:', wb2.SheetNames);
    const sheet2 = wb2.Sheets[wb2.SheetNames[0]];
    const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1, defval: ''});
    console.log('  - 组织架构：');
    data2.slice(1, 8).forEach((row, i) => {
        if (row[0] && row[1]) {
            console.log('    ' + row[0] + ' - ' + row[1] + ' (' + row[2] + ')');
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取创新增长组绩效
const file3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\创新增长组绩效-2023Q3.xlsx';
console.log('\n【文件3】创新增长组绩效-2023Q3.xlsx');
try {
    const wb3 = XLSX.readFile(file3);
    console.log('  - Sheets:', wb3.SheetNames);
    const sheet3 = wb3.Sheets['绩效方案'];
    const data3 = XLSX.utils.sheet_to_json(sheet3, {header: 1, defval: ''});
    console.log('  - 绩效指标：');
    data3.slice(3, 8).forEach((row, i) => {
        if (row[2]) {
            console.log('    ' + row[2] + ' - 权重' + row[4] + ' - ' + row[5]);
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

console.log('\n=== 读取完成 ===');
