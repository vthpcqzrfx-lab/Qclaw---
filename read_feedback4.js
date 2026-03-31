const XLSX = require('xlsx');

console.log('=== 继续读取关键数据 ===\n');

// 读取渠道GMV规划
const file1 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\给销售伙伴.xlsx';
console.log('【文件9】给销售伙伴.xlsx（渠道GMV规划）');
try {
    const wb1 = XLSX.readFile(file1);
    const sheet1 = wb1.Sheets[wb1.SheetNames[0]];
    const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1, defval: ''});
    console.log('  - 渠道GMV规划：');
    data1.slice(1, 6).forEach((row, i) => {
        if (row[0]) {
            console.log('    ' + row[0] + ': 2023年' + row[1] + '/天, 2024年' + row[2] + '/天');
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取高途代理
const file2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\高途代理.xls';
console.log('\n【文件10】高途代理.xls');
try {
    const wb2 = XLSX.readFile(file2);
    const sheet2 = wb2.Sheets[wb2.SheetNames[0]];
    const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1, defval: ''});
    console.log('  - 代理商数量：' + (data2.length - 1));
    console.log('  - 前5家代理商：');
    data2.slice(1, 6).forEach((row, i) => {
        console.log('    ' + (i+1) + '. ' + row[0]);
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取工业化种草拆解
const file3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\工业化种草拆解V1.xlsx';
console.log('\n【文件11】工业化种草拆解V1.xlsx');
try {
    const wb3 = XLSX.readFile(file3);
    const sheet3 = wb3.Sheets[wb3.SheetNames[0]];
    const data3 = XLSX.utils.sheet_to_json(sheet3, {header: 1, defval: ''});
    console.log('  - 总GMV：' + data3[1][1]);
    console.log('  - Q1：' + data3[1][2]);
    console.log('  - Q2：' + data3[1][3]);
    console.log('  - Q3：' + data3[1][4]);
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取星火历史结算
const file4 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\星火历史结算数据\\3-5月结算数据表（0628-1）.xlsx';
console.log('\n【文件12】3-5月结算数据表.xlsx');
try {
    const wb4 = XLSX.readFile(file4);
    console.log('  - Sheets:', wb4.SheetNames);
    const sheet4 = wb4.Sheets[wb4.SheetNames[0]];
    const data4 = XLSX.utils.sheet_to_json(sheet4, {header: 1, defval: ''});
    console.log('  - 数据行数：' + data4.length);
} catch (e) {
    console.log('  Error:', e.message);
}

console.log('\n=== 读取完成 ===');
