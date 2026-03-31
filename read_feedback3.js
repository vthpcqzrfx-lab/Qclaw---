const XLSX = require('xlsx');
const fs = require('fs');

console.log('=== 继续深度读取 ===\n');

// 读取初中信息流数据
const file1 = 'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\初中信息流过程结果数据.xlsx';
console.log('【文件7】初中信息流过程结果数据.xlsx');
try {
    const wb1 = XLSX.readFile(file1);
    console.log('  - Sheets:', wb1.SheetNames);
    const sheet1 = wb1.Sheets['分渠道转化'];
    const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1, defval: ''});
    console.log('  - 初中信息流核心数据：');
    console.log('    总流量：' + data1[4][1]);
    console.log('    P8及以上占比：' + data1[4][2]);
    console.log('    加好友数：' + data1[4][3]);
    console.log('    好友率：' + data1[4][4]);
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取星火每周数据
const file2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\星火每周数据总\\10月.xlsx';
console.log('\n【文件8】星火每周数据总/10月.xlsx');
try {
    const wb2 = XLSX.readFile(file2);
    const sheet2 = wb2.Sheets['每日数据'];
    const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1, defval: ''});
    console.log('  - 10月好友数据（前5天）：');
    data2.slice(1, 6).forEach((row, i) => {
        console.log('    ' + row[0] + ': 新增' + row[1] + ', 删除' + row[2] + ', 总数' + row[3]);
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取渠道GMV规划
const file3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\给销售伙伴.xlsx';
console.log('\n【文件9】给销售伙伴.xlsx（渠道GMV规划）');
try {
    const wb3 = XLSX.readFile(file3);
    const sheet3 = wb3.Sheets[wb3.SheetNames[0]];
    const data3 = XLSX.utils.sheet_to_json(sheet3, {header: 1, defval: ''});
    console.log('  - 渠道GMV规划：');
    data3.slice(1, 6).forEach((row, i) => {
        if (row[0]) {
            console.log('    ' + row[0] + ': 2023年' + row[1] + '/天, 2024年' + row[2] + '/天, 增长' + row[4]);
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取高途代理
const file4 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\高途代理.xls';
console.log('\n【文件10】高途代理.xls');
try {
    const wb4 = XLSX.readFile(file4);
    const sheet4 = wb4.Sheets[wb4.SheetNames[0]];
    const data4 = XLSX.utils.sheet_to_json(sheet4, {header: 1, defval: ''});
    console.log('  - 代理商数据（前5家）：');
    data4.slice(1, 6).forEach((row, i) => {
        console.log('    ' + row[0] + ' - 对接人:' + row[1]);
    });
} catch (e) {
    console.log('  Error:', e.message);
}

console.log('\n=== 读取完成 ===');
