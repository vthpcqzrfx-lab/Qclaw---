const XLSX = require('xlsx');

console.log('=== 读取更多关键数据 ===\n');

// 读取图书部数据汇总
const file1 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\图书部数据汇总\\图书1-3月收款收入-20240416.xlsx';
console.log('【文件13】图书1-3月收款收入.xlsx');
try {
    const wb1 = XLSX.readFile(file1);
    console.log('  - Sheets:', wb1.SheetNames);
    const sheet1 = wb1.Sheets['收款'];
    const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1, defval: ''});
    console.log('  - 数据行数：' + data1.length);
    console.log('  - 收款类型：课程、图书、其他、佣金');
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取业绩前20名单
const file2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\图书部数据汇总\\业绩前20名单.xlsx';
console.log('\n【文件14】业绩前20名单.xlsx');
try {
    const wb2 = XLSX.readFile(file2);
    const sheet2 = wb2.Sheets[wb2.SheetNames[0]];
    const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1, defval: ''});
    console.log('  - 分销员业绩Top5：');
    data2.slice(1, 6).forEach((row, i) => {
        console.log('    ' + (i+1) + '. ' + row[0] + ' - ' + row[1] + '元');
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取KA广州经历数据
const file3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\高途-KA广州经历\\5-11月广州线上例子.xlsx';
console.log('\n【文件15】5-11月广州线上例子.xlsx');
try {
    const wb3 = XLSX.readFile(file3);
    const sheet3 = wb3.Sheets[wb3.SheetNames[0]];
    const data3 = XLSX.utils.sheet_to_json(sheet3, {header: 1, defval: ''});
    console.log('  - 数据行数：' + data3.length);
    console.log('  - 期次示例：');
    data3.slice(1, 5).forEach((row, i) => {
        console.log('    ' + row[0] + ' - ' + row[4]);
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取私域运营资料
const file4 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\私域运营资料\\0615公司账号盘点.xlsx';
console.log('\n【文件16】0615公司账号盘点.xlsx');
try {
    const wb4 = XLSX.readFile(file4);
    const sheet4 = wb4.Sheets[wb4.SheetNames[0]];
    const data4 = XLSX.utils.sheet_to_json(sheet4, {header: 1, defval: ''});
    console.log('  - 账号盘点数量：' + (data4.length - 1));
} catch (e) {
    console.log('  Error:', e.message);
}

console.log('\n=== 读取完成 ===');
