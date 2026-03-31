const XLSX = require('xlsx');
const fs = require('fs');

console.log('=== 继续读取关键Excel文件 ===\n');

// 读取图书产品部目标
const file1 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\2024年ovmom\\图书产品部-24年目标拆解-私域V3版.xlsx';
console.log('【文件4】图书产品部-24年目标拆解-私域V3版.xlsx');
try {
    const wb1 = XLSX.readFile(file1);
    console.log('  - Sheets:', wb1.SheetNames.slice(0, 8));
    const sheet1 = wb1.Sheets['汇总及渠道拆解'];
    const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1, defval: ''});
    console.log('  - 2024年目标：');
    data1.slice(0, 10).forEach((row, i) => {
        if (row[0] && row[6]) {
            console.log('    ' + row[0] + ': ' + row[6]);
        }
    });
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取信息流数据
const file2 = 'C:\\Users\\zhaoyang26\\Desktop\\11-12月素养信息流数据\\2024年信息流-思维杨易.xlsx';
console.log('\n【文件5】2024年信息流-思维杨易.xlsx');
try {
    const wb2 = XLSX.readFile(file2);
    console.log('  - Sheets:', wb2.SheetNames);
    const sheet2 = wb2.Sheets['渠道分期'];
    const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1, defval: ''});
    console.log('  - 腾讯广点通-思维全年数据：');
    console.log('    总流量：' + data2[1][2]);
    console.log('    总收款：' + data2[1][3]);
    console.log('    平均单效：' + data2[1][4]);
    console.log('    转化率：' + data2[1][5]);
    console.log('    渠道ROI：' + data2[1][7]);
} catch (e) {
    console.log('  Error:', e.message);
}

// 读取分销员数据
const file3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档\\7日分销员特训营.xlsx';
console.log('\n【文件6】7日分销员特训营.xlsx');
try {
    const wb3 = XLSX.readFile(file3);
    const sheet3 = wb3.Sheets['Sheet1'];
    const data3 = XLSX.utils.sheet_to_json(sheet3, {header: 1, defval: ''});
    console.log('  - 分销员数据：');
    console.log('    参与人数：' + data3[1][1]);
    console.log('    好友总数：' + data3[1][3]);
    console.log('    日期分布：1122-1128');
} catch (e) {
    console.log('  Error:', e.message);
}

console.log('\n=== 读取完成 ===');
