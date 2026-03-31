const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取薪酬福利费报销的Excel文件
const salaryBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\薪酬福利费报销';
const salaryDirs = fs.readdirSync(salaryBase).filter(f => {
    const fPath = path.join(salaryBase, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 薪酬福利费报销月份 ===');
salaryDirs.forEach(d => {
    const dPath = path.join(salaryBase, d);
    const files = fs.readdirSync(dPath);
    console.log(d + ': ' + files.length + '个文件');
});

// 读取薪酬福利费报销的Excel文件
const salaryExcels = [];
salaryDirs.forEach(d => {
    const dPath = path.join(salaryBase, d);
    const files = fs.readdirSync(dPath).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    files.forEach(f => salaryExcels.push(path.join(dPath, f)));
});

console.log('\n=== 薪酬福利费报销Excel文件数：' + salaryExcels.length + ' ===');

// 读取前3个Excel文件
salaryExcels.slice(0, 3).forEach(f => {
    try {
        const workbook = XLSX.readFile(f);
        console.log('\n--- ' + f.split('\\').pop() + ' ---');
        console.log('Sheets:', workbook.SheetNames);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        console.log('Total rows:', data.length);
        data.slice(0, 5).forEach((row, i) => {
            if (row.some(c => c !== '')) {
                console.log('R' + i + ':', row.slice(0, 5).map(c => String(c).substring(0, 25)));
            }
        });
    } catch (e) {
        console.log('Error:', e.message);
    }
});
