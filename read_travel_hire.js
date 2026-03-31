const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取出差文件夹的Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\出差';
const travelDirs = fs.readdirSync(base).filter(f => {
    const fPath = path.join(base, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 出差文件夹 ===');
travelDirs.forEach(d => {
    const dPath = path.join(base, d);
    const files = fs.readdirSync(dPath);
    const excels = files.filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    console.log(d + ': ' + files.length + '个文件，' + excels.length + '个Excel');
    if (excels.length > 0) {
        console.log('  Excel文件：' + excels.join(', '));
    }
});

// 读取招聘文件夹的Excel文件
const hireBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\招聘';
const hireFiles = fs.readdirSync(hireBase);
console.log('\n=== 招聘文件夹 ===');
hireFiles.forEach(f => {
    const fPath = path.join(hireBase, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
    } else {
        console.log(f);
    }
});
