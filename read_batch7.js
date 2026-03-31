const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取高途经历文件夹
const expBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包';
const expFolders = [
    '高途-KA中价课经历',
    '高途-KA广州经历', 
    '高途-KM青少学部经历',
    '高途-集团市场&X业务线经历',
    '高途体系培训'
];

expFolders.forEach(folder => {
    const fPath = path.join(expBase, folder);
    if (fs.existsSync(fPath)) {
        console.log('\n=== ' + folder + ' ===');
        const files = fs.readdirSync(fPath);
        console.log('文件数：' + files.length);
        files.slice(0, 15).forEach(f => console.log('  ' + f));
        if (files.length > 15) console.log('  ... 还有' + (files.length - 15) + '个文件');
    }
});

// 读取薪酬福利费报销
const salaryBase = path.join(expBase, '薪酬福利费报销');
if (fs.existsSync(salaryBase)) {
    console.log('\n=== 薪酬福利费报销 ===');
    const salaries = fs.readdirSync(salaryBase);
    salaries.forEach(s => {
        const sPath = path.join(salaryBase, s);
        const stat = fs.statSync(sPath);
        if (stat.isDirectory()) {
            const sub = fs.readdirSync(sPath);
            console.log(s + '/ (' + sub.length + '个文件)');
        } else {
            console.log(s);
        }
    });
}

// 读取团建
const teamBase = path.join(expBase, '团建');
if (fs.existsSync(teamBase)) {
    console.log('\n=== 团建 ===');
    const teams = fs.readdirSync(teamBase);
    console.log('文件数：' + teams.length);
    teams.slice(0, 10).forEach(t => console.log('  ' + t));
}
