const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取PPT文件列表
const pptBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT';
const ppts = fs.readdirSync(pptBase);
console.log('=== PPT文件列表 ===');
ppts.forEach(p => console.log(p));

// 读取数据文件夹
const dataBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据';
if (fs.existsSync(dataBase)) {
    console.log('\n=== 数据文件夹 ===');
    const dataFiles = fs.readdirSync(dataBase);
    dataFiles.forEach(d => console.log(d));
}

// 读取个人文件夹
const personalBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\个人';
if (fs.existsSync(personalBase)) {
    console.log('\n=== 个人文件夹 ===');
    const personalFiles = fs.readdirSync(personalBase);
    personalFiles.forEach(p => console.log(p));
}

// 读取出差文件夹详情
const travelBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\出差';
console.log('\n=== 出差报销详情 ===');
const travels = fs.readdirSync(travelBase);
travels.forEach(t => {
    const tPath = path.join(travelBase, t);
    const stat = fs.statSync(tPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(tPath);
        console.log(t + '/ (' + sub.length + '个文件)');
    } else {
        console.log(t);
    }
});

// 读取招聘文件夹
const hireBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\招聘';
console.log('\n=== 招聘文件 ===');
const hires = fs.readdirSync(hireBase);
hires.forEach(h => {
    const hPath = path.join(hireBase, h);
    const stat = fs.statSync(hPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(hPath);
        console.log(h + '/ (' + sub.length + '个文件)');
        sub.slice(0, 2).forEach(s => console.log('  - ' + s));
    } else {
        console.log(h);
    }
});
