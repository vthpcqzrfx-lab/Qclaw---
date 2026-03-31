const fs = require('fs');
const path = require('path');

// 读取不常用的软件
const base1 = 'C:\\Users\\zhaoyang26\\Desktop\\不常用的软件';
const files1 = fs.readdirSync(base1);
console.log('=== 不常用的软件 ===');
files1.forEach(f => console.log(f));

// 读取薪酬福利费报销
const base2 = 'C:\\Users\\zhaoyang26\\Desktop\\薪酬福利费报销';
const files2 = fs.readdirSync(base2);
console.log('\n=== 薪酬福利费报销 ===');
files2.forEach(f => console.log(f));

// 读取薪酬费用发票2026年
const base3 = 'C:\\Users\\zhaoyang26\\Desktop\\薪酬费用发票2026年';
const files3 = fs.readdirSync(base3);
console.log('\n=== 薪酬费用发票2026年 ===');
files3.forEach(f => console.log(f));

// 读取一些好看的简历
const base4 = 'C:\\Users\\zhaoyang26\\Desktop\\一些好看的简历';
const files4 = fs.readdirSync(base4);
console.log('\n=== 一些好看的简历 ===');
files4.forEach(f => console.log(f));

// 读取高途笔记截图
const base5 = 'C:\\Users\\zhaoyang26\\Desktop\\高途笔记截图';
const files5 = fs.readdirSync(base5);
console.log('\n=== 高途笔记截图 ===');
files5.slice(0, 20).forEach(f => console.log(f));
if (files5.length > 20) console.log('... 还有' + (files5.length - 20) + '个文件');
