const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取图书部数据汇总
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\图书部数据汇总';
const files = fs.readdirSync(base);
console.log('=== 图书部数据汇总 ===');
files.forEach(f => console.log(f));

// 读取星火每周数据总
const xhBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\星火每周数据总';
const xhFiles = fs.readdirSync(xhBase);
console.log('\n=== 星火每周数据总 ===');
xhFiles.slice(0, 10).forEach(f => console.log(f));
if (xhFiles.length > 10) console.log('... 还有' + (xhFiles.length - 10) + '个文件');

// 读取星火历史结算数据
const xhHistBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\20240910X业务线交接\\星火历史结算数据';
const xhHistFiles = fs.readdirSync(xhHistBase);
console.log('\n=== 星火历史结算数据 ===');
xhHistFiles.forEach(f => console.log(f));
