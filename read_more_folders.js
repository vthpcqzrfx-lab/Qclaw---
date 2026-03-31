const fs = require('fs');
const path = require('path');

// 读取私域运营资料文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\私域运营资料';
const files = fs.readdirSync(base);

console.log('=== 私域运营资料 ===');
files.forEach(f => {
    const fPath = path.join(base, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.slice(0, 3).forEach(s => console.log('  - ' + s));
    } else {
        console.log(f);
    }
});

// 读取授予协议文件夹
const grantBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\授予协议';
const grants = fs.readdirSync(grantBase);
console.log('\n=== 授予协议 ===');
grants.forEach(g => console.log(g));

// 读取合作方文件夹
const partnerBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\合作方';
const partners = fs.readdirSync(partnerBase);
console.log('\n=== 合作方 ===');
partners.forEach(p => console.log(p));

// 读取信息流数据分享
const infoShareBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\信息流数据分享';
const infoShares = fs.readdirSync(infoShareBase);
console.log('\n=== 信息流数据分享 ===');
infoShares.forEach(i => console.log(i));
