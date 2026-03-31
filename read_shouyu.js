const fs = require('fs');
const path = require('path');

// 读取授予协议文件夹的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\授予协议';
const files = fs.readdirSync(base);

console.log('=== 授予协议文件夹详细 ===');
files.forEach(f => console.log(f));

// 读取合作方文件夹的详细内容
const partnerBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\合作方';
const partnerDirs = fs.readdirSync(partnerBase);
console.log('\n=== 合作方文件夹详细 ===');
partnerDirs.forEach(d => {
    const dPath = path.join(partnerBase, d);
    const files = fs.readdirSync(dPath);
    console.log(d + '/ (' + files.length + '个文件)');
    files.slice(0, 3).forEach(f => console.log('  - ' + f));
});

// 读取头像图片文件夹的详细内容
const avatarBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\头像图片';
const avatarFiles = fs.readdirSync(avatarBase);
console.log('\n=== 头像图片文件夹详细 ===');
console.log('文件数：' + avatarFiles.length);
avatarFiles.forEach(f => console.log(f));
