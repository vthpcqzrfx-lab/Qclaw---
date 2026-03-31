const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取萌娃视频-小班
const mengwaBase = 'C:\\Users\\zhaoyang26\\Desktop\\萌娃视频-小班';
console.log('=== 萌娃视频-小班 ===');
if (fs.existsSync(mengwaBase)) {
    const mengwas = fs.readdirSync(mengwaBase);
    console.log('文件数：' + mengwas.length);
    mengwas.slice(0, 15).forEach(m => console.log('  ' + m));
    if (mengwas.length > 15) console.log('  ... 还有' + (mengwas.length - 15) + '个文件');
}

// 读取公司资质材料
const qualBase = 'C:\\Users\\zhaoyang26\\Desktop\\公司资质材料';
console.log('\n=== 公司资质材料 ===');
if (fs.existsSync(qualBase)) {
    const quals = fs.readdirSync(qualBase);
    quals.forEach(q => console.log('  ' + q));
}

// 读取不常用的软件
const softBase = 'C:\\Users\\zhaoyang26\\Desktop\\不常用的软件';
console.log('\n=== 不常用的软件 ===');
if (fs.existsSync(softBase)) {
    const softs = fs.readdirSync(softBase);
    softs.forEach(s => console.log('  ' + s));
}

// 读取更多我的公文包子文件夹
const moreFolders = ['高途-KA系列', '高途-KM青少学部', '高途-集团市场'];
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包';
moreFolders.forEach(folder => {
    const fPath = path.join(base, folder);
    if (fs.existsSync(fPath)) {
        console.log('\n=== ' + folder + ' ===');
        const files = fs.readdirSync(fPath);
        console.log('文件数：' + files.length);
        files.slice(0, 10).forEach(f => console.log('  ' + f));
    }
});
