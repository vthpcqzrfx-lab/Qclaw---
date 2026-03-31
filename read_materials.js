const fs = require('fs');
const path = require('path');

// 读取1月份素材的详细内容
const base = 'C:\\Users\\zhaoyang26\\Desktop\\1月份素材';
const folders = fs.readdirSync(base).filter(f => {
    const fPath = path.join(base, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 1月份素材详细 ===');
folders.forEach(f => {
    const fPath = path.join(base, f);
    const files = fs.readdirSync(fPath);
    console.log('\n' + f + '/ (' + files.length + '个文件)');
    files.slice(0, 5).forEach(file => console.log('  - ' + file));
    if (files.length > 5) console.log('  ... 还有' + (files.length - 5) + '个文件');
});

// 读取EM业务线的详细内容
const emBase = 'C:\\Users\\zhaoyang26\\Desktop\\EM业务线';
const emFiles = fs.readdirSync(emBase);
console.log('\n=== EM业务线详细 ===');
emFiles.forEach(f => {
    const fPath = path.join(emBase, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
        sub.slice(0, 5).forEach(s => console.log('  - ' + s));
    } else {
        console.log(f);
    }
});
