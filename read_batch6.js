const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取1月份素材
const materialBase = 'C:\\Users\\zhaoyang26\\Desktop\\1月份素材';
console.log('=== 1月份素材 ===');
const materials = fs.readdirSync(materialBase);
materials.forEach(m => {
    const mPath = path.join(materialBase, m);
    const stat = fs.statSync(mPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(mPath);
        console.log(m + '/ (' + sub.length + '个文件)');
    } else {
        console.log(m);
    }
});

// 读取关键词
const keywordBase = 'C:\\Users\\zhaoyang26\\Desktop\\关键词';
console.log('\n=== 关键词 ===');
const keywords = fs.readdirSync(keywordBase);
keywords.forEach(k => {
    const kPath = path.join(keywordBase, k);
    const stat = fs.statSync(kPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(kPath);
        console.log(k + '/ (' + sub.length + '个文件)');
    } else {
        console.log(k);
    }
});

// 读取EM业务线
const emBase = 'C:\\Users\\zhaoyang26\\Desktop\\EM业务线';
console.log('\n=== EM业务线 ===');
const ems = fs.readdirSync(emBase);
ems.forEach(e => {
    const ePath = path.join(emBase, e);
    const stat = fs.statSync(ePath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(ePath);
        console.log(e + '/ (' + sub.length + '个文件)');
        if (sub.length <= 5) {
            sub.forEach(s => console.log('  - ' + s));
        }
    } else {
        console.log(e);
    }
});

// 读取做一个B站号
const biliBase = 'C:\\Users\\zhaoyang26\\Desktop\\做一个B站号';
if (fs.existsSync(biliBase)) {
    console.log('\n=== 做一个B站号 ===');
    const bilis = fs.readdirSync(biliBase);
    bilis.forEach(b => {
        const bPath = path.join(biliBase, b);
        const stat = fs.statSync(bPath);
        if (stat.isDirectory()) {
            const sub = fs.readdirSync(bPath);
            console.log(b + '/ (' + sub.length + '个文件)');
        } else {
            console.log(b);
        }
    });
}

// 读取白马素材
const baima = 'C:\\Users\\zhaoyang26\\Desktop\\白马素材';
if (fs.existsSync(baima)) {
    console.log('\n=== 白马素材 ===');
    const baimaFiles = fs.readdirSync(baima);
    baimaFiles.slice(0, 10).forEach(b => console.log(b));
    if (baimaFiles.length > 10) console.log('... 还有' + (baimaFiles.length - 10) + '个文件');
}
