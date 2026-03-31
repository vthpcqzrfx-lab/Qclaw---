const fs = require('fs');
const path = require('path');

// 读取做一个B站号的详细内容
const biliBase = 'C:\\Users\\zhaoyang26\\Desktop\\做一个B站号';
const biliFiles = fs.readdirSync(biliBase);

console.log('=== 做一个B站号详细 ===');
biliFiles.forEach(f => {
    const fPath = path.join(biliBase, f);
    const stat = fs.statSync(fPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(fPath);
        console.log(f + '/ (' + sub.length + '个文件)');
    } else {
        const size = (stat.size / 1024 / 1024).toFixed(2);
        console.log(f + ' (' + size + ' MB)');
    }
});

// 读取白马素材
const baimaBase = 'C:\\Users\\zhaoyang26\\Desktop\\白马素材';
const baimaFiles = fs.readdirSync(baimaBase);
console.log('\n=== 白马素材 ===');
console.log('文件数：' + baimaFiles.length);
baimaFiles.forEach(f => console.log(f));

// 读取萌娃视频-小班
const mengwaBase = 'C:\\Users\\zhaoyang26\\Desktop\\萌娃视频-小班';
const mengwaFiles = fs.readdirSync(mengwaBase);
console.log('\n=== 萌娃视频-小班 ===');
console.log('文件数：' + mengwaFiles.length);
mengwaFiles.slice(0, 10).forEach(f => console.log(f));
