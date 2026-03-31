const fs = require('fs');
const path = require('path');

// 读取我的公文包中未读的文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包';
const folders = fs.readdirSync(base);

console.log('=== 我的公文包文件夹列表 ===');
folders.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.statSync(fPath).isDirectory()) {
        const files = fs.readdirSync(fPath);
        console.log(f + '/ (' + files.length + '个文件)');
    } else {
        console.log(f);
    }
});
