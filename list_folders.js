const fs = require('fs');
const path = require('path');

// 列出我的公文包的所有子文件夹
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包';
const items = fs.readdirSync(base);

console.log('=== 我的公文包子文件夹 ===');
items.forEach(item => {
    const itemPath = path.join(base, item);
    const stat = fs.statSync(itemPath);
    if (stat.isDirectory()) {
        const files = fs.readdirSync(itemPath);
        console.log(item + '/ (' + files.length + '个文件)');
    } else {
        console.log(item);
    }
});
