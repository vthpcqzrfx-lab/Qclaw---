const fs = require('fs');
const path = require('path');

// 读取桌面顶层文件夹（排除我的公文包、11-12月素养信息流数据等已读文件夹）
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
const folders = fs.readdirSync(desktop).filter(f => {
    const fPath = path.join(desktop, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 桌面顶层文件夹 ===');
folders.forEach(f => {
    const fPath = path.join(desktop, f);
    const files = fs.readdirSync(fPath);
    console.log(f + '/ (' + files.length + '个文件)');
});
