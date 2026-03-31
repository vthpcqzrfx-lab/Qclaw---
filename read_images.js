const fs = require('fs');
const path = require('path');

// 读取桌面上的图片列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

function findImages(dir, depth = 0) {
    if (depth > 1) return; // 限制深度，只读顶层
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory() && depth === 0) {
            findImages(fPath, depth + 1);
        } else if (/\.(jpg|jpeg|png|gif|bmp|webp)$/i.test(f)) {
            console.log(fPath.replace('C:\\Users\\zhaoyang26\\Desktop\\', ''));
        }
    });
}

console.log('=== 桌面图片列表（顶层） ===');
findImages(desktop);
