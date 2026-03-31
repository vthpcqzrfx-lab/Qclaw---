const fs = require('fs');
const path = require('path');

// 读取桌面上的PPT文档列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

function findPPTs(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findPPTs(fPath, depth + 1);
        } else if (f.endsWith('.pptx') || f.endsWith('.ppt')) {
            console.log(fPath.replace('C:\\Users\\zhaoyang26\\Desktop\\', ''));
        }
    });
}

console.log('=== 桌面PPT文档列表 ===');
findPPTs(desktop);
