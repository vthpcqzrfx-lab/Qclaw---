const fs = require('fs');
const path = require('path');

// 读取桌面上的PDF文件列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

function findPDFs(dir, depth = 0) {
    if (depth > 2) return; // 限制深度
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findPDFs(fPath, depth + 1);
        } else if (f.endsWith('.pdf')) {
            console.log(fPath.replace('C:\\Users\\zhaoyang26\\Desktop\\', ''));
        }
    });
}

console.log('=== 桌面PDF文件列表 ===');
findPDFs(desktop);
