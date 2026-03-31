const fs = require('fs');
const path = require('path');

// 读取桌面上的Word文档列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

function findDocs(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findDocs(fPath, depth + 1);
        } else if (f.endsWith('.docx') || f.endsWith('.doc')) {
            console.log(fPath.replace('C:\\Users\\zhaoyang26\\Desktop\\', ''));
        }
    });
}

console.log('=== 桌面Word文档列表 ===');
findDocs(desktop);
