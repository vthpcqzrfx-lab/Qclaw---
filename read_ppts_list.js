const fs = require('fs');
const path = require('path');

// 读取PPT文档列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
let pptFiles = [];

function findPPTs(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findPPTs(fPath, depth + 1);
        } else if (f.endsWith('.pptx') || f.endsWith('.ppt')) {
            pptFiles.push({
                path: fPath,
                name: f,
                size: (stat.size / 1024).toFixed(2) // KB
            });
        }
    });
}

findPPTs(desktop);

console.log('=== PPT文档统计 ===');
console.log('总计PPT文档：' + pptFiles.length + '\n');

// 按大小排序
pptFiles.sort((a, b) => b.size - a.size);

console.log('所有PPT文档（按大小排序）：');
pptFiles.forEach((f, i) => {
    console.log((i+1) + '. ' + f.name + ' (' + f.size + ' KB)');
});

console.log('\nPPT文档分类：');
const categories = {
    '述职汇报': [],
    '培训方法论': [],
    '会议沟通': [],
    '其他': []
};

pptFiles.forEach(f => {
    const name = f.name.toLowerCase();
    if (name.includes('述职') || name.includes('汇报') || name.includes('总结') || name.includes('展望')) categories['述职汇报'].push(f.name);
    else if (name.includes('培训') || name.includes('方法论') || name.includes('私域') || name.includes('裂变')) categories['培训方法论'].push(f.name);
    else if (name.includes('会议') || name.includes('沟通') || name.includes('季度')) categories['会议沟通'].push(f.name);
    else categories['其他'].push(f.name);
});

Object.entries(categories).forEach(([cat, files]) => {
    console.log('  ' + cat + ': ' + files.length + '个');
    if (files.length > 0) {
        files.slice(0, 3).forEach(f => console.log('    - ' + f));
        if (files.length > 3) console.log('    ... 还有' + (files.length - 3) + '个');
    }
});
