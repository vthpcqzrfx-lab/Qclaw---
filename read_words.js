const fs = require('fs');
const path = require('path');

// 读取Word文档列表
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
let wordFiles = [];

function findWords(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findWords(fPath, depth + 1);
        } else if (f.endsWith('.docx') || f.endsWith('.doc')) {
            wordFiles.push({
                path: fPath,
                name: f,
                size: (stat.size / 1024).toFixed(2) // KB
            });
        }
    });
}

findWords(desktop);

console.log('=== Word文档统计 ===');
console.log('总计Word文档：' + wordFiles.length + '\n');

// 按大小排序
wordFiles.sort((a, b) => b.size - a.size);

console.log('最大的15个Word文档：');
wordFiles.slice(0, 15).forEach((f, i) => {
    console.log((i+1) + '. ' + f.name + ' (' + f.size + ' KB)');
});

console.log('\nWord文档分类：');
const categories = {
    'OVMOM目标': [],
    '工作文档': [],
    '合同协议': [],
    '培训文档': [],
    '个人文档': [],
    '其他': []
};

wordFiles.forEach(f => {
    const name = f.name.toLowerCase();
    if (name.includes('ovmom') || name.includes('目标')) categories['OVMOM目标'].push(f.name);
    else if (name.includes('合同') || name.includes('协议') || name.includes('框架')) categories['合同协议'].push(f.name);
    else if (name.includes('培训') || name.includes('讲稿') || name.includes('手册')) categories['培训文档'].push(f.name);
    else if (name.includes('周报') || name.includes('状态') || name.includes('复盘')) categories['工作文档'].push(f.name);
    else if (name.includes('简历') || name.includes('户型') || name.includes('入园')) categories['个人文档'].push(f.name);
    else categories['其他'].push(f.name);
});

Object.entries(categories).forEach(([cat, files]) => {
    console.log('  ' + cat + ': ' + files.length + '个');
    if (files.length > 0 && files.length <= 5) {
        files.forEach(f => console.log('    - ' + f));
    }
});
