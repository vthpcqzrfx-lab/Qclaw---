const fs = require('fs');
const path = require('path');

// 读取PDF文件列表并统计
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
let pdfFiles = [];

function findPDFs(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            findPDFs(fPath, depth + 1);
        } else if (f.endsWith('.pdf')) {
            pdfFiles.push({
                path: fPath,
                name: f,
                size: (stat.size / 1024 / 1024).toFixed(2) // MB
            });
        }
    });
}

findPDFs(desktop);

console.log('=== PDF文件统计 ===');
console.log('总计PDF文件：' + pdfFiles.length + '\n');

// 按大小排序
pdfFiles.sort((a, b) => b.size - a.size);

console.log('最大的10个PDF文件：');
pdfFiles.slice(0, 10).forEach((f, i) => {
    console.log((i+1) + '. ' + f.name + ' (' + f.size + ' MB)');
});

console.log('\nPDF文件分类：');
const categories = {
    '简历': 0,
    'OVMOM': 0,
    '培训': 0,
    '工作文档': 0,
    '个人证件': 0,
    '发票报销': 0,
    '其他': 0
};

pdfFiles.forEach(f => {
    const name = f.name.toLowerCase();
    if (name.includes('简历') || name.includes('cv')) categories['简历']++;
    else if (name.includes('ovmom') || name.includes('目标')) categories['OVMOM']++;
    else if (name.includes('培训') || name.includes('ppt') || name.includes('手册')) categories['培训']++;
    else if (name.includes('合同') || name.includes('协议') || name.includes('文档')) categories['工作文档']++;
    else if (name.includes('合同') || name.includes('证明') || name.includes('体检')) categories['个人证件']++;
    else if (name.includes('发票') || name.includes('报销') || name.includes('元')) categories['发票报销']++;
    else categories['其他']++;
});

Object.entries(categories).forEach(([cat, count]) => {
    console.log('  ' + cat + ': ' + count + '个');
});
