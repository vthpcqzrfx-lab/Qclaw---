const fs = require('fs');
const path = require('path');

// 连续读取 - 批次21（统计所有文件类型）
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

const stats = {
    excel: 0,
    pdf: 0,
    word: 0,
    ppt: 0,
    image: 0,
    video: 0,
    other: 0
};

function scanFiles(dir, depth = 0) {
    if (depth > 2) return;
    const files = fs.readdirSync(dir);
    files.forEach(f => {
        const fPath = path.join(dir, f);
        const stat = fs.statSync(fPath);
        if (stat.isDirectory()) {
            scanFiles(fPath, depth + 1);
        } else {
            const ext = path.extname(f).toLowerCase();
            if (ext === '.xlsx' || ext === '.xls') stats.excel++;
            else if (ext === '.pdf') stats.pdf++;
            else if (ext === '.docx' || ext === '.doc') stats.word++;
            else if (ext === '.pptx' || ext === '.ppt') stats.ppt++;
            else if (['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'].includes(ext)) stats.image++;
            else if (['.mp4', '.avi', '.mov', '.mkv', '.wmv'].includes(ext)) stats.video++;
            else stats.other++;
        }
    });
}

console.log('=== 统计桌面所有文件类型 ===\n');
scanFiles(desktop);

console.log('Excel文件：' + stats.excel);
console.log('PDF文件：' + stats.pdf);
console.log('Word文件：' + stats.word);
console.log('PPT文件：' + stats.ppt);
console.log('图片文件：' + stats.image);
console.log('视频文件：' + stats.video);
console.log('其他文件：' + stats.other);
console.log('');
console.log('总计：' + (stats.excel + stats.pdf + stats.word + stats.ppt + stats.image + stats.video + stats.other));
