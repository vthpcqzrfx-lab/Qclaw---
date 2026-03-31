const fs = require('fs');
const path = require('path');

// 统计所有文件类型
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

function countFiles(dir, ext = null) {
    let count = 0;
    try {
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            const fPath = path.join(dir, f);
            const stat = fs.statSync(fPath);
            if (stat.isDirectory()) {
                count += countFiles(fPath, ext);
            } else if (!ext || f.endsWith(ext)) {
                count++;
            }
        });
    } catch (e) {}
    return count;
}

console.log('=== 文件类型统计 ===');
console.log('Excel文件：' + countFiles(desktop, '.xlsx'));
console.log('Excel文件：' + countFiles(desktop, '.xls'));
console.log('Word文件：' + countFiles(desktop, '.docx'));
console.log('Word文件：' + countFiles(desktop, '.doc'));
console.log('PDF文件：' + countFiles(desktop, '.pdf'));
console.log('PPT文件：' + countFiles(desktop, '.pptx'));
console.log('PPT文件：' + countFiles(desktop, '.ppt'));
console.log('视频文件：' + countFiles(desktop, '.mp4'));
console.log('图片文件：' + countFiles(desktop, '.jpg'));
console.log('图片文件：' + countFiles(desktop, '.png'));
console.log('图片文件：' + countFiles(desktop, '.jpeg'));

// 列出最大的文件夹
console.log('\n=== 最大的文件夹 ===');
const folders = fs.readdirSync(desktop).filter(f => {
    const fPath = path.join(desktop, f);
    return fs.statSync(fPath).isDirectory();
});

const folderSizes = folders.map(f => {
    const fPath = path.join(desktop, f);
    const count = countFiles(fPath);
    return { name: f, count };
}).sort((a, b) => b.count - a.count);

folderSizes.slice(0, 15).forEach(f => {
    console.log(f.name + ': ' + f.count + '个文件');
});
