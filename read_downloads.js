const fs = require('fs');
const path = require('path');

// 读取下载文件夹
const downloadPath = 'C:\\Users\\zhaoyang26\\Downloads';

if (!fs.existsSync(downloadPath)) {
    console.log('下载文件夹不存在');
    process.exit(1);
}

const files = fs.readdirSync(downloadPath);

console.log('=== 下载文件夹内容 ===');
console.log('总文件数：' + files.length + '\n');

// 按类型分类
const types = {
    excel: [],
    pdf: [],
    word: [],
    ppt: [],
    image: [],
    video: [],
    zip: [],
    other: []
};

files.forEach(f => {
    const ext = path.extname(f).toLowerCase();
    const stat = fs.statSync(path.join(downloadPath, f));
    const size = (stat.size / 1024 / 1024).toFixed(2); // MB
    
    if (ext === '.xlsx' || ext === '.xls') types.excel.push({name: f, size});
    else if (ext === '.pdf') types.pdf.push({name: f, size});
    else if (ext === '.docx' || ext === '.doc') types.word.push({name: f, size});
    else if (ext === '.pptx' || ext === '.ppt') types.ppt.push({name: f, size});
    else if (['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'].includes(ext)) types.image.push({name: f, size});
    else if (['.mp4', '.avi', '.mov', '.mkv', '.wmv'].includes(ext)) types.video.push({name: f, size});
    else if (['.zip', '.rar', '.7z'].includes(ext)) types.zip.push({name: f, size});
    else types.other.push({name: f, size});
});

// 输出统计
console.log('Excel文件：' + types.excel.length);
types.excel.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\nPDF文件：' + types.pdf.length);
types.pdf.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\nWord文件：' + types.word.length);
types.word.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\nPPT文件：' + types.ppt.length);
types.ppt.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\n图片文件：' + types.image.length);
types.image.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\n视频文件：' + types.video.length);
types.video.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\n压缩文件：' + types.zip.length);
types.zip.slice(0, 10).forEach(f => console.log('  - ' + f.name + ' (' + f.size + ' MB)'));

console.log('\n其他文件：' + types.other.length);
