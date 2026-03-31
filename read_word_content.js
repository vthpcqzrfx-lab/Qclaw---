const fs = require('fs');
const path = require('path');

// 读取Word文档内容（尝试提取文本）
const files = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\市场部-私域运营部22年OVMOM.doc',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\市场部周报模版.docx',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\市场部私域23年ovmom.docx'
];

console.log('=== 读取Word文档内容 ===\n');

files.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const data = fs.readFileSync(f);
            // 尝试提取文本
            const text = data.toString('utf8', 0, Math.min(data.length, 5000));
            console.log('【' + (i+1) + '】' + path.basename(f));
            console.log('  文件大小：' + (data.length / 1024).toFixed(2) + ' KB');
            // 提取可读的文本片段
            const readable = text.replace(/[^\x20-\x7E\u4e00-\u9fa5]/g, ' ').replace(/\s+/g, ' ').trim();
            console.log('  内容预览：' + readable.substring(0, 200));
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    }
});
