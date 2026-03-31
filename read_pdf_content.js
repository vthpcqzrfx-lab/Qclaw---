const fs = require('fs');
const path = require('path');

// 尝试读取PDF文本内容
const pdfFiles = [
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\APP运营数据.pdf',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\市场部Ovmom-20211028.pdf',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\私域全局培训-徐瑞楠.pdf',
    'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\高途传统裂变-赵奕尧.pdf'
];

console.log('=== 尝试读取PDF内容 ===\n');

pdfFiles.forEach((f, i) => {
    if (fs.existsSync(f)) {
        try {
            const data = fs.readFileSync(f);
            const size = (data.length / 1024).toFixed(2);
            console.log('【' + (i+1) + '】' + path.basename(f));
            console.log('  大小: ' + size + ' KB');
            
            // 尝试提取文本（PDF文本通常在特定位置）
            const text = data.toString('utf8', 0, Math.min(data.length, 50000));
            // 查找可读文本
            const readableParts = text.match(/[\x20-\x7E\u4e00-\u9fa5]{10,}/g);
            if (readableParts && readableParts.length > 0) {
                console.log('  提取文本片段:');
                readableParts.slice(0, 5).forEach((part, idx) => {
                    console.log('    ' + (idx+1) + '. ' + part.substring(0, 100));
                });
            } else {
                console.log('  未提取到可读文本（可能是扫描版PDF）');
            }
            console.log('');
        } catch (e) {
            console.log('  Error: ' + e.message + '\n');
        }
    } else {
        console.log('【' + (i+1) + '】' + path.basename(f) + ' - 不存在\n');
    }
});
