const fs = require('fs');
const path = require('path');

// 尝试读取PDF文件（使用简单的文本提取）
const pdfPath = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\APP运营数据.pdf';

try {
    const data = fs.readFileSync(pdfPath);
    // 尝试提取文本内容
    const text = data.toString('utf8', 0, Math.min(data.length, 10000));
    console.log('=== APP运营数据.pdf (前10000字节) ===');
    console.log(text.substring(0, 2000));
} catch (e) {
    console.log('Error:', e.message);
}

// 读取另一个PDF
const pdfPath2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\PPT\\市场部Ovmom-20211028.pdf';
try {
    const data2 = fs.readFileSync(pdfPath2);
    const text2 = data2.toString('utf8', 0, Math.min(data2.length, 10000));
    console.log('\n=== 市场部Ovmom-20211028.pdf (前10000字节) ===');
    console.log(text2.substring(0, 2000));
} catch (e) {
    console.log('Error:', e.message);
}
