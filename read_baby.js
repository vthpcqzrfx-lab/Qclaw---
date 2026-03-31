const fs = require('fs');
const path = require('path');

// 读取赵星晨 798 一周岁文件夹（孩子照片）
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\赵星晨 798 一周岁';
const files = fs.readdirSync(base);

console.log('=== 赵星晨 798 一周岁 ===');
console.log('文件数：' + files.length);
files.slice(0, 20).forEach(f => console.log(f));
if (files.length > 20) console.log('... 还有' + (files.length - 20) + '个文件');

// 统计文件类型
const types = {};
files.forEach(f => {
    const ext = path.extname(f).toLowerCase();
    types[ext] = (types[ext] || 0) + 1;
});

console.log('\n文件类型分布：');
Object.entries(types).forEach(([ext, count]) => {
    console.log(ext + ': ' + count);
});
