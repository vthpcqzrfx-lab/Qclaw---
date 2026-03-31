const fs = require('fs');
const path = require('path');

// 读取我的公文包中剩余的子文件夹
const bases = [
    { name: '绩效/2022年绩效', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2022年绩效' },
    { name: '绩效/2023年绩效', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2023年绩效' },
    { name: '绩效/2024年绩效', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2024年绩效' },
    { name: '绩效/2025年绩效', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效\\2025年绩效' }
];

bases.forEach(b => {
    if (fs.existsSync(b.path)) {
        const files = fs.readdirSync(b.path);
        console.log('\n=== ' + b.name + ' ===');
        files.forEach(f => console.log(f));
    }
});
