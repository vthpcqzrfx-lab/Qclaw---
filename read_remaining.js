const fs = require('fs');
const path = require('path');

// 读取我的公文包中剩余的重要文件夹
const bases = [
    { name: '信息流数据分享', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\信息流数据分享' },
    { name: '数据/u群后台数据表', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\u群后台数据表' },
    { name: '数据/青少渠道数据', path: 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\数据\\青少渠道数据' }
];

bases.forEach(b => {
    if (fs.existsSync(b.path)) {
        const files = fs.readdirSync(b.path);
        console.log('\n=== ' + b.name + ' ===');
        files.forEach(f => console.log(f));
    }
});
