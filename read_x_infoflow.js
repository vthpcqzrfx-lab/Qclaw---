const fs = require('fs');
const path = require('path');

// 读取X业务线文档
const base1 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\X业务线文档';
const files1 = fs.readdirSync(base1);
console.log('=== X业务线文档 ===');
files1.forEach(f => console.log(f));

// 读取信息流数据分享
const base2 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\信息流数据分享';
const files2 = fs.readdirSync(base2);
console.log('\n=== 信息流数据分享 ===');
files2.forEach(f => console.log(f));

// 读取内部满意度调研问卷
const base3 = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\内部满意度调研问卷';
const files3 = fs.readdirSync(base3);
console.log('\n=== 内部满意度调研问卷 ===');
files3.forEach(f => console.log(f));
