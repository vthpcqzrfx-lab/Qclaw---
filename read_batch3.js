const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取绩效各年度文件
const perfBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\绩效';
const years = ['2022年绩效', '2023年绩效', '2024年绩效', '2025年绩效'];

years.forEach(year => {
    const dir = path.join(perfBase, year);
    if (fs.existsSync(dir)) {
        const files = fs.readdirSync(dir);
        console.log('\n=== ' + year + ' ===');
        files.forEach(f => console.log('  ' + f));
    }
});

// 读取会议文件列表（更多）
const meetingBase = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\会议';
const meetings = fs.readdirSync(meetingBase);
console.log('\n=== 全部会议文件 ===');
meetings.forEach(m => {
    const mPath = path.join(meetingBase, m);
    const stat = fs.statSync(mPath);
    if (stat.isDirectory()) {
        const sub = fs.readdirSync(mPath);
        console.log(m + '/ (' + sub.length + '个文件)');
        sub.slice(0, 3).forEach(s => console.log('  - ' + s));
    } else {
        console.log(m);
    }
});
