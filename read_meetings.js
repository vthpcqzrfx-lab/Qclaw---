const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 读取会议文件夹的Excel文件
const base = 'C:\\Users\\zhaoyang26\\Desktop\\我的公文包\\会议';
const meetingDirs = fs.readdirSync(base).filter(f => {
    const fPath = path.join(base, f);
    return fs.statSync(fPath).isDirectory();
});

console.log('=== 会议文件夹 ===');
meetingDirs.forEach(d => {
    const dPath = path.join(base, d);
    const files = fs.readdirSync(dPath);
    const excels = files.filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    console.log(d + ': ' + files.length + '个文件，' + excels.length + '个Excel');
});

// 读取几个关键会议的Excel
const keyMeetings = [
    '20221010Q3晋升会/附件2-晋升答辩汇总名单-【私域运营】.xlsx',
    '20230905三六零人员盘点/3【战略规划&组织架构调整计划】.xlsx',
    '20230905三六零人员盘点/4【盘点信息收集表】-哓丹.xlsx'
];

keyMeetings.forEach(f => {
    const fPath = path.join(base, f);
    if (fs.existsSync(fPath)) {
        try {
            const workbook = XLSX.readFile(fPath);
            console.log('\n=== ' + f.split('/').pop() + ' ===');
            console.log('Sheets:', workbook.SheetNames);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
            console.log('Total rows:', data.length);
        } catch (e) {
            console.log('Error:', e.message);
        }
    }
});
