const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

// 读取OVMOM年度目标文件
const ovmomDir = path.join(desktop, '我的公文包\\2024年ovmom');
if (fs.existsSync(ovmomDir)) {
    console.log('=== OVMOM年度目标文件 ===');
    fs.readdirSync(ovmomDir).forEach(f => {
        console.log(f);
    });
}

// 读取会议文件
const meetingDir = path.join(desktop, '我的公文包\\会议');
if (fs.existsSync(meetingDir)) {
    console.log('\n=== 会议文件 ===');
    fs.readdirSync(meetingDir).slice(0, 15).forEach(f => {
        console.log(f);
    });
}

// 读取X业务线文档
const xDir = path.join(desktop, '我的公文包\\X业务线文档');
if (fs.existsSync(xDir)) {
    console.log('\n=== X业务线文档 ===');
    fs.readdirSync(xDir).forEach(f => {
        console.log(f);
    });
}

// 读取绩效文件
const perfDir = path.join(desktop, '我的公文包\\绩效');
if (fs.existsSync(perfDir)) {
    console.log('\n=== 绩效文件 ===');
    fs.readdirSync(perfDir).forEach(f => {
        console.log(f);
    });
}

// 读取招聘文件
const hireDir = path.join(desktop, '我的公文包\\招聘');
if (fs.existsSync(hireDir)) {
    console.log('\n=== 招聘文件 ===');
    fs.readdirSync(hireDir).slice(0, 10).forEach(f => {
        console.log(f);
    });
}

// 读取出差报销
const travelDir = path.join(desktop, '我的公文包\\出差');
if (fs.existsSync(travelDir)) {
    console.log('\n=== 出差报销城市 ===');
    fs.readdirSync(travelDir).forEach(f => {
        console.log(f);
    });
}
