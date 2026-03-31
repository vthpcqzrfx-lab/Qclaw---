const fs = require('fs');
const path = require('path');

// 读取1月份落地页
const base1 = 'C:\\Users\\zhaoyang26\\Desktop\\1月份落地页';
const files1 = fs.readdirSync(base1);
console.log('=== 1月份落地页 ===');
files1.forEach(f => console.log(f));

// 读取信息流招聘
const base2 = 'C:\\Users\\zhaoyang26\\Desktop\\信息流招聘';
const files2 = fs.readdirSync(base2);
console.log('\n=== 信息流招聘 ===');
files2.forEach(f => console.log(f));

// 读取信息流结算
const base3 = 'C:\\Users\\zhaoyang26\\Desktop\\信息流结算';
const files3 = fs.readdirSync(base3);
console.log('\n=== 信息流结算 ===');
files3.forEach(f => console.log(f));

// 读取剑桥英语分析
const base4 = 'C:\\Users\\zhaoyang26\\Desktop\\剑桥英语分析';
const files4 = fs.readdirSync(base4);
console.log('\n=== 剑桥英语分析 ===');
files4.forEach(f => console.log(f));

// 读取学而思流量调研
const base5 = 'C:\\Users\\zhaoyang26\\Desktop\\学而思流量调研';
const files5 = fs.readdirSync(base5);
console.log('\n=== 学而思流量调研 ===');
files5.forEach(f => console.log(f));

// 读取公司资质材料
const base6 = 'C:\\Users\\zhaoyang26\\Desktop\\公司资质材料';
const files6 = fs.readdirSync(base6);
console.log('\n=== 公司资质材料 ===');
files6.forEach(f => console.log(f));

// 读取给代理的剑桥课程介绍
const base7 = 'C:\\Users\\zhaoyang26\\Desktop\\给代理的剑桥课程介绍';
const files7 = fs.readdirSync(base7);
console.log('\n=== 给代理的剑桥课程介绍 ===');
files7.forEach(f => console.log(f));

// 读取商品头图
const base8 = 'C:\\Users\\zhaoyang26\\Desktop\\商品头图';
const files8 = fs.readdirSync(base8);
console.log('\n=== 商品头图 ===');
files8.forEach(f => console.log(f));
