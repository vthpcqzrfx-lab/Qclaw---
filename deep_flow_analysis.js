const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('🔍 深度分析 - 信息流投放明细 (70,000+条记录)');
console.log('='.repeat(70));

const workbook = XLSX.readFile(path.join(desktop, '信息流分析-头条&广点通-2月&3月.xlsx'));
const sheet = workbook.Sheets['信息流投放明细'];
const data = XLSX.utils.sheet_to_json(sheet);

console.log(`\n总投放记录数: ${data.length}条\n`);

// 1. 流量池分布
console.log('📊 1. 流量池分布 (媒体渠道)');
const poolCount = {};
data.forEach(row => {
    const pool = row['流量池'] || '未知';
    poolCount[pool] = (poolCount[pool] || 0) + 1;
});
Object.entries(poolCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`   ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 2. 产品分布
console.log('\n📊 2. 产品分布 (Top 15)');
const productCount = {};
data.forEach(row => {
    const p = row['产品'] || '未知';
    productCount[p] = (productCount[p] || 0) + 1;
});
Object.entries(productCount).sort((a,b) => b[1]-a[1]).slice(0,15).forEach(([k,v]) => {
    console.log(`   ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 3. 投放人/优化师
console.log('\n📊 3. 投放人/优化师分布 (Top 15)');
const optimizerCount = {};
data.forEach(row => {
    const o = row['投放人名称'] || row['管理人'] || '未知';
    optimizerCount[o] = (optimizerCount[o] || 0) + 1;
});
Object.entries(optimizerCount).sort((a,b) => b[1]-a[1]).slice(0,15).forEach(([k,v]) => {
    console.log(`   ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 4. 投放账户
console.log('\n📊 4. 投放账户分布 (Top 15)');
const accountCount = {};
data.forEach(row => {
    const a = row['投放账户名称'] || '未知';
    accountCount[a] = (accountCount[a] || 0) + 1;
});
Object.entries(accountCount).sort((a,b) => b[1]-a[1]).slice(0,15).forEach(([k,v]) => {
    const shortName = k.length > 40 ? k.substring(0,40) + '...' : k;
    console.log(`   ${shortName}: ${v}条`);
});

// 5. 落地页
console.log('\n📊 5. 落地页分布 (Top 10)');
const landingCount = {};
data.forEach(row => {
    const l = row['落地页名称'] || '未知';
    landingCount[l] = (landingCount[l] || 0) + 1;
});
Object.entries(landingCount).sort((a,b) => b[1]-a[1]).slice(0,10).forEach(([k,v]) => {
    const shortName = k.length > 45 ? k.substring(0,45) + '...' : k;
    console.log(`   ${shortName}: ${v}条`);
});

// 6. 投放计划
console.log('\n📊 6. 投放计划分布 (Top 10)');
const planCount = {};
data.forEach(row => {
    const p = row['投放计划名称'] || '未知';
    planCount[p] = (planCount[p] || 0) + 1;
});
Object.entries(planCount).sort((a,b) => b[1]-a[1]).slice(0,10).forEach(([k,v]) => {
    const shortName = k.length > 45 ? k.substring(0,45) + '...' : k;
    console.log(`   ${shortName}: ${v}条`);
});

// 7. 是否搜索 vs 通投
console.log('\n📊 7. 搜索 vs 通投分布');
const searchCount = {};
data.forEach(row => {
    const s = row['是否是搜索'] || '未知';
    searchCount[s] = (searchCount[s] || 0) + 1;
});
Object.entries(searchCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`   ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

console.log('\n' + '='.repeat(70));
console.log('分析完成!');
console.log('='.repeat(70));
