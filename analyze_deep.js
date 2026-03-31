const XLSX = require('xlsx');
const path = require('path');
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

const workbook = XLSX.readFile(path.join(desktop, '线索明细数据 (13).xlsx'));
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet);

console.log('=== 线索数据深度分析报告 ===\n');
console.log(`总线索数: ${data.length}条\n`);

// 1. 产品分析
console.log('=== 1. 产品分布 ===');
const productCount = {};
data.forEach(row => {
    const products = row['产品名称'] || '';
    products.split(',').forEach(p => {
        p = p.trim();
        if(p) {
            productCount[p] = (productCount[p] || 0) + 1;
        }
    });
});
Object.entries(productCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`  ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 2. 获客方式
console.log('\n=== 2. 获客方式分布 ===');
const channelCount = {};
data.forEach(row => {
    const ch = row['获客方式'] || '未知';
    channelCount[ch] = (channelCount[ch] || 0) + 1;
});
Object.entries(channelCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`  ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 3. 投放计划
console.log('\n=== 3. 投放计划分布 (Top10) ===');
const planCount = {};
data.forEach(row => {
    const plan = row['投放计划'] || '未知';
    planCount[plan] = (planCount[plan] || 0) + 1;
});
Object.entries(planCount).sort((a,b) => b[1]-a[1]).slice(0,10).forEach(([k,v]) => {
    console.log(`  ${k}: ${v}条`);
});

// 4. 投放账户
console.log('\n=== 4. 投放账户分布 ===');
const accountCount = {};
data.forEach(row => {
    const acc = row['投放账户名称'] || '未知';
    accountCount[acc] = (accountCount[acc] || 0) + 1;
});
Object.entries(accountCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`  ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 5. 代理商
console.log('\n=== 5. 代理商分布 ===');
const agentCount = {};
data.forEach(row => {
    const agent = row['代理商简称'] || '未知';
    agentCount[agent] = (agentCount[agent] || 0) + 1;
});
Object.entries(agentCount).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
    console.log(`  ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
});

// 6. 时间分析
console.log('\n=== 6. 时间分析 ===');
const timeData = data.filter(row => row['投放时间']);
if(timeData.length > 0) {
    // 按小时分布
    const hourCount = {};
    timeData.forEach(row => {
        const time = row['投放时间'];
        if(time) {
            const hour = new Date(time).getHours();
            hourCount[hour] = (hourCount[hour] || 0) + 1;
        }
    });
    console.log('  投放时间小时分布:');
    for(let h=0; h<24; h++) {
        if(hourCount[h]) {
            console.log(`    ${h}:00-${h}:59: ${hourCount[h]}条`);
        }
    }
}

// 7. 落地页分析
console.log('\n=== 7. 落地页分布 (Top10) ===');
const landingCount = {};
data.forEach(row => {
    const lp = row['落地页'] || '未知';
    landingCount[lp] = (landingCount[lp] || 0) + 1;
});
Object.entries(landingCount).sort((a,b) => b[1]-a[1]).slice(0,10).forEach(([k,v]) => {
    console.log(`  ${k.substring(0,50)}: ${v}条`);
});

console.log('\n=== 分析完成 ===');
