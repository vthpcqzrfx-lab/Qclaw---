const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('📊 转化漏斗分析');
console.log('='.repeat(70));

const workbook = XLSX.readFile(path.join(desktop, '信息流分析-头条&广点通-2月&3月.xlsx'));
const sheet = workbook.Sheets['信息流投放明细'];
const data = XLSX.utils.sheet_to_json(sheet);

// 按媒体统计完整漏斗
const funnel = {};

data.forEach(row => {
    const media = row['流量池'] || '未知';
    if (!funnel[media]) {
        funnel[media] = {
            impressions: 0,
            clicks: 0,
            leads: 0,
            friends: 0,
            attend: 0,
            cost: 0
        };
    }
    
    funnel[media].impressions += parseFloat(row['展示']) || 0;
    funnel[media].clicks += parseFloat(row['点击']) || 0;
    funnel[media].leads += parseFloat(row['线索数']) || 0;
    funnel[media].friends += parseFloat(row['加好友数']) || 0;
    funnel[media].attend += parseFloat(row['到课人数']) || 0;
    funnel[media].cost += parseFloat(row['现金消耗-元']) || 0;
});

console.log('\n漏斗数据:');
Object.entries(funnel).forEach(([media, stats]) => {
    console.log(`\n${media}:`);
    console.log(`  展示: ${(stats.impressions/10000).toFixed(1)}万`);
    console.log(`  点击: ${(stats.clicks/1000).toFixed(1)}千`);
    console.log(`  线索: ${stats.leads}`);
    console.log(`  加好友: ${stats.friends}`);
    console.log(`  到课: ${stats.attend}`);
    console.log(`  消耗: ${(stats.cost/10000).toFixed(1)}万`);
    
    // 计算转化率
    const ctr = stats.impressions > 0 ? (stats.clicks / stats.impressions * 100).toFixed(2) : 0;
    const leadRate = stats.clicks > 0 ? (stats.leads / stats.clicks * 100).toFixed(1) : 0;
    const friendRate = stats.leads > 0 ? (stats.friends / stats.leads * 100).toFixed(1) : 0;
    const attendRate = stats.friends > 0 ? (stats.attend / stats.friends * 100).toFixed(1) : 0;
    
    console.log(`  点击率: ${ctr}%`);
    console.log(`  线索率: ${leadRate}%`);
    console.log(`  加好友率: ${friendRate}%`);
    console.log(`  到课率: ${attendRate}%`);
});

// 按产品统计
console.log('\n\n' + '='.repeat(70));
console.log('📊 各产品转化漏斗');
console.log('='.repeat(70));

const productFunnel = {};
data.forEach(row => {
    const product = row['产品'] || '未知';
    if (!productFunnel[product]) {
        productFunnel[product] = {
            impressions: 0, clicks: 0, leads: 0, friends: 0, attend: 0, cost: 0
        };
    }
    
    productFunnel[product].impressions += parseFloat(row['展示']) || 0;
    productFunnel[product].clicks += parseFloat(row['点击']) || 0;
    productFunnel[product].leads += parseFloat(row['线索数']) || 0;
    productFunnel[product].friends += parseFloat(row['加好友数']) || 0;
    productFunnel[product].attend += parseFloat(row['到课人数']) || 0;
    productFunnel[product].cost += parseFloat(row['现金消耗-元']) || 0;
});

Object.entries(productFunnel)
    .filter(([k,v]) => k !== '未知' && v.leads > 100)
    .sort((a,b) => b[1].leads - a[1].leads)
    .slice(0, 6)
    .forEach(([product, stats]) => {
        const cpl = stats.leads > 0 ? (stats.cost / stats.leads).toFixed(0) : 0;
        const friendRate = stats.leads > 0 ? (stats.friends / stats.leads * 100).toFixed(0) : 0;
        const attendRate = stats.friends > 0 ? (stats.attend / stats.friends * 100).toFixed(0) : 0;
        
        console.log(`\n${product}:`);
        console.log(`  线索: ${stats.leads} | CPL: ${cpl}元`);
        console.log(`  加好友率: ${friendRate}% | 到课率: ${attendRate}%`);
    });

console.log('\n\n✅ 分析完成');
