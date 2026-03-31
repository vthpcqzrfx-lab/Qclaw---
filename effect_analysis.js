const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('📊 关键指标交叉分析');
console.log('='.repeat(70));

const workbook = XLSX.readFile(path.join(desktop, '信息流分析-头条&广点通-2月&3月.xlsx'));
const sheet = workbook.Sheets['信息流投放明细'];
const data = XLSX.utils.sheet_to_json(sheet);

// 按产品分析投放效果
console.log('\n📈 各产品投放效果分析');
console.log('-'.repeat(70));

const productStats = {};
data.forEach(row => {
    const product = row['产品'] || '未知';
    if (!productStats[product]) {
        productStats[product] = {
            impressions: 0,
            clicks: 0,
            leads: 0,
            cost: 0
        };
    }
    
    const impression = parseFloat(row['展示']) || 0;
    const click = parseFloat(row['点击']) || 0;
    const lead = parseFloat(row['线索数']) || 0;
    const c = parseFloat(row['现金消耗-元']) || 0;
    
    productStats[product].impressions += impression;
    productStats[product].clicks += click;
    productStats[product].leads += lead;
    productStats[product].cost += c;
});

console.log('产品 | 展示 | 点击 | 点击率 | 线索数 | 消耗(元) | CPC | CPL');
console.log('-'.repeat(70));

Object.entries(productStats)
    .filter(([k,v]) => k !== '未知' && v.impressions > 0)
    .sort((a,b) => b[1].leads - a[1].leads)
    .slice(0, 15)
    .forEach(([product, stats]) => {
        const ctr = stats.impressions > 0 ? (stats.clicks / stats.impressions * 100).toFixed(2) + '%' : '0%';
        const cpc = stats.clicks > 0 ? (stats.cost / stats.clicks).toFixed(2) : '0';
        const cpl = stats.leads > 0 ? (stats.cost / stats.leads).toFixed(2) : '0';
        const pname = product.length > 15 ? product.substring(0,15) + '...' : product;
        console.log(`${pname} | ${(stats.impressions/10000).toFixed(1)}万 | ${(stats.clicks/1000).toFixed(1)}千 | ${ctr} | ${stats.leads} | ${(stats.cost/10000).toFixed(1)}万 | ${cpc} | ${cpl}`);
    });

// 按媒体分析
console.log('\n\n📊 各媒体渠道效果对比');
console.log('-'.repeat(70));

const mediaStats = {};
data.forEach(row => {
    const media = row['流量池'] || '未知';
    if (!mediaStats[media]) {
        mediaStats[media] = { impressions: 0, clicks: 0, leads: 0, cost: 0 };
    }
    
    mediaStats[media].impressions += parseFloat(row['展示']) || 0;
    mediaStats[media].clicks += parseFloat(row['点击']) || 0;
    mediaStats[media].leads += parseFloat(row['线索数']) || 0;
    mediaStats[media].cost += parseFloat(row['现金消耗-元']) || 0;
});

console.log('媒体 | 展示 | 点击 | 点击率 | 线索数 | 消耗 | CPC | CPL');
Object.entries(mediaStats).forEach(([media, stats]) => {
    const ctr = stats.impressions > 0 ? (stats.clicks / stats.impressions * 100).toFixed(2) + '%' : '0%';
    const cpc = stats.clicks > 0 ? (stats.cost / stats.clicks).toFixed(2) : '0';
    const cpl = stats.leads > 0 ? (stats.cost / stats.leads).toFixed(2) : '0';
    console.log(`${media} | ${(stats.impressions/10000).toFixed(1)}万 | ${(stats.clicks/1000).toFixed(1)}千 | ${ctr} | ${stats.leads} | ${(stats.cost/10000).toFixed(1)}万 | ${cpc} | ${cpl}`);
});

// 优化师效果排名
console.log('\n\n🏆 优化师效果排行榜 (Top 10)');
console.log('-'.repeat(70));

const optimizerStats = {};
data.forEach(row => {
    const opt = row['投放人名称'] || row['管理人'] || '未知';
    if (!optimizerStats[opt]) {
        optimizerStats[opt] = { impressions: 0, clicks: 0, leads: 0, cost: 0 };
    }
    
    optimizerStats[opt].impressions += parseFloat(row['展示']) || 0;
    optimizerStats[opt].clicks += parseFloat(row['点击']) || 0;
    optimizerStats[opt].leads += parseFloat(row['线索数']) || 0;
    optimizerStats[opt].cost += parseFloat(row['现金消耗-元']) || 0;
});

console.log('优化师 | 线索数 | 消耗(元) | CPL | 点击率 | 每万展示线索');
console.log('-'.repeat(70));

Object.entries(optimizerStats)
    .filter(([k,v]) => v.leads > 0)
    .sort((a,b) => b[1].leads - a[1].leads)
    .slice(0, 10)
    .forEach(([opt, stats]) => {
        const cpl = stats.leads > 0 ? (stats.cost / stats.leads).toFixed(2) : '0';
        const ctr = stats.impressions > 0 ? (stats.clicks / stats.impressions * 100).toFixed(2) + '%' : '0%';
        const leadsPerW = stats.impressions > 0 ? (stats.leads / stats.impressions * 10000).toFixed(2) : '0';
        const optName = opt.length > 10 ? opt.substring(0,10) : opt;
        console.log(`${optName} | ${stats.leads} | ${stats.cost.toFixed(0)} | ${cpl} | ${ctr} | ${leadsPerW}`);
    });

console.log('\n' + '='.repeat(70));
