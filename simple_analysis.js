const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('📊 关键指标交叉分析');
console.log('='.repeat(70));

const workbook = XLSX.readFile(path.join(desktop, '信息流分析-头条&广点通-2月&3月.xlsx'));
const sheet = workbook.Sheets['信息流投放明细'];
const data = XLSX.utils.sheet_to_json(sheet);

console.log('\n总记录数:', data.length);

// 检查列名
console.log('\n列名:', Object.keys(data[0]).join(', '));

// 按产品分析投放效果
console.log('\n\n📈 各产品投放效果分析');
console.log('-'.repeat(70));

const productStats = {};
data.forEach(row => {
    const product = row['产品'] || '未知';
    if (!productStats[product]) {
        productStats[product] = {
            count: 0,
            impressions: 0,
            clicks: 0,
            leads: 0,
            cost: 0
        };
    }
    
    productStats[product].count++;
    productStats[product].impressions += parseFloat(row['展示']) || 0;
    productStats[product].clicks += parseFloat(row['点击']) || 0;
    productStats[product].leads += parseFloat(row['线索数']) || 0;
    productStats[product].cost += parseFloat(row['现金消耗-元']) || 0;
});

Object.keys(productStats).slice(0, 5).forEach(p => {
    console.log(p, productStats[p]);
});

// 简化输出
console.log('\n产品 | 记录数 | 展示 | 点击 | 线索 | 消耗');
Object.entries(productStats)
    .filter(([k,v]) => k !== '未知')
    .sort((a,b) => b[1].leads - a[1].leads)
    .slice(0, 10)
    .forEach(([product, stats]) => {
        console.log(`${product} | ${stats.count} | ${stats.impressions} | ${stats.clicks} | ${stats.leads} | ${stats.cost.toFixed(0)}`);
    });

console.log('\n完成');
