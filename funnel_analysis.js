const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('🔍 素材管理分析');
console.log('='.repeat(70));

try {
    const wb = XLSX.readFile(path.join(desktop, '工作簿-素材.xlsx'));
    
    wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);
        
        if (data.length > 0) {
            console.log(`\n📋 Sheet: ${sheetName} (${data.length}条)`);
            
            // 显示列名
            const headers = Object.keys(data[0]);
            console.log('   列:', headers.slice(0, 8).join(' | '));
            
            // 显示样本
            if (data.length >= 2) {
                console.log('   样本1:', Object.values(data[0]).slice(0, 6).join(' | '));
                console.log('   样本2:', Object.values(data[1]).slice(0, 6).join(' | '));
            }
        }
    });
} catch(e) {
    console.log('素材文件读取失败:', e.message);
}

console.log('\n\n' + '='.repeat(70));
console.log('🔍 按媒体计算完整转化漏斗');
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
            signup: 0,
            attend: 0,
            cost: 0
        };
    }
    
    funnel[media].impressions += parseFloat(row['展示']) || 0;
    funnel[media].clicks += parseFloat(row['点击']) || 0;
    funnel[media].leads += parseFloat(row['线索数']) || 0;
    funnel[media].friends += parseFloat(row['加好友数']) || 0;
    funnel[media].signup += parseFloat(row['引流课报名人数']) || 0;
    funnel[media].attend += parseFloat(row['到课人数']) || 0;
    funnel[media].cost += parseFloat(row['现金消耗-元']) || 0;
});

console.log('\n媒体 | 展示 | 点击→线索→加好友→报名→到课 | 消耗');
Object.entries(funnel).forEach(([media, stats]) => {
    console.log(`${media}:`);
    console.log(`  展示: ${(stats.impressions/10000).toFixed(1)}万`);
    console.log(`  点击: ${(stats.clicks/1000).toFixed(1)}千 (${(stats.clicks/stats.impressions*100).toFixed(2)}%)`);
    console.log(`  线索: ${stats.leads} (${(stats.leads/stats.clicks*100).toFixed(1)}%)`);
    console.log(`  加好友: ${stats.friends} (${(stats.friends/stats.leads*100).toFixed(1)}%)`);
    console.log(`  报名: ${stats.signup} (${(stats.signup/stats.friends*100).toFixed(1)}%)`);
    console.log(`  到课: ${stats.attend} (${(stats.attend/stats.signup*100).toFixed(1)}%)`);
    console.log(`  消耗: ${(stats.cost/10000).toFixed(1)}万`);
    console.log('');
});

console.log('\n' + '='.repeat(70));
console.log('✅ 分析完成');
