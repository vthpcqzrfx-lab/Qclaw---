const XLSX = require('xlsx');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

console.log('='.repeat(70));
console.log('📊 各期线索数据对比分析');
console.log('='.repeat(70));

// 分析多个线索文件
const leadFiles = [
    { name: '线索明细数据 (13).xlsx', label: '3月最新' },
    { name: '线索明细数据 (0211期).xlsx', label: '2月11期' },
    { name: '线索明细数据 0205使用.xlsx', label: '2月5期' },
    { name: '线索明细数据.xlsx', label: '综合' }
];

leadFiles.forEach(file => {
    const fp = path.join(desktop, file.name);
    try {
        const wb = XLSX.readFile(fp);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet);
        
        console.log(`\n📌 ${file.name}`);
        console.log(`   记录数: ${data.length}条`);
        
        if (data.length > 0) {
            // 产品分布
            const products = {};
            data.forEach(row => {
                const p = row['产品名称'] || '未知';
                products[p] = (products[p] || 0) + 1;
            });
            console.log('   产品分布:');
            Object.entries(products).slice(0, 3).forEach(([k,v]) => {
                console.log(`     - ${k.substring(0,40)}: ${v}条`);
            });
            
            // 渠道分布
            const channels = {};
            data.forEach(row => {
                const c = row['获客方式'] || '未知';
                channels[c] = (channels[c] || 0) + 1;
            });
            console.log('   渠道:');
            Object.entries(channels).forEach(([k,v]) => {
                console.log(`     - ${k}: ${v}条`);
            });
        }
    } catch (e) {
        console.log(`   读取失败: ${e.message}`);
    }
});

console.log('\n\n' + '='.repeat(70));
console.log('📊 城市等级分布分析');
console.log('='.repeat(70));

// 分析城市等级
const cityWB = XLSX.readFile(path.join(desktop, 'city_level_5_251122.xlsx'));
const citySheet = cityWB.Sheets['Sheet1'];
const cityData = XLSX.utils.sheet_to_json(citySheet);

const levelCount = {};
cityData.forEach(row => {
    const level = row['city_level_p'] || '未知';
    levelCount[level] = (levelCount[level] || 0) + 1;
});

console.log('\n城市等级分布:');
Object.entries(levelCount).sort((a,b) => b[1]-a[1]).forEach(([level, count]) => {
    console.log(`  ${level}: ${count}个城市`);
});

console.log('\n\n' + '='.repeat(70));
console.log('📊 收款目标分析 (星光项目部)');
console.log('='.repeat(70));

// 分析收款目标
const goalWB = XLSX.readFile(path.join(desktop, '星光项目部2026年收款及效率目标_执行版0311.xlsx'));
const goalSheet = goalWB.Sheets['Sheet2'];
const goalData = XLSX.utils.sheet_to_json(goalSheet);

console.log('\n月度预算与退费率:');
console.log('月份 | 渠道 | 市场费用 | 退费率');
console.log('-'.repeat(50));
goalData.forEach(row => {
    console.log(`${row['月份']} | ${row['一级渠道']} | ${row['市场费用']} | ${(row['退费率']*100).toFixed(1)}%`);
});

console.log('\n\n' + '='.repeat(70));
console.log('✅ 分析完成');
console.log('='.repeat(70));
