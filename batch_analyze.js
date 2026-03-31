const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';
const output = [];

function log(msg) {
    console.log(msg);
    output.push(msg);
}

function analyzeLeads(filePath, fileName) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(sheetName);
        
        if (!data || data.length === 0) return;
        
        log(`\n${'='.repeat(60)}`);
        log(`📊 ${fileName}`);
        log(`${'='.repeat(60)}`);
        log(`总记录数: ${data.length}条\n`);
        
        // Get headers
        const headers = Object.keys(data[0]);
        
        // Key metrics
        const metrics = {};
        
        // Analyze key columns
        const keyCols = ['产品名称', '获客方式', '投放计划', '投放账户名称', '代理商简称', '落地页', '流量池', '城市'];
        
        keyCols.forEach(col => {
            if (headers.includes(col)) {
                const count = {};
                data.forEach(row => {
                    const val = row[col] || '未知';
                    const vals = val.toString().split(',').map(v => v.trim()).filter(v => v);
                    vals.forEach(v => {
                        count[v] = (count[v] || 0) + 1;
                    });
                });
                
                const sorted = Object.entries(count).sort((a,b) => b[1] - a[1]).slice(0, 8);
                if (sorted.length > 0) {
                    log(`📌 ${col}分布:`);
                    sorted.forEach(([k, v]) => {
                        log(`   ${k}: ${v}条 (${(v/data.length*100).toFixed(1)}%)`);
                    });
                    log('');
                }
            }
        });
        
        return true;
    } catch (e) {
        log(`Error analyzing ${fileName}: ${e.message}`);
        return false;
    }
}

function analyzeFlow(filePath, fileName) {
    try {
        const workbook = XLSX.readFile(filePath);
        
        log(`\n${'='.repeat(60)}`);
        log(`📈 ${fileName}`);
        log(`${'='.repeat(60)}`);
        
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet);
            
            if (!data || data.length === 0) return;
            
            log(`\n📋 工作表: ${sheetName} (${data.length}行)`);
            
            const headers = Object.keys(data[0]);
            log(`列数: ${headers.length}`);
            
            // Show first few columns
            const sampleCols = headers.slice(0, 10);
            log(`主要列: ${sampleCols.join(' | ')}`);
            
            // Show first 3 rows
            log('\n数据样本:');
            data.slice(0, 3).forEach((row, i) => {
                const vals = sampleCols.map(c => row[c] || '-');
                log(`  Row${i+1}: ${vals.join(' | ')}`);
            });
        });
        
        return true;
    } catch (e) {
        log(`Error: ${e.message}`);
        return false;
    }
}

function analyzeGoals(filePath, fileName) {
    try {
        const workbook = XLSX.readFile(filePath);
        
        log(`\n${'='.repeat(60)}`);
        log(`🎯 ${fileName}`);
        log(`${'='.repeat(60)}`);
        
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet);
            
            if (!data || data.length === 0) return;
            
            log(`\n📋 工作表: ${sheetName}`);
            
            const headers = Object.keys(data[0]);
            log(`列: ${headers.join(' | ')}`);
            
            log('\n数据样本:');
            data.slice(0, 5).forEach((row, i) => {
                const vals = headers.slice(0, 8).map(h => row[h] || '-');
                log(`  ${vals.join(' | ')}`);
            });
        });
        
        return true;
    } catch (e) {
        log(`Error: ${e.message}`);
        return false;
    }
}

// Main analysis
log('🔍 开始批量分析Excel文件...');
log(`时间: ${new Date().toLocaleString()}\n`);

// 1. 线索明细数据 files
const leadFiles = [
    '线索明细数据 (13).xlsx',
    '线索明细数据 (0211期).xlsx',
    '线索明细数据 0205使用.xlsx',
    '线索明细数据.xlsx',
    '线索明细数据to新雨.xlsx'
];

leadFiles.forEach(f => {
    const fp = path.join(desktop, f);
    if (fs.existsSync(fp)) {
        analyzeLeads(fp, f);
    }
});

// 2. 信息流分析
analyzeFlow(path.join(desktop, '信息流分析-头条&广点通-2月&3月.xlsx'), '信息流分析-头条&广点通-2月&3月.xlsx');

// 3. 城市等级
analyzeFlow(path.join(desktop, '【春二轮&春暑秋】-城市P8等级表.xlsx'), '【春二轮&春暑秋】-城市P8等级表.xlsx');

// 4. 收款目标
analyzeGoals(path.join(desktop, '星光项目部2026年收款及效率目标_执行版0311.xlsx'), '星光项目部2026年收款及效率目标_执行版0311.xlsx');

// 5. 城市分级
analyzeGoals(path.join(desktop, 'city_level_5_251122.xlsx'), 'city_level_5_251122.xlsx');

log('\n' + '='.repeat(60));
log('✅ 分析完成!');
log('='.repeat(60));

// Save to file
fs.writeFileSync(path.join(desktop, 'analysis_report.txt'), output.join('\n'), 'utf8');
console.log('\n报告已保存到桌面: analysis_report.txt');
