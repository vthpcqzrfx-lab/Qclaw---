const XLSX = require('xlsx');
const path = require('path');
const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

const files = [
    '线索明细数据 (13).xlsx',
    '线索明细数据 0205使用.xlsx',
    '线索明细数据.xlsx'
];

let workbook = null;
for (const f of files) {
    const filePath = path.join(desktop, f);
    try {
        workbook = XLSX.readFile(filePath);
        console.log('Loaded:', f);
        break;
    } catch (e) {
        console.log('Failed:', f, e.message);
    }
}

if (!workbook) {
    console.log('No file loaded');
    process.exit(1);
}

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('\n=== 文件概览 ===');
console.log('工作表数量:', workbook.SheetNames.length);
console.log('工作表名称:', workbook.SheetNames.join(', '));

const rowCount = data.length;
const colCount = data.length > 0 ? data[0].length : 0;
console.log('\n=== 数据规模 ===');
console.log('总行数:', rowCount);
console.log('总列数:', colCount);

console.log('\n=== 表头 (前20列) ===');
if (data.length > 0) {
    for (let col = 0; col < Math.min(20, colCount); col++) {
        console.log(`  ${col + 1}: ${data[0][col] || ''}`);
    }
}

console.log('\n=== 前5行数据样本 (前10列) ===');
for (let row = 1; row < Math.min(6, rowCount); row++) {
    const rowData = [];
    for (let col = 0; col < Math.min(10, colCount); col++) {
        rowData.push(data[row][col] || '-');
    }
    console.log(`Row${row + 1}:`, rowData.join(' | '));
}

// Simple statistics
console.log('\n=== 关键列数据分析 ===');
if (data.length > 1) {
    const headers = data[0];
    for (let col = 0; col < Math.min(15, colCount); col++) {
        const header = headers[col] || '';
        let numCount = 0;
        let numSum = 0;
        for (let row = 1; row < Math.min(100, rowCount); row++) {
            const val = data[row][col];
            if (typeof val === 'number') {
                numCount++;
                numSum += val;
            } else if (typeof val === 'string' && val.match(/^[\d,.]+$/)) {
                const num = parseFloat(val.replace(/,/g, ''));
                if (!isNaN(num)) {
                    numCount++;
                    numSum += num;
                }
            }
        }
        if (numCount > 50) {
            const avg = numSum / numCount;
            console.log(`  Column ${col + 1} [${header}]: 有效数值行=${numCount}, 平均值=${avg.toFixed(2)}`);
        }
    }
}
