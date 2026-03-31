const mammoth = require('mammoth');
const path = require('path');
const fs = require('fs');

const desktop = 'C:\\Users\\zhaoyang26\\Desktop';

const files = [
    '杂记.docx',
    '建议文档.doc',
    '文字库.doc',
    '客户档案功能服务开通确认函.docx',
    '视频著作权侵权和解及一次性赔偿协议.docx'
];

async function readDocx(filePath) {
    try {
        const result = await mammoth.extractRawText({ path: filePath });
        return result.value;
    } catch (e) {
        return `Error: ${e.message}`;
    }
}

async function readLegacyDoc(filePath) {
    // For .doc files, try to read as text with encoding detection
    try {
        // First try UTF-8
        let content = fs.readFileSync(filePath);
        
        // Try to detect if it's GBK (Chinese encoding)
        let text = content.toString('utf8');
        
        // Check for garbled characters (common pattern)
        if (text.includes('��') || text.includes('�')) {
            // Try GBK
            text = content.toString('gbk');
        }
        
        return text;
    } catch (e) {
        return `Error reading .doc: ${e.message}`;
    }
}

async function main() {
    console.log('='.repeat(60));
    console.log('📄 Word文档分析');
    console.log('='.repeat(60));

    for (const file of files) {
        const filePath = path.join(desktop, file);
        
        if (!fs.existsSync(filePath)) {
            console.log(`\n❌ 文件不存在: ${file}`);
            continue;
        }
        
        console.log(`\n\n${'='.repeat(60)}`);
        console.log(`📄 ${file}`);
        console.log('='.repeat(60));
        
        let content;
        if (file.endsWith('.docx')) {
            content = await readDocx(filePath);
        } else {
            content = await readLegacyDoc(filePath);
        }
        
        // Show first 3000 characters
        const preview = content.substring(0, 3000);
        console.log(preview);
        
        if (content.length > 3000) {
            console.log(`\n... (还有 ${content.length - 3000} 字符)`);
        }
    }
    
    console.log('\n\n' + '='.repeat(60));
    console.log('✅ 分析完成!');
    console.log('='.repeat(60));
}

main();
