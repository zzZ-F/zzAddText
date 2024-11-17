const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const folderName = '1114';
const type = '法语';
const types = ['英语', '法语', '西班牙语'];

// 定义变量
let columnName = '';

if (type === '英语') {
    columnName = '英语';
} else if (type === '法语') {
    columnName = '法语';
} else if (type === '西班牙语') {
    columnName = '西班牙语';
}

// 文件路径
const filePath = path.join(process.env.HOME, `Desktop/EnglishListening/${folderName}/逐字稿.xlsx`);
const outputDir = path.join(__dirname, 'output');
const outputFile = path.join(outputDir, 'words_sentences.txt');

// 读取 Excel 文件
const workbook = xlsx.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]]; // 获取第一个工作表
const data = xlsx.utils.sheet_to_json(sheet);

// 确保输出目录存在
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// 处理数据
let outputText = '';
data.forEach((row, index) => {
    if (row[columnName]) {
        const words = row[columnName].split(' ');
        const processedWords = words.map((word, i) => {
            // 在最后一个单词后加((⏱️=3000))，但每第十句时不加
            return i === words.length - 1 && (index + 1) % 10 !== 0
                ? `${word}((⏱️=3000))\n`
                : `${word}\n`;
        }).join('');
        outputText += processedWords;

        // 每 10 句添加一个额外的换行分隔
        if ((index + 1) % 10 === 0) {
            outputText += '\n\n\n';
        }
    }
});

// 写入到文本文件
fs.writeFileSync(outputFile, outputText, 'utf8');
console.log('内容已成功输出到 words_sentences.txt 文件中');
