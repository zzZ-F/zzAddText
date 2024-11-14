const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// 文件路径
const filePath = path.join(process.env.HOME, 'Desktop/EnglishListening/1114/逐字稿.xlsx');
// const outputDir = path.join(process.env.HOME, 'Desktop/EnglishListening/1114/output');
const outputDir = path.join(__dirname, 'output');

const outputFile = path.join(outputDir, 'processed_sentences.txt');

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
    if (row['英语']) {
        const words = row['英语'].split(' ');
        const processedWords = words.map((word, i) => {
            // 在最后一个单词后加((⏱️=3000))
            return i === words.length - 1 ? `${word}((⏱️=3000))\n` : `${word}\n`;
        }).join('');
        outputText += processedWords;

        // 每 10 句添加一个额外的换行分隔
        if ((index + 1) % 10 === 0) {
            outputText += '\n';
        }
    }
});

// 写入到文本文件
fs.writeFileSync(outputFile, outputText, 'utf8');
console.log('内容已成功输出到 processed_sentences.txt 文件中');
