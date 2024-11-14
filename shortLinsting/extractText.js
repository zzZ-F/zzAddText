const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const type = '法语';
const types = ['英语', '法语', '西班牙语']


// 定义变量
let folderName = '';
let column1 = '';
let column2 = '';

if(type === '英语') {
    folderName = '1114';
    column1 = '英语短句1';
    column2 = '英语短句2';
}
if(type === '法语') {
    folderName = '1114';
    column1 = '法语短句1';
    column2 = '法语短句2';
}
if(type === '西班牙语') {
    folderName = '1114';
    column1 = '西班牙语短句1';
    column2 = '西班牙语短句2';
}


// 文件路径
const filePath = path.join(process.env.HOME, `Desktop/EnglishListening/${folderName}/逐字稿.xlsx`);
const outputDir = path.join(__dirname, 'output');
const outputFile = path.join(outputDir, 'tmp.txt');

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
    if (row[column1] && row[column2]) {
        outputText += `${row[column1]} \n`;
        outputText += `${row[column2]} \n`;
    }
});

// 写入到文本文件
fs.writeFileSync(outputFile, outputText, 'utf8');
console.log('内容已成功输出到 tmp.txt 文件中');
