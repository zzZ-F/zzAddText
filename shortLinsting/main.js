const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

// Load the Excel file
const filePath = path.join(__dirname, '逐字稿.xlsx'); // Replace with the path to your Excel file
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0]; // Assuming there is only one sheet
const worksheet = workbook.Sheets[sheetName];

// Parse the worksheet
const data = xlsx.utils.sheet_to_json(worksheet);

// Create the output folder if it doesn't exist
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

// Prepare the content for writing
let content = '';

data.forEach((row, index) => {
    const originalSentence = row['原始句子'];   // Column: 原始句子
    const sentencePart1 = row['拆分短句 1'];    // Column: 拆分短句 1
    const sentencePart2 = row['拆分短句 2'];    // Column: 拆分短句 2

    // 1. Original sentence with [Ǩ:1] and [Ǩ:5]
    content += `${originalSentence}[Ǩ:1]\n`;
    content += `${originalSentence}[Ǩ:5]\n`;

    // 2. Split the sentence into words and output each word individually
    originalSentence.split(' ').forEach(word => {
        content += `${word.replace(',', '').replace('.', '')}\n`;
    });

    // content += `${originalSentence.split(' ').join(' ')}[Ǩ:2]\n`;

    // 3. Sentence part 1 repeated four times with [Ǩ:3]
    for (let i = 0; i < 4; i++) {
        content += `${sentencePart1}[Ǩ:3]\n`;
    }

    // 4. Sentence part 2 repeated four times with [Ǩ:3]
    for (let i = 0; i < 4; i++) {
        content += `${sentencePart2}[Ǩ:3]\n`;
    }

    // 5. Original sentence repeated four times with [Ǩ:4]
    for (let i = 0; i < 4; i++) {
        content += `${originalSentence}[Ǩ:4]\n`;
    }

    // Add a separator or newline to distinguish between entries
    content += '\n';
});

// Write the content to content.txt
const outputPath = path.join(outputDir, 'content.txt');
fs.writeFileSync(outputPath, content, 'utf8');

console.log('Processing complete. Check the output folder for content.txt');
