const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

// Load the Excel file
const filePath = path.join(__dirname, '逐字稿.xlsx'); // Replace with your file path
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

data.forEach((row) => {
    const question = row['问句'];   // Accessing "问句"
    const word1 = row['单词'];      // Accessing "单词" (first one)
    const word2 = row['单词2'];     // Accessing second "单词" (if named differently)
    const exampleSentence = row['例句'];  // Accessing "例句"

    // Process each column and append the specified suffixes
    content += `${question}[Ǩ:3]\n`;
    content += `${word1}[Ǩ:1]\n`;
    content += `${word2}[Ǩ:1]\n`;
    content += `${exampleSentence}[Ǩ:1]\n`;

    // Check if word1 or word2 are present in the example sentence
    if (!exampleSentence.toLowerCase().includes(word1.toLowerCase()) || !exampleSentence.toLowerCase().includes(word2.toLowerCase())) {
        console.log('内容不匹配:', row);
    }

    // Add a separator or newline to distinguish between entries
    content += '\n';
});

// Write the content to content.txt
const outputPath = path.join(outputDir, 'content.txt');
fs.writeFileSync(outputPath, content, 'utf8');

console.log('Processing complete. Check the output folder for content.txt');
