const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

// Load the Excel file
const filePath = path.join(__dirname, '长的听力.xlsx'); // Replace with your file path
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0]; // Assuming we only have one sheet
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
    const englishSentence = row['英文']; // Accessing the '英文' column

    // Repeat and modify the sentence
    content += `${englishSentence} [Ǩ:2]\n`;
    content += `${englishSentence} [Ǩ:5]\n`;
    content += `${englishSentence} [Ǩ:2]\n`;
    content += `${englishSentence} [Ǩ:2]\n\n`;
});

// Write the content to content.txt
const outputPath = path.join(outputDir, 'content.txt');
fs.writeFileSync(outputPath, content, 'utf8');

console.log('Processing complete. Check the output folder for content.txt');
