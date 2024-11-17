const xlsx = require("xlsx");

function processExcel(filePath, outputPath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const headers = data[0];
    const rows = data.slice(1);

    // 从整句中提取短句的函数
    function extractShortPhrase(mainText, shortPhrase) {
        const mainWords = mainText.split(/\s+/); // 按空格拆分整句
        const shortWords = shortPhrase.split(/\s+/); // 按空格拆分短句

        return mainWords.slice(0, shortWords.length).join(" "); // 截取对应长度的单词并拼接
    }

    rows.forEach((row) => {
        // 跳过中文及中文短句的处理
        if (row[2] && row[4] && row[6]) {
            row[4] = row[4]; // 中文短句1
            row[6] = row[6]; // 中文短句2
        }

        // 处理英语、法语、西班牙语
        [["英语", 1, 3, 5], ["法语", 7, 8, 9], ["西班牙语", 10, 11, 12]].forEach(
            ([lang, mainColIndex, short1Index, short2Index]) => {
                const mainText = row[mainColIndex]?.trim() || "";
                const shortPhrase1 = row[short1Index]?.trim() || "";
                const shortPhrase2 = row[short2Index]?.trim() || "";

                if (mainText && shortPhrase1) {
                    // 根据整句重新提取短句1
                    row[short1Index] = extractShortPhrase(mainText, shortPhrase1);

                    // 根据整句重新提取短句2
                    if (shortPhrase2) {
                        const remainingMainText = mainText
                            .split(/\s+/)
                            .slice(shortPhrase1.split(/\s+/).length)
                            .join(" "); // 获取短句1后的部分
                        row[short2Index] = extractShortPhrase(remainingMainText, shortPhrase2);
                    }
                }
            }
        );
    });

    // 写回新的 Excel
    const newSheetData = [headers, ...rows];
    const newSheet = xlsx.utils.aoa_to_sheet(newSheetData);
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, sheetName);

    xlsx.writeFile(newWorkbook, outputPath);
    console.log(`文件已处理并保存到: ${outputPath}`);
}

const inputFilePath = "逐字稿.xlsx"; // 输入文件路径
const outputFilePath = "output.xlsx"; // 输出文件路径

processExcel(inputFilePath, outputFilePath);
