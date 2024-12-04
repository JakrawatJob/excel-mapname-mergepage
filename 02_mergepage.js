const fs = require('fs').promises;
const path = require('path');
const ExcelJS = require('exceljs');

const truthRenameFolderPath = path.join(__dirname, './Folder/01_TruthRename');
const folderMergeExcelPath = path.join(__dirname, './Folder/02_TruthMerge');

async function ensureDirectoryExists(dirPath) {
    try {
        await fs.mkdir(dirPath, { recursive: true });
    } catch (err) {
        if (err.code !== 'EEXIST') {
            throw err;
        }
    }
}

async function excelToJson(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];

    const jsonData = [];
    let columnMask = null;

    sheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
        const rowData = [];
        let hasData = false;

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
                hasData = true;
                if (rowIndex === 1) {
                    if (!columnMask) columnMask = [];
                    columnMask[colNumber] = true;
                }
            }
            if (!columnMask || columnMask[colNumber]) {
                rowData.push(cell.value);
            }
        });

        if (hasData) {
            jsonData.push(rowData);
        }
    });

    return jsonData;
}

async function jsonToExcel(jsonData, outputFilePath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('MergedData');

    jsonData.forEach((row) => {
        worksheet.addRow(row);
    });

    worksheet.columns.forEach((column) => {
        let maxWidth = 10;
        column.eachCell({ includeEmpty: true }, (cell) => {
            if (cell.value) {
                const textLength = String(cell.value).length;
                maxWidth = Math.max(maxWidth, textLength + 2);
            }
        });
        column.width = maxWidth;
    });

    await workbook.xlsx.writeFile(outputFilePath);
}

async function main() {
    try {
        await ensureDirectoryExists(folderMergeExcelPath);
        const subfolders = await fs.readdir(truthRenameFolderPath);

        for (const subfolder of subfolders) {
            const subfolderPath = path.join(truthRenameFolderPath, subfolder);

            const stat = await fs.lstat(subfolderPath);
            if (!stat.isDirectory()) continue;

            const newSubfolderPath = path.join(folderMergeExcelPath, subfolder);
            await ensureDirectoryExists(newSubfolderPath);

            const files = await fs.readdir(subfolderPath);
            const fileGroups = {};

            files.forEach(file => {
                if (!file.toLowerCase().includes('.xlsx')) return;
                const match = file.match(/^(.*)_page(\d+)\.xlsx$/i);
                if (match) {
                    const baseName = match[1];
                    const pageNum = parseInt(match[2], 10);
                    if (!fileGroups[baseName]) fileGroups[baseName] = [];
                    fileGroups[baseName].push({ file, pageNum });
                }
            });

            for (const [baseName, groupFiles] of Object.entries(fileGroups)) {
                groupFiles.sort((a, b) => a.pageNum - b.pageNum);

                const mergedJson = [];
                let headersAdded = false;

                for (const { file } of groupFiles) {
                    const filePath = path.join(subfolderPath, file);
                    const jsonData = await excelToJson(filePath);

                    if (!headersAdded) {
                        mergedJson.push(...jsonData);
                        headersAdded = true;
                    } else {
                        mergedJson.push(...jsonData.slice(1));
                    }
                }

                const newFileName = `${baseName}.xlsx`;
                const newFilePath = path.join(newSubfolderPath, newFileName);
                await jsonToExcel(mergedJson, newFilePath);
                console.log(`Merged files for ${baseName} into ${newFilePath}`);
            }
        }
    } catch (err) {
        console.error(err);
    }
}

main();
