const fs = require('fs').promises;
const path = require('path');

const jsonFilePath = path.join(__dirname, 'sony-label-db.json');
const truthFolderPath = path.join(__dirname, './Folder/Truth');
const truthRenameFolderPath = path.join(__dirname, './Folder/01_TruthRename');

async function readJSONFile(filePath) {
    const data = await fs.readFile(filePath, 'utf8');
    return JSON.parse(data);
}

async function ensureDirectoryExists(dirPath) {
    try {
        await fs.mkdir(dirPath, { recursive: true });
    } catch (err) {
        if (err.code !== 'EEXIST') {
            throw err;
        }
    }
}

async function main() {
    try {
        const jsonData = await readJSONFile(jsonFilePath);
        const idToFilenameMap = {};
        jsonData.forEach(entry => {
            idToFilenameMap[entry.id] = entry.filename;
        });

        let counter = 1;
        const subfolders = await fs.readdir(truthFolderPath);

        for (const subfolder of subfolders) {
            const subfolderPath = path.join(truthFolderPath, subfolder);
            const stat = await fs.lstat(subfolderPath);

            if (stat.isDirectory()) {
                const newSubfolderPath = path.join(truthRenameFolderPath, subfolder);
                await ensureDirectoryExists(newSubfolderPath);

                const files = await fs.readdir(subfolderPath);

                for (const file of files) {
                    const filePath = path.join(subfolderPath, file);
                    const ext = path.extname(file);
                    const baseName = path.basename(file, ext);
                    const id = parseInt(baseName, 10);

                    if (isNaN(id)) {
                        continue;
                    }

                    let newFilename = idToFilenameMap[id];
                    if (!newFilename) {
                        newFilename = `Unknown_page1`;
                    } else {
                        newFilename = path.basename(newFilename, path.extname(newFilename)) + '.xlsx';
                    }

                    const counterStr = String(counter).padStart(2, '0');
                    newFilename = ` ${newFilename}`;
                    const newFilePath = path.join(newSubfolderPath, newFilename);

                    await fs.copyFile(filePath, newFilePath);
                    counter++;
                }
            }
        }
    } catch (err) {
        console.error(err);
    }
}

main();
