const { exec } = require('child_process');

// Function to run a script
function runScript(scriptName) {
    return new Promise((resolve, reject) => {
        console.log(`Running ${scriptName}...`);
        exec(`node ${scriptName}`, (error, stdout, stderr) => {
            if (error) {
                console.error(`Error running ${scriptName}:`, error.message);
                reject(error);
                return;
            }
            if (stderr) {
                console.error(`Error output from ${scriptName}:`, stderr);
            }
            console.log(`Output from ${scriptName}:\n`, stdout);
            resolve();
        });
    });
}

// Main function to run scripts sequentially
async function main() {
    try {
        await runScript('01_renameExcel.js');
        await runScript('02_mergepage.js');
        console.log('Both scripts finished successfully.');
    } catch (error) {
        console.error('Error during script execution:', error.message);
    }
}

main();
