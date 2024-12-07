const { spawn } = require('child_process');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

/**
 * Convert Excel worksheet data to Parquet format
 * @param {string} excelFilePath - Path to the Excel file
 * @param {string} sheetName - Name of the worksheet to convert
 * @param {string} outputPath - Path where the Parquet file will be saved
 * @returns {Promise<string>} - Path to the created Parquet file
 */
async function convertExcelToParquet(excelFilePath, sheetName, outputPath) {
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn('python', [
            'converter.py',
            'convert',
            excelFilePath,
            sheetName,
            outputPath
        ]);

        let result = '';

        pythonProcess.stdout.on('data', (data) => {
            result += data.toString();
        });

        pythonProcess.stderr.on('data', (data) => {
            console.error(`Python Error: ${data}`);
        });

        pythonProcess.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(`Python process exited with code ${code}`));
                return;
            }

            try {
                const response = JSON.parse(result);
                if (response.success) {
                    resolve(outputPath);
                } else {
                    reject(new Error(response.message));
                }
            } catch (error) {
                reject(new Error('Failed to parse Python response'));
            }
        });
    });
}

/**
 * Read data from a Parquet file
 * @param {string} parquetFilePath - Path to the Parquet file
 * @returns {Promise<Array>} - Array of objects containing the Parquet file data
 */
async function readParquetFile(parquetFilePath) {
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn('python', [
            'converter.py',
            'read',
            parquetFilePath
        ]);

        let result = '';

        pythonProcess.stdout.on('data', (data) => {
            result += data.toString();
        });

        pythonProcess.stderr.on('data', (data) => {
            console.error(`Python Error: ${data}`);
        });

        pythonProcess.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(`Python process exited with code ${code}`));
                return;
            }

            try {
                const data = JSON.parse(result);
                resolve(data);
            } catch (error) {
                reject(new Error('Failed to parse Python response'));
            }
        });
    });
}

module.exports = {
    convertExcelToParquet,
    readParquetFile
};
