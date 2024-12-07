const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.set('view engine', 'ejs');
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Ensure directories exist
const UPLOAD_DIR = path.join(__dirname, 'uploads');
const OUTPUT_DIR = path.join(__dirname, 'output');

[UPLOAD_DIR, OUTPUT_DIR].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir);
    }
});

// Function to convert Excel data to JSON and save
function saveExcelData(jsonData, outputPath) {
    fs.writeFileSync(outputPath, JSON.stringify(jsonData, null, 2));
}

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/upload', upload.single('excelFile'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).send('No file uploaded');
        }

        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;

        if (sheetNames.length === 1) {
            // If only one sheet, convert it directly
            const worksheet = workbook.Sheets[sheetNames[0]];
            const jsonData = xlsx.utils.sheet_to_json(worksheet);
            
            if (jsonData.length === 0) {
                throw new Error('No data in worksheet');
            }

            // Save as JSON file (temporary solution)
            const outputPath = path.join(OUTPUT_DIR, `${sheetNames[0]}.json`);
            saveExcelData(jsonData, outputPath);

            res.redirect('/view-data');
        } else {
            // If multiple sheets, let user choose
            res.render('select-sheet', { 
                sheets: sheetNames, 
                fileName: req.file.filename 
            });
        }
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error processing file: ' + error.message);
    }
});

app.post('/convert/:fileName', (req, res) => {
    try {
        const { sheetName } = req.body;
        const { fileName } = req.params;
        
        const filePath = path.join(UPLOAD_DIR, fileName);
        const workbook = xlsx.readFile(filePath);
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet);
        
        if (jsonData.length === 0) {
            throw new Error('No data in worksheet');
        }

        // Save as JSON file
        const outputPath = path.join(OUTPUT_DIR, `${sheetName}.json`);
        saveExcelData(jsonData, outputPath);

        res.redirect('/view-data');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error converting sheet: ' + error.message);
    }
});

app.get('/view-data', (req, res) => {
    try {
        const jsonFiles = fs.readdirSync(OUTPUT_DIR)
            .filter(file => file.endsWith('.json'));
        
        const allData = [];
        for (const file of jsonFiles) {
            const filePath = path.join(OUTPUT_DIR, file);
            const jsonContent = fs.readFileSync(filePath, 'utf8');
            const data = JSON.parse(jsonContent);
            
            allData.push({
                fileName: file,
                data: data
            });
        }
        
        res.render('view-data', { data: allData });
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error reading files: ' + error.message);
    }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
