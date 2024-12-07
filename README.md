# Excel to JSON Converter

A web application that allows users to upload Excel files and convert them to JSON format. The application supports both single and multiple worksheet Excel files.

## Features

- Excel file upload
- Support for multiple worksheets
- Automatic conversion to JSON
- Data preview in tabular format
- Simple and intuitive user interface

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Amitkumar4920/ExcelToParquetConvertor.git
```

2. Install dependencies:
```bash
npm install
```

3. Start the application:
```bash
npm start
```

The application will be available at `http://localhost:3001`

## Usage

1. Open your web browser and navigate to `http://localhost:3001`
2. Click "Choose File" to select an Excel file
3. If the file has multiple worksheets, select the one you want to convert
4. The data will be converted and displayed in a table format

## Dependencies

- Express.js
- Multer
- xlsx
- EJS

## Directory Structure

- `/uploads` - Temporary storage for uploaded Excel files
- `/output` - Storage for converted JSON files
- `/views` - EJS templates for the web interface

## Contributing

Feel free to submit issues and enhancement requests.
