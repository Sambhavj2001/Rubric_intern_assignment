
const fs = require('fs');
const XLSX = require('xlsx');

// Function to read JSON file
function readJsonFile(filePath) {
  try {
    const jsonData = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(jsonData);
  } catch (error) {
    console.error(`Error reading JSON file: ${error.message}`);
    process.exit(1);
  }
}

// Function to flatten nested JSON structure
function flattenJson(jsonObj, prefix = '') {
  let flatObj = {};
  for (const [key, value] of Object.entries(jsonObj)) {
    const newKey = prefix + key;
    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      flatObj = { ...flatObj, ...flattenJson(value, newKey + '_') };
    } else {
      flatObj[newKey] = value;
    }
  }
  return flatObj;
}

// Function to convert JSON to Excel
function jsonToExcel(jsonData, excelFileName) {
  // Flatten the nested JSON
  const flatData = flattenJson(jsonData);

  // Create a worksheet
  const ws = XLSX.utils.json_to_sheet([flatData]);

  // Create a workbook with the worksheet
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');

  // Write the workbook to an Excel file
  XLSX.writeFile(wb, excelFileName);
}

// Main script
// const jsonFilePath = 'path/to/your/nested/data.json';
const jsonFilePath = './sample.json';
const excelFileName = 'output_excel_file.xlsx';

// Read the JSON file
const jsonData = readJsonFile(jsonFilePath);

// Convert JSON to Excel
jsonToExcel(jsonData, excelFileName);

console.log(`Conversion successful. Excel file saved as ${excelFileName}.`);
