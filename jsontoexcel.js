
const fs = require('fs');
const XLSX = require('xlsx');

//code for reading the json file
function readJsonFile(filePath) {
  try {
    const jsonData = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(jsonData);
  } catch (error) {
    console.error(`Error reading JSON file: ${error.message}`);
    process.exit(1);
  }
}

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

// code to convert json to excel
function jsonToExcel(jsonData, excelFileName) {
  const flatData = flattenJson(jsonData);

  const ws = XLSX.utils.json_to_sheet([flatData]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');

  XLSX.writeFile(wb, excelFileName);
}

// declare the path of the file
const jsonFilePath = './sample.json';
// name the excel file 
const excelFileName = 'output_excel_file.xlsx';

//reading the json file
const jsonData = readJsonFile(jsonFilePath);

// converting the json file to excel sheet
jsonToExcel(jsonData, excelFileName);

// message to verify that the execution is done successfully.
console.log(`Conversion successful. Excel file saved as ${excelFileName}.`);
