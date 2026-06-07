/**
 * Generates .xlsx versions of all sample CSV files.
 * Run with: node sample-data/generate-excel.js
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const dir = __dirname;

const csvFiles = [
  'sales-by-month.csv',
  'market-share.csv',
  'population-gdp.csv',
  'product-ratings.csv',
  'temperature-trend.csv',
  'study-hours-scores.csv',
];

csvFiles.forEach(csvFile => {
  const csvPath = path.join(dir, csvFile);
  const xlsxFile = csvFile.replace('.csv', '.xlsx');
  const xlsxPath = path.join(dir, xlsxFile);

  const workbook = XLSX.readFile(csvPath);
  XLSX.writeFile(workbook, xlsxPath);
  console.log(`Created: ${xlsxFile}`);
});

console.log('\nAll Excel files generated successfully.');
