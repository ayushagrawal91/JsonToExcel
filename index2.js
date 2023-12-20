const fs = require('fs');
const XLSX = require('xlsx');

function ConvertToExcel(data, outputFilename = 'output.xlsx') {
    // /** Converts an array of JS objects to a worksheet. */
    // const worksheet = XLSX.utils.json_to_sheet(data);
    // /** Converts an array of JS objects to a worksheet. */
    // const workBook = XLSX.utils.book_new();
    // /** Append a worksheet to a workbook */
    // XLSX.utils.book_append_sheet(workBook, worksheet, 'Sheet 1');
    // XLSX.writeFile(workBook, outputFilename);
    // console.log(`Data exported to ${outputFilename}`);
    try {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workBook, worksheet, 'Sheet 1');
    XLSX.writeFile(workBook, outputFilename);
    console.log(`Data exported to ${outputFilename}`);
  } catch (error) {
    console.error(`Error occurred during Excel conversion: ${error.message}`);}
}

function main() {

    const jsonData = JSON.parse(fs.readFileSync('./data.json', 'utf-8'));
    ConvertToExcel(jsonData.employees, 'output.xlsx');
}

main();
