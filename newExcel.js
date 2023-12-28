const ExcelJS = require('D:/kirim-pdf-ke-telegram/node_modules/exceljs');

const path = process.cwd();
// Membuat file Excel pertama
const workbook1 = new ExcelJS.Workbook();
const worksheet1 = workbook1.addWorksheet('Sheet1');
// Tambahkan data ke worksheet pertama (contoh)
worksheet1.addRow(['Path']);

const excelFileName1 = path+'\\reportFailed.xlsx';

workbook1.xlsx.writeFile(excelFileName1)
  .then(function() {
    console.log(`File Excel pertama '${excelFileName1}' telah dibuat.`);
  })
  .catch(function(error) {
    console.error(`Terjadi kesalahan saat membuat file Excel pertama: ${error}`);
  });

// Membuat file Excel kedua
const workbook2 = new ExcelJS.Workbook();
const worksheet2 = workbook2.addWorksheet('Sheet1');
// Tambahkan data ke worksheet kedua (contoh)
worksheet2.addRow(['Result']);


const excelFileName2 = path+'\\resultCase.xlsx';

workbook2.xlsx.writeFile(excelFileName2)
  .then(function() {
    console.log(`File Excel kedua '${excelFileName2}' telah dibuat.`);
  })
  .catch(function(error) {
    console.error(`Terjadi kesalahan saat membuat file Excel kedua: ${error}`);
  });
