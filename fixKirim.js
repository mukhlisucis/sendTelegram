const XLSX = require('D:/kirim-pdf-ke-telegram/node_modules/xlsx');
const TelegramBot = require('D:/kirim-pdf-ke-telegram/node_modules/node-telegram-bot-api');
const path = process.cwd();

const token = 'TOKEN ANDA';
const bot = new TelegramBot(token, { polling: true });
const user_id = 'USER ID ANDA';

const delayAfter = 20; // Tambahkan penundaan setelah setiap 20 permintaan
const delayDuration = 5000; // 5 detik penundaan, sesuaikan kebutuhan Anda

let totalFilesSent = 0; // Menyimpan total jumlah file yang berhasil dikirim

function bacaDataDariExcel(namaFileExcel) {
  const workbook = XLSX.readFile(namaFileExcel);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const data = [];

  let i = 2; 
  let cellAddress = `A${i}`;

  while (worksheet[cellAddress]) {
    data.push(worksheet[cellAddress].v); 
    i++;
    cellAddress = `A${i}`;
  }

  return data;
}

function kirimFilePDF(alamatFilePDF) {
  const fileOptions = {
    filename: '',
    contentType: 'application/pdf',
  };

  return new Promise((resolve, reject) => {
    bot.sendDocument(user_id, alamatFilePDF, {}, fileOptions)
      .then(() => {
        console.log('File PDF telah dikirim.');
        totalFilesSent++; // Increment total file yang berhasil dikirim
        resolve();
      })
      .catch((error) => {
        if (error.response && error.response.statusCode === 429) {
          // Jika terjadi rate limit, tambahkan penundaan dan coba lagi
          console.warn('Rate limit terdeteksi. Menunggu beberapa saat sebelum mencoba lagi...');
          setTimeout(() => {
            kirimFilePDF(alamatFilePDF).then(resolve).catch(reject);
          }, delayDuration);
        } else {
          console.error('Terjadi kesalahan saat mengirim file PDF:', error);
          reject(error);
        }
      });
  });
}

const dataExcel = bacaDataDariExcel(path + '\\reportFailed.xlsx');

const promises = dataExcel.map((pdfPath, index) => {
  const promise = kirimFilePDF(pdfPath);

  if ((index + 1) % delayAfter === 0 && index + 1 !== dataExcel.length) {
    // Tambahkan penundaan setelah setiap `delayAfter` permintaan, kecuali untuk batch terakhir
    return promise.then(() => new Promise((resolve) => setTimeout(resolve, delayDuration)));
  }

  return promise;
});

Promise.all(promises)
  .then(() => {
    console.log(`Total ${totalFilesSent} file berhasil dikirim.`);
    console.log('Tidak ada file yang dikirim lagi.');
    process.exit(0);
  })
  .catch((error) => {
    console.error('Terjadi kesalahan saat mengirim pesan:', error);
    process.exit(1);
  });
