const TelegramBot = require('D:/kirim-pdf-ke-telegram/node_modules/node-telegram-bot-api');
const xlsx = require('D:/kirim-pdf-ke-telegram/node_modules/node-xlsx');
const fs = require('fs');
const path = process.cwd();
const bot = new TelegramBot('TOKEN ANDA', { polling: true });

const workSheetsFromFile = xlsx.parse(fs.readFileSync(path+'\\resultCase.xlsx'));

const sheetData = workSheetsFromFile[0].data;

let combinedMessage = '';

for (const row of sheetData) {
   const dataA = row[0];

   combinedMessage += dataA + '\n' + '\n';
}

bot.sendMessage('ID GROUP', combinedMessage)
  .then(() => {
    console.log('Pesan telah terkirim.');
    process.exit(0);
  })
  .catch((error) => {
    console.error('Terjadi kesalahan saat mengirim pesan:', error);
    process.exit(1);
  });
