const ExcelJS = require('exceljs');
const fs = require('fs');
const { exec } = require('child_process');
const path = require('path');

const TELEGRAM_BOT_TOKEN = '8562177835:AAErdCmz3911yx-x2fqsE6kOyW04E6cSSSU';
const CHAT_ID = '1136692275';
const EXCEL_FILE_PATH = './accounts_phones.xlsx';
const TEMP_FILE_PATH = './temp_accounts_phones.xlsx';
const LAST_PHONE_FILE = './last_phone.txt';

async function sendNewNumbers() {
  console.log('--- بدأ الفحص في: ' + new Date().toLocaleTimeString() + ' ---');
  
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    console.log('ملف الإكسيل غير موجود!');
    return;
  }

  try {
    // حل مشكلة فتح الملف: بنسخ الملف لنسخة مؤقتة ونقرا منها لتجنب الـ Lock
    fs.copyFileSync(EXCEL_FILE_PATH, TEMP_FILE_PATH);
  } catch (err) {
    console.log('الملف الأساسي مفتوح، جاري المحاولة...');
  }

  let lastSentPhone = "";
  if (fs.existsSync(LAST_PHONE_FILE)) {
    lastSentPhone = fs.readFileSync(LAST_PHONE_FILE, 'utf8').trim();
  }

  const workbook = new ExcelJS.Workbook();
  const fileToRead = fs.existsSync(TEMP_FILE_PATH) ? TEMP_FILE_PATH : EXCEL_FILE_PATH;
  
  await workbook.xlsx.readFile(fileToRead);
  const worksheet = workbook.getWorksheet('Results');
  
  let rows = [];
  if (worksheet) {
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && row.getCell(1).value) {
          rows.push(row.getCell(1).value.toString().trim());
      }
    });
  }

  const lastIndex = rows.indexOf(lastSentPhone);
  const newNumbers = lastIndex === -1 ? (rows.length > 0 ? [rows[rows.length-1]] : []) : rows.slice(lastIndex + 1);

  console.log(`إجمالي الأرقام بالشيت: ${rows.length} | الأرقام الجديدة غير المرسلة: ${newNumbers.length}`);

  if (newNumbers.length === 0) {
    console.log('لا توجد أرقام جديدة.');
    exec(`curl -s "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage?chat_id=${CHAT_ID}&text=${encodeURIComponent("ℹ️ لا توجد أرقام جديدة في هذا الفحص.")}"`);
  } else {
    // إرسال التقرير المجمع
    let summaryMessage = `🔔 تم العثور على ${newNumbers.length} أرقام جديدة:\n\n`;
    newNumbers.forEach(phone => summaryMessage += `📞 ${phone}\n`);
    exec(`curl -s "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage?chat_id=${CHAT_ID}&text=${encodeURIComponent(summaryMessage)}"`);

    // إرسال الصور منفصلة
    const filesInFolder = fs.readdirSync('./');
    for (const phone of newNumbers) {
      const imageFileName = filesInFolder.find(file => file.startsWith(phone) && file.toLowerCase().endsWith('.jpg'));
      if (imageFileName) {
        const imagePath = path.join('./', imageFileName);
        exec(`curl -s -F chat_id=${CHAT_ID} -F photo=@"${imagePath}" -F caption="📞 ${phone}" https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendPhoto`);
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
    }

    // تحديث ملف آخر رقم
    if (rows.length > 0) {
      fs.writeFileSync(LAST_PHONE_FILE, rows[rows.length - 1]);
      console.log('✅ تم تحديث ملف السجل بأحدث رقم: ' + rows[rows.length - 1]);
    }
  }

  // رسالة النهاية
  setTimeout(() => {
    exec(`curl -s "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage?chat_id=${CHAT_ID}&text=${encodeURIComponent("✅ انتهى البوت من فحص جميع الحسابات في هذه الدورة.")}"`);
    console.log('--- انتهى الفحص ---');
  }, 3000);
}

// التشغيل كل ربع ساعة والتنفيذ الفوري عند البدء
setInterval(sendNewNumbers, 15 * 60 * 1000);
sendNewNumbers();
console.log('البوت يعمل الآن بنظام النسخ المؤقت لتجنب تعارض الملف المفتوح...');