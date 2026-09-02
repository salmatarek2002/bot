const puppeteer = require('puppeteer');
const fs = require('fs');
const ExcelJS = require('exceljs');

const HOME_URL = 'https://eg1xbet.com/ar';

const ACCOUNTS = [
  // 1. Alex
  { username: "salma.alex.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.alex.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },

  // 2. Assuit
  { username: "salma.assuit.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.assuit.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
// 3. Aswan & Behaira & Beni
  { username: "salma.aswan.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.aswan.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.aswan.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.behaira.16@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.beni.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },

  // 4. Cairo
  { username: "salma.cairo.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.16@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.17@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.18@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.19@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.20@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.21@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.22@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.23@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.24@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.25@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.26@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.27@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.28@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.29@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.30@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.cairo.31@gmail.com", password: "ssrytw2492002", name: "Tarek" },
// 5. Dakha & Damitta & Fayoum
  { username: "salma.dakha.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.16@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.17@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.dakha.18@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.damitta.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.fayoum.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },

  // 6. Ganwabsinaa & Gharbia & Giza & Ismail
  { username: "salma.ganwabsinaa.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ganwabsinaa.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ganwabsinaa.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ganwabsinaa.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ganwabsinaa.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ganwabsinaa.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.gharbia.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.16@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.17@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.18@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.giza.19@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ismail.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ismail.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ismail.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ismail.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.ismail.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
// 7. Kafr & Luxor & Matroh & Menouf & Minya
  { username: "salma.kafr.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.kafr.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.luxor.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.luxor.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.matroh.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.matroh.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.matroh.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.matroh.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.menouf.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.minya.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },

  // 8. Northsinaa & Port & Qalioub & Qena & Red
  { username: "salma.northsinaa.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.northsinaa.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.northsinaa.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.northsinaa.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.port.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.port.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.port.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qalioub.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.qena.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.red.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },

  // 9. Sharqia & Sohag & Suez & T & Wadi
  { username: "salma.sharqia.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.14@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.15@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.16@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.17@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.18@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.19@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.20@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sharqia.21@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.05@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.06@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.07@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.08@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.09@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.10@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.11@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.12@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.sohag.13@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.suez.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.suez.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.suez.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.suez.04@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.t.2492002@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.wadi.01@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.wadi.02@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.wadi.03@gmail.com", password: "ssrytw2492002", name: "Tarek" },
  { username: "salma.wadi.04@gmail.com", password: "ssrytw2492002", name: "Tarek" }
];

const SEEN_PHONE_NUMBERS = new Set();

const OPEN_LOGIN_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(5) > div > div > div > div > button';
const USERNAME_SELECTOR = '#username';
const PASSWORD_SELECTOR = '#username-password';
const SUBMIT_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(5) > div > div > div > div.ui-inline-dropdown__content > div > div > div > div > form > button';

const NAME_INPUT_SELECTOR = '#__WELCOME_APP__ > div.default-layout-container > div > div > section > main > div > div > div > div > div.user-verify-common-type__form.user-verify-common-type-form > div.ui-input-base-default--theme-default.ui-input-base-default--size-m.ui-input-base-default.ui-input.user-verify-common-type-form__input.ui-input-base.ui-input.user-verify-common-type-form__input > div > div.ui-input-base-default__content > input';

const DEPOSIT_BTN_SELECTOR = '#app > div.layout-content.layout-content--theme-primary--40 > div.layout-content__header > header > div.header-top.header__top > div.user-control-dashboard--padding-right.user-control-dashboard.header-top__controls--outside.header-top__controls > div:nth-child(4) > a';
const VODAFONE_CASH_SELECTOR = '#vodafone_1';
const ORANGE_CASH_SELECTOR = '.orange_egypt_bt_webdefault';
const PHONE_NUMBER_SELECTOR = '#payment_modal_container > div.payment_modal_body > form > div:nth-child(2) > div > span.modal-message-address';

const EXCEL_FILE_PATH = './accounts_phones.xlsx';

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function getFormattedFileName(phoneNumber, isDuplicate = false) {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  
  const cleanPhone = phoneNumber ? phoneNumber.replace(/[^0-9]/g, '') : 'UNKNOWN_PHONE';
  const dupTag = isDuplicate ? '_DUPLICATED' : '';

  return `${cleanPhone}_${year}-${month}-${day}_${hours}-${minutes}${dupTag}.jpg`;
}

function getCurrentDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}
async function appendToExcel(phoneNumber) {
  if (!phoneNumber) return;

  const cleanPhone = phoneNumber.replace(/[^0-9]/g, '');
  const formattedDateTime = getCurrentDateTime();
  const workbook = new ExcelJS.Workbook();
  let worksheet;

  if (fs.existsSync(EXCEL_FILE_PATH)) {
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);
    worksheet = workbook.getWorksheet('Results');
    
    let exists = false;
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && row.getCell(1).value == cleanPhone) {
        exists = true;
      }
    });

    if (exists) {
      console.log(`ℹ️ الرقم (${cleanPhone}) موجود مسبقاً في الإكسل، لن يتم إضافته مرة أخرى.`);
      return;
    }
  } else {
    worksheet = workbook.addWorksheet('Results');
    worksheet.addRow(['رقم الموبايل (Phone)', 'وقت الاستخراج (Date & Time)']);
    worksheet.getColumn(1).width = 25;
    worksheet.getColumn(2).width = 25;
  }

  worksheet.addRow([cleanPhone, formattedDateTime]);
  await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
  console.log(`📝 تم حفظ الرقم الفريد (${cleanPhone}) مع الوقت (${formattedDateTime}) في ملف الإكسل بنجاح.`);
}

async function extractPhoneNumber(browser) {
  const pages = await browser.pages();
  
  for (let i = pages.length - 1; i >= 0; i--) {
    const page = pages[i];
    
    for (const frame of page.frames()) {
      try {
        await frame.waitForSelector(PHONE_NUMBER_SELECTOR, { timeout: 1500 }).catch(() => {});
        const phoneText = await frame.evaluate((sel) => {
          const el = document.querySelector(sel);
          return el ? (el.innerText || el.textContent || el.value || '').trim() : null;
        }, PHONE_NUMBER_SELECTOR);

        if (phoneText && phoneText.length > 5) {
          return phoneText;
        }
      } catch (e) {}
    }

    try {
      await page.waitForSelector(PHONE_NUMBER_SELECTOR, { timeout: 1500 }).catch(() => {});
      const phoneText = await page.evaluate((sel) => {
        const el = document.querySelector(sel);
        return el ? (el.innerText || el.textContent || el.value || '').trim() : null;
      }, PHONE_NUMBER_SELECTOR);

      if (phoneText && phoneText.length > 5) {
        return phoneText;
      }
    } catch (e) {}
  }

  return null;
}

async function processAccount(account, index, total) {
  console.log(`\n[${index + 1}/${total}] 🔑 جاري معالجة: ${account.username}`);

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: {
      width: 1920,
      height: 1080,
      deviceScaleFactor: 2
    },
    args: ['--start-maximized', '--no-sandbox', '--disable-setuid-sandbox']
  });

  try {
    const pages = await browser.pages();
    let page = pages[0];

    await page.goto(HOME_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });

    await page.waitForSelector(OPEN_LOGIN_BTN_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(OPEN_LOGIN_BTN_SELECTOR);
    await delay(1000);

    await page.waitForSelector(USERNAME_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(USERNAME_SELECTOR);
    await delay(300);
    await page.type(USERNAME_SELECTOR, account.username, { delay: 50 });

    await page.click(PASSWORD_SELECTOR);
    await delay(300);
    await page.type(PASSWORD_SELECTOR, account.password, { delay: 50 });
    await delay(500);

    try {
      await page.click(SUBMIT_BTN_SELECTOR);
    } catch (e) {
      await page.keyboard.press('Enter');
    }

    await delay(3000);

    try {
      const nameInput = await page.waitForSelector(NAME_INPUT_SELECTOR, { timeout: 4000 });
      if (nameInput) {
        console.log(`👤 ظهرت صفحة الاسم! جاري إدخال: ${account.name}`);
        await nameInput.click();
        await delay(300);
        await nameInput.type(account.name, { delay: 50 });
        await delay(500);
        await page.keyboard.press('Enter');
        await delay(3000);
      }
    } catch (err) {}

    await page.waitForSelector(DEPOSIT_BTN_SELECTOR, { visible: true, timeout: 30000 });
    await page.click(DEPOSIT_BTN_SELECTOR);

    await delay(6000);

    // 1. خطوة فودافون
    let allPages = await browser.pages();
    page = allPages[allPages.length - 1];

    let clicked = false;
    for (const frame of page.frames()) {
      try {
        const el = await frame.$(VODAFONE_CASH_SELECTOR);
        if (el) {
          await el.scrollIntoViewIfNeeded();
          await el.click();
          clicked = true;
          break;
        }
      } catch (e) {}
    }

    if (!clicked) {
      await page.waitForSelector(VODAFONE_CASH_SELECTOR, { visible: true, timeout: 15000 }).catch(() => {});
      await page.evaluate((sel) => {
        const btn = document.querySelector(sel);
        if (btn) {
          btn.scrollIntoView();
          btn.click();
        }
      }, VODAFONE_CASH_SELECTOR);
    }

    await delay(8000);

    const phoneNumber = await extractPhoneNumber(browser);
    let isDuplicate = false;

    if (phoneNumber) {
      const cleanPhone = phoneNumber.replace(/[^0-9]/g, '');
      if (SEEN_PHONE_NUMBERS.has(cleanPhone)) {
        isDuplicate = true;
        console.warn(`⚠️ تنبيه: رقم فودافون (${phoneNumber}) مكرر خلال جلسة التنفيذ الحالية!`);
      } else {
        SEEN_PHONE_NUMBERS.add(cleanPhone);
      }

      let finalPages = await browser.pages();
      let activePage = finalPages[finalPages.length - 1];
      let fileName = getFormattedFileName(phoneNumber, isDuplicate);

      await activePage.screenshot({
        path: fileName,
        type: 'jpeg',
        quality: 100,
        fullPage: false
      });

      console.log(`📞 رقم فودافون المستخرج: ${phoneNumber}`);
      console.log(`📸 تم التقاط اللقطة باسم: ${fileName}`);

      await appendToExcel(phoneNumber);
    } else {
      console.log(`❌ لم يتم العثور على رقم فودافون.`);
    }

    // 2. خطوة أورانج (بشرط 010) - بالتعديل الآمن لمنع الإيرور
    console.log("جاري الانتقال والدفع عبر أورانج...");
    
    let orangeClicked = false;
    allPages = await browser.pages();
    page = allPages[allPages.length - 1];

    for (const frame of page.frames()) {
      try {
        const clickedInFrame = await frame.evaluate((sel) => {
          const btn = document.querySelector(sel);
          if (btn) {
            btn.scrollIntoView();
            btn.click();
            return true;
          }
          return false;
        }, ORANGE_CASH_SELECTOR);

        if (clickedInFrame) {
          orangeClicked = true;
          break;
        }
      } catch (e) {}
    }

    if (!orangeClicked) {
      try {
        await page.waitForSelector(ORANGE_CASH_SELECTOR, { visible: true, timeout: 10000 });
        await page.evaluate((sel) => {
          const btn = document.querySelector(sel);
          if (btn) {
            btn.scrollIntoView();
            btn.click();
          }
        }, ORANGE_CASH_SELECTOR);
        orangeClicked = true;
      } catch (e) {
        console.log("⚠️ زرار أورانج لم يستجب بالطريقة العادية، جاري محاولة التخطي أو المتابعة...");
      }
    }

    console.log("⏳ جاري انتظار ظهور نافذة ورقم أورانج...");
    await delay(7000);

    const orangePhoneNumber = await extractPhoneNumber(browser);

    if (orangePhoneNumber) {
      const cleanOrangePhone = orangePhoneNumber.replace(/[^0-9]/g, '');

      if (cleanOrangePhone.startsWith('010')) {
        let isOrangeDuplicate = false;
        if (SEEN_PHONE_NUMBERS.has(cleanOrangePhone)) {
          isOrangeDuplicate = true;
          console.warn(`⚠️ تنبيه: رقم أورانج (${orangePhoneNumber}) مكرر خلال جلسة التنفيذ الحالية!`);
        } else {
          SEEN_PHONE_NUMBERS.add(cleanOrangePhone);
        }

        let finalPagesOrange = await browser.pages();
        let activePageOrange = finalPagesOrange[finalPagesOrange.length - 1];
        let orangeFileName = getFormattedFileName(orangePhoneNumber, isOrangeDuplicate);

        await activePageOrange.screenshot({
          path: orangeFileName,
          type: 'jpeg',
          quality: 100,
          fullPage: false
        });

        console.log(`📞 رقم أورانج المستخرج (يبدأ بـ 010): ${orangePhoneNumber}`);
        console.log(`📸 تم التقاط اللقطة باسم: ${orangeFileName}`);

        await appendToExcel(orangePhoneNumber);
      } else {
        console.log(`🚫 تم تجاهل رقم أورانج (${orangePhoneNumber}) لأنه لا يبدأ بـ 010.`);
      }
    } else {
      console.log(`❌ لم يتم العثور على رقم أورانج أو أن النافذة لم تفتح.`);
    }

  } catch (error) {
    console.log(`❌ حدث خطأ مع الحساب (${account.username}): ${error.message}`);
  } finally {
    await browser.close();
  }
}

(async () => {
  console.log(`🚀 بدء تنفيذ العملية لـ ${ACCOUNTS.length} حساب...`);

  for (let i = 0; i < ACCOUNTS.length; i++) {
    await processAccount(ACCOUNTS[i], i, ACCOUNTS.length);
    await delay(2000);
  }

  console.log(`\n🎉 اكتملت العملية تماماً لجميع الحسابات وتوقف السكربت!`);
})();