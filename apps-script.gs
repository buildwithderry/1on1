/**
 * Penerima lead untuk landing page Sesi 1 on 1.
 *
 * Cara pasang:
 * 1. Buka spreadsheet leads -> Ekstensi -> Apps Script, tempel file ini.
 * 2. Ganti SHEET_NAME kalau nama tab bukan "leads".
 * 3. Terapkan -> Deployment baru -> Jenis: Aplikasi web
 *    Jalankan sebagai: Saya | Akses: Siapa saja
 * 4. Salin URL /exec ke CONFIG.endpoint di index.html.
 *
 * Kolom: Timestamp | Nama | No HP | Kebutuhan | Sumber | Harga | Waktu Klien | Halaman
 */
var SHEET_NAME = 'leads';
var TZ = 'Asia/Jakarta';

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    var row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 8).setValues([[
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      String(d.nama || '').slice(0, 80),
      String(d.hp || ''),
      String(d.kebutuhan || '').slice(0, 800),
      String(d.sumber || ''),
      Number(d.harga) || '',
      String(d.waktuKlien || ''),
      String(d.halaman || '')
    ]]);
    sh.getRange(row, 3).setNumberFormat('@'); // nomor HP tetap teks (+62...)
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  }
}
