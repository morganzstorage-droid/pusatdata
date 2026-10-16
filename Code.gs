/**
 * ====================================================================
 * APLIKASI PENDATAAN ANAK YAYASAN KARYA MURNI RUTENG
 * SCRIPT GOOGLE APPS SCRIPT (Code.gs)
 * ====================================================================
 * Petunjuk Pemasangan:
 * 1. Buka Google Sheets baru di https://sheets.google.com
 * 2. Beri nama spreadsheet: "Database Pendataan Anak - Yayasan Karya Murni Ruteng"
 * 3. Klik menu: Ekstensi (Extensions) > Apps Script
 * 4. Hapus semua kode default di "Code.gs", lalu tempelkan seluruh kode di bawah ini.
 * 5. (Opsional) Jika ingin tampilan langsung di Google Apps Script:
 *    Klik tanda '+' di samping File > pilih HTML > beri nama "Index" > tempelkan kode Index.html.
 * 6. Klik tombol "Deploy" (Terapkan) berwarna biru di kanan atas > "New deployment" (Penerapan baru).
 * 7. Pilih tipe: "Web app" (Aplikasi Web).
 * 8. Pada kolom:
 *    - Description: "SIPENDA Karya Murni API v1"
 *    - Execute as: "Me" (Email Google Anda)
 *    - Who has access: "Anyone" (Siapa saja - penting agar aplikasi web bisa kirim data)
 * 9. Klik "Deploy", izinkan hak akses (Authorize access), pilih email Anda > Advanced > Go to (unsafe).
 * 10. Salin "Web app URL" (akhiran /exec) dan tempelkan ke Pengaturan Aplikasi!
 * ====================================================================
 */

var SHEET_NAME = "Data_Anak_YKM";

var HEADERS = [
  "ID",
  "Nama Lengkap",
  "Jenis Kelamin",
  "Tempat Lahir",
  "Tanggal Lahir",
  "Usia",
  "NIK",
  "Nomor KK",
  "Alamat Lengkap",
  "Desa/Kelurahan",
  "Kecamatan",
  "Kabupaten",
  "Jenjang Pendidikan",
  "Kelas",
  "Nama Sekolah",
  "Jenis Disabilitas",
  "Desil",
  "Layanan Kesehatan",
  "Status Layanan Kesehatan",
  "DTSEN",
  "Status KPD",
  "Nama Ayah",
  "NIK Ayah",
  "Pekerjaan Ayah",
  "Nama Ibu",
  "NIK Ibu",
  "Pekerjaan Ibu",
  "Domisili Anak Saat Ini",
  "Keterangan Tambahan",
  "Diinput Oleh",
  "Waktu Input"
];

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#1e1b4b");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    for (var i = 1; i <= HEADERS.length; i++) {
      sheet.setColumnWidth(i, 160);
    }
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(9, 250);
  }
  return sheet;
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "view";
  
  if (action === "getAll") {
    return ContentService.createTextOutput(JSON.stringify(getAllRecords()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "ping") {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Koneksi Google Sheet Yayasan Karya Murni Ruteng Berhasil!",
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    return HtmlService.createTemplateFromFile("Index")
      .evaluate()
      .setTitle("SIPENDA - Yayasan Karya Murni Ruteng")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "online",
      message: "Google Apps Script API Yayasan Karya Murni Ruteng Aktif.",
      totalData: getAllRecords().length
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var action = data.action || "SAVE";
    var sheet = getOrCreateSheet();
    
    if (action === "SAVE") {
      var child = data.record;
      var result = saveOrUpdateRecord(sheet, child);
      return createJsonResponse({ status: "success", message: "Data berhasil disimpan ke Google Sheet", result: result });
    }
    
    if (action === "DELETE") {
      var idToDelete = data.id;
      var deleted = deleteRecord(sheet, idToDelete);
      return createJsonResponse({ status: "success", message: "Data berhasil dihapus dari Google Sheet", deleted: deleted });
    }
    
    if (action === "BULK_SYNC") {
      var records = data.records || [];
      var syncResult = bulkSyncRecords(sheet, records);
      return createJsonResponse({ status: "success", message: "Sinkronisasi massal berhasil", totalSynced: syncResult });
    }
    
    return createJsonResponse({ status: "error", message: "Aksi tidak dikenali: " + action });
    
  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveOrUpdateRecord(sheet, item) {
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var targetId = item.id || ("YKM-" + new Date().getTime());
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(targetId)) {
      rowIndex = i + 1;
      break;
    }
  }
  
  var jenjangFull = item.jenjangPendidikan || "";
  if (jenjangFull === "Lainnya" && item.jenjangPendidikanLainnya) {
    jenjangFull = "Lainnya: " + item.jenjangPendidikanLainnya;
  }
  
  var rowValues = [
    targetId,
    item.nama || "",
    item.jenisKelamin || "",
    item.tempatLahir || "",
    item.tanggalLahir || "",
    item.usia || "",
    item.nik ? "'" + item.nik : "",
    item.kk ? "'" + item.kk : "",
    item.alamatLengkap || "",
    item.desaKelurahan || "",
    item.kecamatan || "",
    item.kabupaten || "",
    jenjangFull,
    item.kelas || "",
    item.namaSekolah || "",
    item.jenisDisabilitas || "",
    item.desil || "",
    item.layananKesehatan || "",
    item.statusLayananKesehatan || "",
    item.dtsen || "",
    item.statusKpd || "",
    item.namaAyah || "",
    item.nikAyah ? "'" + item.nikAyah : "",
    item.pekerjaanAyah || "",
    item.namaIbu || "",
    item.nikIbu ? "'" + item.nikIbu : "",
    item.pekerjaanIbu || "",
    item.domisiliAnakSaatIni || "",
    item.keteranganTambahan || "",
    item.createdBy || "Pengguna",
    new Date().toLocaleString("id-ID", { timeZone: "Asia/Makassar" })
  ];
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    return { type: "updated", id: targetId, row: rowIndex };
  } else {
    sheet.appendRow(rowValues);
    return { type: "inserted", id: targetId, row: sheet.getLastRow() };
  }
}

function deleteRecord(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function getAllRecords() {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    list.push({
      id: String(r[0] || ""),
      nama: String(r[1] || ""),
      jenisKelamin: String(r[2] || ""),
      tempatLahir: String(r[3] || ""),
      tanggalLahir: String(r[4] || ""),
      usia: String(r[5] || ""),
      nik: String(r[6] || "").replace(/^'/, ""),
      kk: String(r[7] || "").replace(/^'/, ""),
      alamatLengkap: String(r[8] || ""),
      desaKelurahan: String(r[9] || ""),
      kecamatan: String(r[10] || ""),
      kabupaten: String(r[11] || ""),
      jenjangPendidikan: String(r[12] || ""),
      kelas: String(r[13] || ""),
      namaSekolah: String(r[14] || ""),
      jenisDisabilitas: String(r[15] || ""),
      desil: String(r[16] || ""),
      layananKesehatan: String(r[17] || ""),
      statusLayananKesehatan: String(r[18] || ""),
      dtsen: String(r[19] || ""),
      statusKpd: String(r[20] || ""),
      namaAyah: String(r[21] || ""),
      nikAyah: String(r[22] || "").replace(/^'/, ""),
      pekerjaanAyah: String(r[23] || ""),
      namaIbu: String(r[24] || ""),
      nikIbu: String(r[25] || "").replace(/^'/, ""),
      pekerjaanIbu: String(r[26] || ""),
      domisiliAnakSaatIni: String(r[27] || ""),
      keteranganTambahan: String(r[28] || ""),
      createdBy: String(r[29] || ""),
      createdAt: String(r[30] || "")
    });
  }
  return list;
}

function bulkSyncRecords(sheet, records) {
  if (!records || records.length === 0) return 0;
  for (var i = 0; i < records.length; i++) {
    saveOrUpdateRecord(sheet, records[i]);
  }
  return records.length;
}
