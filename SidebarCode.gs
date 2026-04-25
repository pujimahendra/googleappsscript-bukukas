/**
 * Menambahkan Menu "Aplikasi RT" di atas Google Sheets
 * Refresh / muat ulang spreadsheet Anda agar menu ini muncul.
 * nanti hasil dari menu yang kamu buat adalah dropdown
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Menu Admin")
    .addItem("Buka Input Iuran", "openIuranSidebar")
    .addToUi();
}

/**
 * Fungsi untuk membuka Sidebar
 */
function openIuranSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("SidebarInput")
    .setTitle("Input Iuran Anggota")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Mengambil daftar sheet yang mengandung kata "Iuran", apabila kas sudah lebih dari satu tahun
 * Digunakan untuk mengisi dropdown di Sidebar
 */
function getIuranSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = [];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    // Filter hanya sheet yang ada kata "Iuran" (tidak case-sensitive)
    if (name.toLowerCase().indexOf("iuran") > -1) {
      sheetNames.push(name);
    }
  }

  // Fallback: Jika tidak ada sheet bernama "Iuran", kembalikan semua nama sheet
  if (sheetNames.length === 0) {
    for (var i = 0; i < sheets.length; i++) {
      sheetNames.push(sheets[i].getName());
    }
  }
  return sheetNames;
}

/**
 * Mengambil daftar nama anggota dari Sheet yang dipilih secara dinamis
 */
function getDaftarNama(sheetName) {
  if (!sheetName) sheetName = "Iuran Anggota 2026"; // Default fallback
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan.");

  var data = sheet.getDataRange().getValues();
  var names = [];

  // Asumsi baris 1 kosong, baris 2 adalah header (No, Nama Anggota, dst)
  // Data anggota mulai di baris 3 (index array 2)
  for (var i = 2; i < data.length; i++) {
    var nama = data[i][1]; // Kolom B (indeks 1)
    if (nama && nama !== "") {
      names.push(nama);
    }
  }
  return names;
}

/**
 * Memproses penyimpanan iuran masal ke sheet yang dipilih
 */
function prosesIuran(selectedNames, bulan, nominal, sheetName) {
  if (!sheetName)
    throw new Error("Silakan pilih Tahun/Sheet Iuran terlebih dahulu!");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan.");

  var data = sheet.getDataRange().getValues();
  var headers = data[1]; // Header ada di baris 2 (indeks 1)

  // Cari indeks kolom untuk bulan yang dipilih
  var colIndex = headers.indexOf(bulan) + 1;
  if (colIndex === 0) {
    throw new Error("Kolom bulan '" + bulan + "' tidak ditemukan di header!");
  }

  // Lakukan perulangan dan isi nominal jika nama masuk dalam daftar terpilih
  for (var i = 2; i < data.length; i++) {
    var nama = data[i][1];
    if (selectedNames.indexOf(nama) > -1) {
      sheet.getRange(i + 1, colIndex).setValue(nominal);
    }
  }
  return (
    "Berhasil menyimpan iuran untuk " +
    selectedNames.length +
    " anggota di " +
    sheetName +
    "."
  );
}

/**
 * Menambahkan anggota baru ke baris paling bawah pada sheet yang dipilih
 */
function tambahAnggota(namaBaru, sheetName) {
  if (!sheetName)
    throw new Error("Silakan pilih Tahun/Sheet Iuran terlebih dahulu!");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan.");

  var data = sheet.getDataRange().getValues();
  var maxNo = 0;

  // Mencari Nomor Urut (No) terbesar untuk dilanjutkan
  for (var i = 2; i < data.length; i++) {
    var no = parseInt(data[i][0], 10);
    if (!isNaN(no) && no > maxNo) {
      maxNo = no;
    }
  }

  var newNo = maxNo + 1;
  var lastRow = sheet.getLastRow();

  // Menambahkan No dan Nama di baris baru
  sheet.getRange(lastRow + 1, 1).setValue(newNo);
  sheet.getRange(lastRow + 1, 2).setValue(namaBaru);

  // Kembalikan daftar nama yang baru diupdate agar UI Sidebar langsung sinkron
  return getDaftarNama(sheetName);
}

/**
 * Fungsi untuk menduplikasi sheet iuran (untuk ganti tahun dsb)
 */
function duplikasiSheetIuran(sourceSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAsli = ss.getSheetByName(sourceSheetName);

  if (!sheetAsli) {
    throw new Error("Sheet sumber tidak ditemukan!");
  }

  var sheetBaru = sheetAsli.copyTo(ss);
  var namaBaru = sourceSheetName + " (Copy)";

  // Cek jika nama sudah ada, hapus yang lama atau beri nama unik
  var cekSheet = ss.getSheetByName(namaBaru);
  if (cekSheet) {
    ss.deleteSheet(cekSheet);
  }

  sheetBaru.setName(namaBaru);
  sheetBaru.activate();

  return (
    "Sheet " + sourceSheetName + " berhasil diduplikasi menjadi: " + namaBaru
  );
}
