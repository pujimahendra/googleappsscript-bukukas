/**
 * Fungsi utama untuk menampilkan halaman HTML
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Laporan Keuangan Kas")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi untuk mengambil dan mengolah data dari Spreadsheet
 * Fungsi ini akan dipanggil oleh frontend (Index.html)
 */
function getDashboardData() {
  // Buka Spreadsheet yang sedang aktif (tempat script ini berada)
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ambil Data dari Sheet "Kas Utama"
  const kasSheet = ss.getSheetByName("Kas Utama");
  if (!kasSheet) return { error: "Sheet 'Kas Utama' tidak ditemukan" };

  const kasData = kasSheet.getDataRange().getValues();
  kasData.shift(); // Hapus baris pertama (Header)
  kasData.shift(); // Sesuai dengan format CSV Anda, sepertinya ada 2 baris awal sebelum data aktual (atau 1 header utama)
  // Perhatian: Jika header Anda hanya 1 baris, cukup gunakan 1 kali kasData.shift()

  let totalSaldo = 0;
  let pemasukanBulanIni = 0;
  let pengeluaranBulanIni = 0;

  // Dapatkan bulan dan tahun saat ini (Untuk filter dinamis)
  // Catatan: Di Javascript bulan dimulai dari 0 (0 = Januari, 3 = April)
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();

  let lastValidRowIndex = -1;

  // Loop data kas untuk mencari saldo terakhir dan total bulan ini
  for (let i = 0; i < kasData.length; i++) {
    const row = kasData[i];
    const tanggalRaw = row[1]; // Kolom B: Tanggal

    // Pastikan baris tersebut memiliki data tanggal (bukan baris kosong di bawah)
    if (tanggalRaw && tanggalRaw !== "") {
      lastValidRowIndex = i;

      const rowDate = new Date(tanggalRaw);
      const masuk = parseFloat(row[4]) || 0; // Kolom E: Masuk
      const keluar = parseFloat(row[5]) || 0; // Kolom F: Keluar

      // Jika transaksi terjadi pada bulan dan tahun ini
      if (
        rowDate.getMonth() === currentMonth &&
        rowDate.getFullYear() === currentYear
      ) {
        pemasukanBulanIni += masuk;
        pengeluaranBulanIni += keluar;
      }
    }
  }

  // Ambil saldo dari baris valid terakhir
  if (lastValidRowIndex !== -1) {
    totalSaldo = parseFloat(kasData[lastValidRowIndex][6]) || 0; // Kolom G: Saldo
  }

  // mencari tahu tahun sekarang
  var tahun = new Date().getFullYear();

  // 2. Ambil Data dari Sheet "Iuran Anggota 'tahun sekarang'"
  const iuranSheet = ss.getSheetByName("Iuran Anggota " + tahun);
  if (!iuranSheet)
    return { error: "Sheet 'Iuran Anggota " + tahun + "' tidak ditemukan" };

  const iuranDataRaw = iuranSheet.getDataRange().getValues();
  iuranDataRaw.shift(); // Hapus Header
  iuranDataRaw.shift(); // Tergantung struktur, sesuaikan shift ini jika nama anggota langsung di baris 2

  const iuranList = [];

  for (let i = 0; i < iuranDataRaw.length; i++) {
    const row = iuranDataRaw[i];
    const nama = row[1]; // Kolom B: Nama Anggota

    // Abaikan baris kosong
    if (nama && nama !== "") {
      iuranList.push({
        no: row[0],
        nama: nama,
        jan: row[2] || "",
        feb: row[3] || "",
        mar: row[4] || "",
        apr: row[5] || "",
        mei: row[6] || "",
        jun: row[7] || "",
        jul: row[8] || "",
        agu: row[9] || "",
        sep: row[10] || "",
        okt: row[11] || "",
        nov: row[12] || "",
        des: row[13] || "",
      });
    }
  }

  // Kembalikan data dalam format JSON/Object
  return {
    saldo: totalSaldo,
    pemasukan: pemasukanBulanIni,
    pengeluaran: pengeluaranBulanIni,
    inventaris: 0, // Dummy data untuk fitur tambahan
    iuran: iuranList,
    bulanAktif: getNamaBulan(currentMonth) + " " + currentYear,
  };
}

// Fungsi pembantu untuk nama bulan
function getNamaBulan(index) {
  const bulan = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  return bulan[index];
}
