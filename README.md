# Sistem Manajemen Kas Berbasis Google Sheets + Apps Script

Proyek ini merupakan sistem sederhana namun powerful untuk mengelola kas anggota menggunakan **Google Sheets** yang diperluas dengan **Google Apps Script**.

Dirancang untuk mempermudah pencatatan iuran, monitoring keuangan, dan transparansi data bagi seluruh anggota.

---

## Fitur Utama

### 1. Sidebar Interaktif

- Menambahkan anggota iuran dengan mudah
- Checklist pembayaran kas per bulan (tanpa input manual)
- Mendukung multi-periode iuran (contoh: 2026, 2027, dst)

---

### 2. Landing Page (Dashboard)

Halaman khusus untuk anggota melihat ringkasan keuangan:

- Saldo kas saat ini
- Pemasukan per bulan
- Pengeluaran per bulan
- Inventory (sementara menggunakan data dummy)
- Tabel iuran anggota

---

### 3. Automasi dengan Apps Script

- Mengurangi input manual
- Meminimalisir human error
- Integrasi langsung dengan spreadsheet

---

## Struktur Proyek

```
├── Code.gs              # Logic untuk Index.html Google Apps Script
├── Index.html           # Landing page / dashboard
├── SidebarInput.html    # Tampilan sidebar interaktif
├── SidebarCode.gs       # Logic untuk SidebarInput.html Google Apps Script
└── README.md            # Dokumentasi proyek
```

---

## Cara Menggunakan

1. Buka Google Sheets yang ingin digunakan
2. Masuk ke menu **Extensions > Apps Script**
3. Copy semua file dari repository ini ke project Apps Script Anda
4. Deploy atau jalankan script
5. Reload spreadsheet
6. Sidebar dan fitur dashboard akan muncul otomatis

---

## Kebutuhan

- Akun Google
- Akses ke Google Sheets
- Basic pemahaman Apps Script (opsional)

---

## Tujuan Proyek

- Mempermudah pengelolaan kas organisasi / komunitas
- Meningkatkan transparansi keuangan
- Mengurangi pekerjaan manual dalam pencatatan

---

## Preview

- Sidebar
- Dashboard
- Spreadsheet

---

## Lisensi

Gunakan lisensi sesuai kebutuhan (MIT / GPL / dll)

---

## Author

Dibuat untuk mempermudah manajemen kas berbasis spreadsheet dengan pendekatan sederhana namun efisien.

---

# googleappsscript-bukukas
