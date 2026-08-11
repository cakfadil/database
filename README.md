# 💰 Aplikasi Pengelolaan Pinjaman Mingguan Nasabah

Aplikasi web pengelolaan pinjaman mingguan nasabah berbasis **Single Page Application (SPA)** yang dideploy via **GitHub Pages** dengan backend API menggunakan **Google Apps Script** dan database **Google Sheets**.

---

## 📁 Struktur Berkas

```
pinjaman-mingguan/
├── Code.gs         # Kode Backend Google Apps Script (Database & API Logic)
├── index.html      # Frontend Web Application (HTML5, Glassmorphism CSS, & JS Fetch API)
└── README.md       # Dokumentasi & Panduan Setup Projek
```

---

## 🛠️ 1. Setup Backend Google Apps Script & Google Sheets

1. **Buat Google Sheets Baru**:
   - Akses [sheets.new](https://sheets.new) di browser Anda.
   - Beri nama Spreadsheet, contoh: `Database Pinjaman Mingguan`.

2. **Buka Editor Apps Script**:
   - Klik menu **Extensions > Apps Script** (Ekstensi > Apps Script).

3. **Salin Kode Backend**:
   - Hapus semua kode default pada file `Code.gs`.
   - Buka file [`Code.gs`](Code.gs) dari proyek VS Code ini, lalu salin dan tempel seluruh isinya ke editor Apps Script.

4. **Inisialisasi Database (Dummy Data)**:
   - Di bagian atas editor Apps Script, pilih fungsi `setupDatabase`.
   - Klik tombol **Run (Jalankan)** dan berikan izin akses (authorization).
   - Periksa kembali Spreadsheet Anda. 3 Sheet (`Setting`, `Pinjaman`, `Transaksi`) telah otomatis dibuat beserta data sampelnya.

5. **Deploy Web App API**:
   - Klik tombol **Deploy > New deployment** (Penerapan > Penerapan baru).
   - Pilih jenis: **Web app**.
   - **Description**: `v1 API Backend`.
   - **Execute as**: `Me` (Email Anda).
   - **Who has access**: **`Anyone`** *(Penting: Wajib pilih "Anyone/Siapa saja" agar frontend GitHub Pages bisa mengakses API)*.
   - Klik **Deploy**.
   - **Salin Web App URL** yang dihasilkan (contoh format: `https://script.google.com/macros/s/AKfycbx.../exec`).

---

## 💻 2. Pengembangan & Uji Coba di VS Code

1. Buka folder ini di **VS Code**:
   ```bash
   code .
   ```
2. Jalankan lokal server untuk pengujian (bisa menggunakan ekstensi VS Code **Live Server** atau `npx serve .`).
3. Buka halaman aplikasi di browser.
4. Klik tombol **⚙️ Pengaturan URL Backend API** di halaman login.
5. Tempelkan **Web App URL** yang didapatkan dari Google Apps Script, lalu klik **Simpan URL**.

---

## 🚀 3. Deployment ke GitHub Pages

1. **Inisialisasi Git & Push ke GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - Aplikasi Pinjaman Mingguan"
   git branch -M main
   git remote add origin https://github.com/USERNAME/NAMA-REPO.git
   git push -u origin main
   ```

2. **Aktifkan GitHub Pages**:
   - Buka repositori Anda di GitHub.
   - Buka tab **Settings** > menu **Pages** di bilah samping kiri.
   - Pada bagian **Build and deployment > Source**, pilih **Deploy from a branch**.
   - Pada **Branch**, pilih `main` dan folder `/ (root)`, lalu klik **Save**.
   - Tunggu 1-2 menit hingga link GitHub Pages Anda aktif (contoh: `https://USERNAME.github.io/NAMA-REPO/`).

3. **Gunakan Aplikasi di GitHub Pages**:
   - Buka situs GitHub Pages Anda.
   - Masukkan URL Apps Script API via modal **⚙️ Pengaturan URL Backend API** (URL disimpan dengan aman di `localStorage` browser Anda).

---

## 🔑 Kredensial Uji Coba Default

| Username | Password | Role | Akses Khusus |
|---|---|---|---|
| `admin` | `admin123` | Administrator | Akses penuh (Batch Payment, Renewal, User Management) |
| `budi` | `pass123` | Nasabah | Resot22 - Transaksi & Tabungan Terbuka |
| `siti` | `pass123` | Nasabah | Resot22 - Transaksi & Tabungan Terbuka |
| `andi` | `pass123` | Nasabah | Resot23 - Transaksi & Tabungan Disembunyikan Admin |
