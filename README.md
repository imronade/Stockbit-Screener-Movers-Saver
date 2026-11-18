# 🚀 Stockbit Screener & Movers Saver (All-in-One)

Ekstensi browser terpadu untuk **Google Chrome**, **Microsoft Edge**, **Brave Browser**, dan **Mozilla Firefox** yang menggabungkan fitur **Stockbit Screener Saver** dan **Stockbit Movers Saver**. Ekstensi ini memungkinkan Anda mengekspor data ke file Excel (`.xlsx`) dan menyinkronkan data secara otomatis ke dalam **Satu Google Spreadsheet Terpadu** dengan tampilan **Unified Calendar View (Side-by-Side)**.

---

## 📑 Daftar Isi
1. [Fitur Utama](#-fitur-utama)
2. [Panduan Instalasi Ekstensi Browser](#-panduan-instalasi-ekstensi-browser)
   * [Google Chrome](#-1-google-chrome)
   * [Microsoft Edge](#-2-microsoft-edge)
   * [Brave Browser](#-3-brave-browser)
   * [Mozilla Firefox](#-4-mozilla-firefox)
3. [Panduan Setup Google Spreadsheet (1-Klik)](#-panduan-setup-google-spreadsheet-1-klik)
4. [Menghubungkan Ekstensi ke Google Spreadsheet](#-menghubungkan-ekstensi-ke-google-spreadsheet)
5. [Panduan Penggunaan Lengkap](#-panduan-penggunaan-lengkap)
   * [1. Menggunakan Fitur Stockbit Screener](#1-menggunakan-fitur-stockbit-screener)
   * [2. Menggunakan Fitur Stockbit Movers](#2-menggunakan-fitur-stockbit-movers)
   * [3. Membaca Calendar View (Side-by-Side Center)](#3-membaca-calendar-view-side-by-side-center)
   * [4. Menggunakan Dashboard Rekap & Filter Tanggal](#4-menggunakan-dashboard-rekap--filter-tanggal)
   * [5. Mengatur Kategori Screener](#5-mengatur-kategori-screener-setting_kategori)
   * [6. Menggunakan Menu Toolbar Google Sheets](#6-menggunakan-menu-toolbar-google-sheets)
6. [Struktur Lembar Kerja Spreadsheet](#-struktur-lembar-kerja-spreadsheet)
7. [FAQ & Troubleshooting](#-faq--troubleshooting)

---

## ✨ Fitur Utama

1. **Kompatibilitas Multi-Browser (Manifest V3):**
   * Berjalan lancar di **Google Chrome**, **Microsoft Edge**, **Brave**, dan **Mozilla Firefox**.
   * Menggunakan *background service worker* yang aman dari proteksi *Content Security Policy (CSP)* Stockbit.
2. **Satu Google Spreadsheet Terpadu (All-in-One):**
   * **Side-by-Side Calendar View:**
     * **Sisi Kiri:** Kalender **Movers** dengan 7 checklist kategori (Top Gainer, Top Loser, Top Value, Top Volume, Top Frequency, Net Foreign Buy, Net Foreign Sell) untuk memantau irisan saham harian.
     * **Sisi Kanan:** Kalender **Screener** dengan filter dropdown kategori strategi.
     * **Center Filter:** Pemilihan **Tahun (E3)** dan **Bulan (H3)** berada di posisi tengah (*Center*), cukup ubah 1 kali untuk menyinkronkan kedua kalender sekaligus.
   * **Struktur Tab Rapi & Baku:**
     1. `panduan_penggunaan` *(halaman depan panduan)*
     2. `calendar_view` *(tampilan kalender komparatif kiri-kanan)*
     3. `movers_dashboard` *(rekap frekuensi & irisan kategori movers)*
     4. `screener_dashboard` *(rekap frekuensi & irisan screener)*
     5. `movers_records` *(database mentah movers)*
     6. `screener_records` *(database mentah screener dengan auto-detect kolom metrik baru)*
     7. `setting_kategori` *(pemetaan nama screener ke kategori strategi)*
3. **Ekspor Offline ke Excel (.xlsx):**
   * Download data Screener atau seluruh tab Movers langsung ke file `.xlsx` lokal tanpa memerlukan koneksi Google Sheets.
4. **1-Klik Setup Otomatis:**
   * Script menyusun seluruh layout, border, rumus, conditional formatting, warna tema, dan dropdown secara mandiri tanpa perlu upload template manual.

---

## 💻 Panduan Instalasi Ekstensi Browser

Folder ekstensi berada di komputer lokal Anda pada direktori:
```
C:\Personal\Project\Saham\Stockbit_Screener_Movers_Saver
```

---

### 🌐 1. Google Chrome

1. Buka browser **Google Chrome**.
2. Masuk ke halaman pengelolaan ekstensi dengan mengetikkan:
   ```
   chrome://extensions
   ```
   di address bar lalu tekan **Enter**.
3. Di pojok kanan atas, aktifkan tombol geser **Developer mode** (Mode pengembang).
4. Di pojok kiri atas, klik tombol **Load unpacked** (Muat yang belum dibongkar).
5. Pada jendela file explorer yang muncul, pilih folder:
   ```
   C:\Personal\Project\Saham\Stockbit_Screener_Movers_Saver
   ```
6. Klik **Select Folder** (Pilih Folder).
7. Ekstensi **Stockbit Screener & Movers Saver** akan langsung terpasang dan aktif.

---

### 🌊 2. Microsoft Edge

1. Buka browser **Microsoft Edge**.
2. Masuk ke halaman ekstensi:
   ```
   edge://extensions
   ```
3. Di panel menu sebelah kiri bawah, aktifkan toggle **Developer mode** (Mode pengembang).
4. Klik tombol **Load unpacked** (Muat ekstensi yang belum dibongkar) di bagian atas.
5. Pilih folder:
   ```
   C:\Personal\Project\Saham\Stockbit_Screener_Movers_Saver
   ```
6. Klik **Select Folder**. Ekstensi siap digunakan.

---

### 🦁 3. Brave Browser

1. Buka browser **Brave**.
2. Buka halaman ekstensi:
   ```
   brave://extensions
   ```
3. Aktifkan toggle **Developer mode** di pojok kanan atas.
4. Klik tombol **Load unpacked** di pojok kiri atas.
5. Arahkan dan pilih folder:
   ```
   C:\Personal\Project\Saham\Stockbit_Screener_Movers_Saver
   ```
6. Ekstensi langsung terpasang.

---

### 🦊 4. Mozilla Firefox

1. Buka browser **Mozilla Firefox**.
2. Ketik di address bar:
   ```
   about:debugging#/runtime/this-firefox
   ```
   lalu tekan **Enter**.
3. Pada bagian **Temporary Extensions** (Pengaya Sementara), klik tombol **Load Temporary Add-on...** (Muat Pengaya Sementara...).
4. Arahkan ke folder:
   ```
   C:\Personal\Project\Saham\Stockbit_Screener_Movers_Saver
   ```
5. Pilih file **`manifest.json`** lalu klik **Open**.
6. Ekstensi akan langsung aktif di browser Firefox Anda.

> [!TIP]
> Di Firefox versi standar, temporary add-on aktif selama sesi browser berjalan. Untuk penggunaan permanen di Firefox Developer Edition atau Firefox Nightly, Anda dapat mematikan verifikasi tanda tangan pengaya lewat `about:config` (`xpinstall.signatures.required` = `false`).

---

## 📊 Panduan Setup Google Spreadsheet (1-Klik)

Sebelum data dari browser dapat dikirim ke spreadsheet, Anda perlu menyiapkan Google Spreadsheet penerima data:

### Langkah 1: Buat Spreadsheet & Tempel Script
1. Buka [Google Sheets](https://sheets.new) untuk membuat spreadsheet baru.
2. Beri nama spreadsheet Anda (misal: `Stockbit Database Terpadu`).
3. Pada menu atas, buka **Extensions (Ekstensi)** > **Apps Script**.
4. Hapus seluruh baris kode default yang ada di dalam editor `Code.gs`.
5. Buka file `google_apps_script.js` yang ada di folder proyek, salin seluruh isinya (**Ctrl+A**, **Ctrl+C**), lalu tempel (**Ctrl+V**) ke dalam editor Apps Script.
6. Simpan project dengan menekan ikon disket **Save project** (Ctrl+S).

### Langkah 2: Jalankan Setup Template Awal
1. Di bilah menu atas Apps Script (di samping ikon Run/Debug), pastikan dropdown fungsi memilih **`setupTemplateViews`**.
2. Klik tombol **Run (Jalankan)**.
3. Saat pertama kali dijalankan, Google akan meminta izin keamanan (*Authorization Required*):
   * Klik **Review permissions** (Tinjau izin).
   * Pilih akun Google Anda.
   * Jika muncul peringatan *"Google hasn't verified this app"*, klik tautan **Advanced** (Lanjutan) di kiri bawah.
   * Klik **Go to Untitled project (unsafe)**.
   * Klik **Allow** (Izinkan).
4. Tunggu beberapa detik hingga proses selesai. Seluruh sheet (`panduan_penggunaan`, `calendar_view`, `movers_dashboard`, `screener_dashboard`, `movers_records`, `screener_records`, `setting_kategori`) akan langsung tersusun rapi beserta formula, border, dan warnanya!

### Langkah 3: Deploy sebagai Web App
1. Di pojok kanan atas Apps Script, klik tombol biru **Deploy** > **New deployment** (Penyebaran baru).
2. Di sebelah kiri tulisan *Select type*, klik ikon gerigi ⚙️ lalu pilih **Web app**.
3. Atur konfigurasi sebagai berikut:
   * **Description:** `Stockbit Saver Webhook`
   * **Execute as:** Pilih **`Me (email Anda)`**
   * **Who has access:** Wajib pilih **`Anyone`** *(Siapa saja)* agar ekstensi browser dapat mengirimkan data tanpa kendala autentikasi.
4. Klik tombol **Deploy**.
5. Salin (**Copy**) URL yang muncul di bawah tulisan **Web app URL** (URL berakhiran `/exec`). Simpan URL ini untuk dimasukkan ke ekstensi di browser.

---

## 🔗 Menghubungkan Ekstensi ke Google Spreadsheet

Pengaturan URL Web App ini cukup dilakukan **1 kali saja** di browser Anda, dan langsung berlaku untuk sinkronisasi Screener maupun Movers:

1. Buka situs [Stockbit](https://stockbit.com).
2. Anda bisa membuka menu pengaturan melalui salah satu cara berikut:
   * **Dari Panel Movers:** Klik ikon gerigi **⚙️** di samping tombol *Sync Movers*.
   * **Dari Halaman Screener:** Klik tombol bundar hijau di pojok kanan bawah, lalu pilih **Pengaturan Web App URL**.
3. Tempel (*paste*) URL Web App yang telah Anda salin tadi ke kolom yang tersedia:
   ```
   https://script.google.com/macros/s/.../exec
   ```
4. Klik **💾 Simpan URL**. Notifikasi hijau akan muncul menandakan konfigurasi berhasil disimpan.

---

## 🎯 Panduan Penggunaan Lengkap

### 1. Menggunakan Fitur Stockbit Screener

1. Buka halaman Screener di Stockbit (misal: `https://stockbit.com/#/screener`).
2. Jalankan screener saham pilihan Anda sampai tabel saham muncul.
3. Di pojok kanan bawah layar akan tampil tombol bulat hijau (**Floating Action Button**):
   * **📥 Download Screener (.xlsx):** Mengunduh tabel screener saat ini ke file Excel dengan nama file otomatis sesuai nama screener dan tanggal hari ini.
   * **☁️ Sync Screener ke Sheets:** Mengirimkan seluruh emiten hasil screener beserta metrik-metriknya ke spreadsheet:
     * Data tersimpan di sheet `screener_records`.
     * Jika Anda menambahkan indikator/kolom baru di Stockbit, script secara otomatis mendeteksi dan membuatkan kolom baru (*Dynamic Columns*) tanpa merusak data lama!
     * Kalender Screener dan Screener Dashboard otomatis ter-update seketika.

---

### 2. Menggunakan Fitur Stockbit Movers

1. Buka halaman Stockbit di mana panel **Movers** berada (biasanya di sisi kanan layar).
2. Di samping judul *Movers*, Anda akan melihat toolbar tombol baru:
   * **📊 Download Movers (.xlsx):**
     * Ekstensi akan otomatis berpindah dan mengumpulkan data dari seluruh tab Movers (*Top Gainer, Top Loser, Top Value, Top Volume, Top Frequency, Net Foreign Buy, Net Foreign Sell*).
     * Hasilnya otomatis dirangkum ke file Excel lokal dengan dua sheet: data mentah dan rekap irisan.
   * **☁️ Sync Movers ke Sheets:**
     * Ekstensi meng-crawl seluruh tab movers dan langsung mengirimkannya ke sheet `movers_records` di Google Spreadsheet Anda.
     * Kalender Movers dan Movers Dashboard otomatis ter-update seketika.

---

### 3. Membaca Calendar View (Side-by-Side Center)

Buka sheet **`calendar_view`** di Google Spreadsheet Anda. Tampilan ini didesain berdampingan untuk memudahkan komparasi pergerakan pasar (*Movers*) dengan hasil strategi (*Screener*):

```
+---------------------------------------------------------------------------------------------------+
|                           📅 CALENDAR VIEW - MOVERS & SCREENER                                    |
|                                                                                                   |
|               Tahun: [ 2026 ▼ ] (E3)       |       Bulan: [ November (11) ▼ ] (H3)                |
+--------------------------------------------+---+--------------------------------------------------+
| ⚡ MOVERS CALENDAR & FILTER IRISAN         | | | 🎯 SCREENER CALENDAR                             |
| [Top Gainer][Top Loser]... (A5:E5)         | | | Kategori Screener: (G5:H5)                       |
| [   ✓   ]   [       ]  ... (A6:E6)         | | | [ SEMUA KATEGORI ▼ ] (G6:H6)                     |
| [Net Foreign Buy] [Net Foreign Sell] (A7:B7)| | | 🟢 Emiten hijau = muncul ≥ 2 kali dalam bulan    |
| [   ✓   ]         [       ]          (A8:B8)| | | ⭐ Bintang = irisan screener (⭐⭐ = 2, ⭐⭐⭐ = 3) |
+--------------------------------------------+ | +--------------------------------------------------+
| Senin   | Selasa  | Rabu    | Kamis  | Jumat| | | Senin   | Selasa  | Rabu    | Kamis  | Jumat     |
| (Kol A) | (Kol B) | (Kol C) | (Kol D)|(Kol E)| | (Kol G) | (Kol H) | (Kol I) | (Kol J)| (Kol K)   |
+---------+---------+---------+--------+------+---+---------+---------+---------+--------+----------+
| 03/11   | 04/11   | 05/11   | 06/11  | 07/11| | | 03/11   | 04/11   | 05/11   | 06/11  | 07/11    |
| BREN    | BBCA    | ASII    | TLKM   | ...  | | | BRMS    | BBCA    | ...     | ...    | ...      |
+---------+---------+---------+--------+------+---+---------+---------+---------+--------+----------+
```

#### Cara Penggunaan Calendar View:
1. **Pengaturan Tahun (`E3`) & Bulan (`H3`) di Posisi Center:**
   * Cukup ubah Tahun di `E3` atau Bulan di `H3`, maka **kedua sisi kalender (Movers dan Screener) langsung ter-update serentak** ke bulan dan tahun perdagangan tersebut.
2. **Sisi Kiri (Movers - 7 Checklist Kategori & Filter Irisan):**
   * **Baris 5 & 6:** Checklist `Top Gainer`, `Top Loser`, `Top Value`, `Top Volume`, `Top Frequency`.
   * **Baris 7 & 8:** Checklist `Net Foreign Buy` dan `Net Foreign Sell`.
   * **Fitur Irisan Penuh (*Full Intersection*):** Jika Anda mencentang *Top Gainer* dan *Net Foreign Buy*, kalender hanya akan menampilkan saham yang berhasil masuk ke **kedua** kategori tersebut (saham yang naik dan diborong asing pada hari yang sama).
   * Anda bebas mengombinasikan kategori manapun (*misal: Top Value + Net Foreign Buy*).
3. **Sisi Kanan (Screener):**
   * Di sel `G6` terdapat dropdown kategori strategi (misal: `SEMUA KATEGORI`, `MOMENTUM`, `BREAKOUT`, `VALUE`, `SWING`).
   * Pilih kategori untuk menyaring emiten yang lolos pada hari tersebut.
4. **Analisis Silang (*Cross-Check*):**
   * Karena baris tanggalnya sejajar lurus dari kiri ke kanan, Anda bisa melihat dalam sekejap pada baris hari yang sama (misal hari Selasa, 4 November):
     > *"Saham apa yang lolos Screener Momentum hari ini sekaligus masuk ke jajaran Top Movers / Net Foreign Buy?"* (Contoh: jika `BBCA` muncul di kedua sisi).
5. **Simbol Bintang Irisan Screener (Multi-Strategy Hit) & Stabilo Hijau:**
   * **Simbol Bintang (⭐):** Jika suatu emiten lolos di $\ge 2$ screener berbeda pada tanggal yang sama, otomatis diberi simbol bintang sebanyak jumlah screener yang ditembus:
     * `BBCA ⭐⭐` = Lolos di 2 strategi screener sekaligus hari itu.
     * `BBCA ⭐⭐⭐` = Lolos di 3 strategi screener sekaligus hari itu, dst.
   * **Stabilo Hijau & Badge Tanggal (🟢):** Saham yang konsisten muncul di beberapa tanggal berbeda dalam bulan yang sama otomatis ditandai:
     * Diberi **warna teks hijau tebal** dengan badge jumlah kemunculan: `MSJA ⭐⭐ (2x) 🟢`.
     * Kotak sel tanggal otomatis diberi **stabilo background hijau lembut** (`#f0fdf4`) sehingga hari-hari bursa yang memiliki saham konsisten langsung terlihat mencolok.
     * Di `movers_dashboard` & `screener_dashboard`, baris emiten yang muncul $\ge 2$ hari otomatis diberi **stabilo hijau** (`#dcfce7`).

---

### 4. Menggunakan Dashboard Rekap & Filter Tanggal

Terdapat dua dashboard analisis:
* **`movers_dashboard`:** Menyajikan daftar emiten, total hari kemunculan, total irisan kategori, rincian kategori yang pernah dimasuki, dan harga penutupan terakhir.
* **`screener_dashboard`:** Menyajikan **Tabel Terpisah per Kategori Strategi** (misal: *MOMENTUM*, *BREAKOUT*, *VALUE*, dll.) lengkap dengan banner warna khas, daftar emiten, total hari lolos filter, frekuensi screener, dan nama screener terkait.

> [!TIP]
> **Pilihan Semua Data (All Records):**
> Jika sel tanggal (**Dari Tanggal** atau **Sampai Tanggal**) pada baris 3 **dikosongkan / dihapus**, sistem secara otomatis **memilih dan menampilkan seluruh data histori** tanpa pembatasan rentang waktu.

---

### 5. Mengatur Kategori Screener (`setting_kategori`)

Di sheet **`setting_kategori`**, Anda dapat mengelompokkan berbagai nama screener Stockbit Anda ke dalam kelompok strategi:
* **Kolom A:** Nama Screener di Stockbit (harus sama persis dengan judul screener di web).
* **Kolom B:** Kategori Strategi (misal: `MOMENTUM`, `BREAKOUT`, `SWING`, `VALUE`, `DIVIDEND`).

Daftar kategori di Kolom B ini otomatis menjadi pilihan dropdown di `calendar_view` (sel `G6`) serta mengelompokkan tabel-tabel di `screener_dashboard`.

---

### 6. Menggunakan Menu Toolbar Google Sheets

Di toolbar atas Google Spreadsheet, Anda memiliki menu khusus:
```
🚀 Stockbit Tools
├── ✨ Setup Awal / Reset
├── 🔄 Update Script
└── 🔄 Refresh Tampilan
```

* **`✨ Setup Awal / Reset`:**
  * Membangun ulang seluruh template, layout kanan-kiri, tabel per kategori screener, styling border, warna tema, validasi dropdown, dan merapikan urutan tab sheet dari awal tanpa menghapus histori records.
* **`🔄 Update Script`:**
  * Memeriksa versi script saat ini dengan repository GitHub resmi secara live:
    * Jika sudah sama, menampilkan status **`✅ Script Up to Date`** sehingga Anda tidak perlu repot update.
    * Jika ada versi baru, menampilkan badge **`⚡ Update Tersedia`**, ringkasan versi baru, dan 1 tombol **"📋 Salin Kode Terbaru"** ke clipboard beserta panduan deploy 2 langkah.
* **`🔄 Refresh Tampilan`:**
  * **1-Klik untuk Semua:** Otomatis memperbarui data Kalender Movers, Kalender Screener, Movers Dashboard, Screener Dashboard, dan memastikan urutan sheet tetap rapi.

---

## 📁 Struktur Lembar Kerja Spreadsheet

| No | Nama Sheet | Keterangan |
|:--:|:-----------|:-----------|
| 1 | `panduan_penggunaan` | Panduan interaktif cara penggunaan langsung di spreadsheet. |
| 2 | `calendar_view` | Tampilan kalender komparatif side-by-side (Movers kiri, Screener kanan, filter Center). |
| 3 | `movers_dashboard` | Rekapitulasi statistik frekuensi & irisan kategori Movers. |
| 4 | `screener_dashboard` | Rekapitulasi statistik frekuensi & irisan strategi Screener. |
| 5 | `movers_records` | Database histori data mentah seluruh tab Movers. |
| 6 | `screener_records` | Database histori data mentah Screener (kolom dinamis). |
| 7 | `setting_kategori` | Tabel pemetaan nama screener ke kategori strategi. |

---

## ❓ FAQ & Troubleshooting

### 1. Muncul notifikasi "Ekstensi baru direload. Silakan tekan F5" saat klik Sync
* **Penyebab:** Ekstensi baru saja di-refresh atau di-update di halaman `chrome://extensions` atau `about:debugging`.
* **Solusi:** Cukup tekan tombol **F5** (Refresh) pada halaman Stockbit Anda, lalu klik tombol Sync kembali.

### 2. Status HTTP 401 / 403 saat Sync ke Google Sheets
* **Penyebab:** Pada pengaturan deployment Google Apps Script, akses belum diatur ke publik.
* **Solusi:** Buka Apps Script > klik **Deploy** > **Manage deployments** > klik ikon pensil (Edit) > ubah opsi **Who has access** menjadi **`Anyone`** (Siapa saja) > klik **Deploy**.

### 3. Saya mengedit script di Apps Script, apakah URL Web App berubah?
* Agar URL tidak berubah, setiap kali melakukan edit script, klik **Deploy** > **Manage deployments** > klik ikon pensil (Edit) > pada dropdown *Version*, pilih **New version** (Versi baru) > klik **Deploy**. URL Web App Anda akan tetap sama.

### 4. Library Excel belum selesai dimuat saat klik Download (.xlsx)
* Tunggu sekitar 2-3 detik setelah halaman Stockbit terbuka sempurna sebelum menekan tombol download agar file `xlsx.full.min.js` selesai diinisialisasi oleh browser.

### 5. Tampilan kalender tampak tertimpa atau balik ke versi lama setelah Sync
* **Penyebab:** Web App URL masih mengeksekusi versi snapshot deployment lama di server Google karena setelah script diperbarui, deployment belum di-update ke *New version*.
* **Solusi:** Di Google Apps Script, klik **Deploy** > **Manage deployments** > Edit (ikon pensil) > pada dropdown *Version* pilih **New version** (Versi baru) > klik **Deploy**. Setelah itu, buka spreadsheet dan klik **🚀 Stockbit Tools** > **✨ Setup Awal / Reset**. Layout akan terkunci rapi dengan 7 kategori.
