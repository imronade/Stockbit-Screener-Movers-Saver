/**
 * =========================================================================================
 * STOCKBIT SCREENER & MOVERS SAVER - ALL-IN-ONE GOOGLE APPS SCRIPT
 * =========================================================================================
 * Versi: 1.3.7 (Enforce Timezone Asia/Jakarta WIB & Smart Market Date Automation)
 * 
 * FITUR UTAMA:
 * 1. Side-by-Side Calendar View (Senin - Jumat):
 *    - Tabel KIRI  : Movers Calendar (Didukung 7 Checklist Kategori & Filter Irisan)
 *    - Tabel KANAN : Screener Calendar (Didukung Dropdown Kategori Strategi)
 *    - Pengaturan Tahun (E3) & Bulan (H3) CENTER TERPADU: Cukup ubah 1 kali, kedua kalender sinkron!
 *    - Sejajar Tanggal per Minggu: Langsung lihat perbandingan Movers vs Screener per hari!
 * 2. Multi-Record Storage:
 *    - 'movers_records'   : Menyimpan histori data mentah Movers (Top Gainer, Loser, Foreign Flow, dll.)
 *    - 'screener_records' : Menyimpan histori data Screener dengan deteksi kolom metrik dinamis.
 * 3. Dedicated Dashboards:
 *    - 'movers_dashboard'   : Rekap statistik frekuensi & irisan Movers dalam rentang tanggal.
 *    - 'screener_dashboard' : Rekap statistik frekuensi & irisan Screener dalam rentang tanggal.
 * 4. Webhook Receiver Cerdas (doPost):
 *    - Otomatis mendeteksi payload apakah dari Screener atau Movers.
 *    - Memperbarui data mentah dan me-refresh tampilan calendar & dashboard secara instan.
 * 5. Event onEdit Otomatis:
 *    - Ubah Tahun/Bulan -> Refresh kedua kalender.
 *    - Ubah Checkbox Movers -> Refresh kalender Movers.
 *    - Ubah Kategori Screener -> Refresh kalender Screener.
 * =========================================================================================
 */

// =========================================================================================
// KONFIGURASI VERSI SCRIPT
// =========================================================================================
const SCRIPT_VERSION = "1.3.7";

/**
 * Menu Kustom di Google Sheets saat file dibuka
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Stockbit Tools")
    .addItem("✨ Setup Awal / Reset", "setupTemplateViews")
    .addSeparator()
    .addItem("🛠️ Koreksi Tanggal Records", "showFixRecordsDateModal")
    .addSeparator()
    .addItem("🔄 Update Script", "showUpdateScriptModal")
    .addSeparator()
    .addItem("🔄 Refresh Tampilan", "updateAllViews")
    .addToUi();
}

/**
 * Otomatis mendeteksi perubahan cell (onEdit)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // 1. Edit di sheet 'calendar_view'
    if (sheetName === "calendar_view") {
      // Edit Tahun (E3) atau Bulan (H3) di posisi Center -> Refresh KEDUA kalender
      if (row === 3 && (col === 5 || col === 8)) {
        renderUnifiedCalendarView();
        return;
      }
      // Edit Kategori Screener (G6 atau H6) -> Refresh Kalender Screener
      if (row === 6 && (col === 7 || col === 8)) {
        renderScreenerCalendarPart();
        return;
      }
      // Edit Checkbox Movers:
      // Baris 6 (Kolom A s/d E: Top Gainer, Loser, Value, Volume, Freq)
      // Baris 8 (Kolom A s/d B: Net Foreign Buy, Net Foreign Sell)
      if ((row === 6 && col >= 1 && col <= 5) || (row === 8 && col >= 1 && col <= 2)) {
        renderMoversCalendarPart();
        return;
      }
    }

    // 2. Edit filter tanggal di 'movers_dashboard' (B3 atau E3)
    if (sheetName === "movers_dashboard" && row === 3 && (col === 2 || col === 5)) {
      renderMoversDashboard();
      return;
    }

    // 3. Edit filter tanggal di 'screener_dashboard' (C3 atau E3)
    if (sheetName === "screener_dashboard" && row === 3 && (col === 3 || col === 5)) {
      renderScreenerDashboard();
      return;
    }

    // 4. Edit di sheet 'setting_kategori' -> Perbarui dropdown kategori Screener
    if (sheetName === "setting_kategori") {
      updateCategoryDropdown();
      return;
    }
  } catch (err) {
    Logger.log("Error onEdit: " + err.toString());
  }
}

/**
 * Webhook Receiver (doPost) untuk Chrome/Firefox Extension
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Tidak ada data post yang diterima." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureWIBTimeZone(ss);

    // -----------------------------------------------------------------------
    // A. JALUR MOVERS
    // -----------------------------------------------------------------------
    if (payload.type === "movers" || payload.action === "appendMovers" || payload.records) {
      const records = payload.records || [];
      if (records.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({ status: "warning", message: "Data records Movers kosong." }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      let wsMovers = getSheetCaseInsensitive(ss, "movers_records");
      if (!wsMovers) {
        wsMovers = ss.insertSheet("movers_records");
        const moversHeaders = [["Tanggal", "Kategori", "Emiten", "Price", "IDX Points", "% IDX Mov", "Value", "Volume", "Freq", "Net Foreign"]];
        wsMovers.getRange("A1:J1").setValues(moversHeaders)
          .setFontWeight("bold")
          .setBackground("#1e40af")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("center");
        wsMovers.setFrozenRows(1);
      }

      const rows = records.map(r => [
        r.tanggal, r.kategori, r.symbol, r.price, r.idxPoints, r.idxMov, r.value, r.volume, r.freq, r.netForeign
      ]);

      const lastRow = wsMovers.getLastRow();
      wsMovers.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);

      // Refresh tampilan Movers
      renderMoversCalendarPart();
      renderMoversDashboard();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        count: rows.length,
        message: `Berhasil menyimpan ${rows.length} baris Movers! Calendar & Dashboard ter-update.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // -----------------------------------------------------------------------
    // B. JALUR SCREENER
    // -----------------------------------------------------------------------
    if (payload.type === "screener" || payload.action === "appendScreener" || payload.screenerName || payload.rows) {
      const screenerName = payload.screenerName || "Stockbit Screener";

      // Jika tanggal tidak dikirim dari payload, hitung tanggal bursa terakhir berbasis jam buka BEI (09.00 WIB)
      let dateStr = payload.date;
      if (!dateStr) {
        const now = new Date();
        const hourWIB = parseInt(Utilities.formatDate(now, "Asia/Jakarta", "HH"), 10);
        const dayWIB = parseInt(Utilities.formatDate(now, "Asia/Jakarta", "u"), 10); // 1 = Senin, ..., 7 = Minggu

        let target = new Date(now.getTime());
        if (dayWIB === 6) {
          // Sabtu -> mundur 1 hari ke Jumat
          target.setDate(target.getDate() - 1);
        } else if (dayWIB === 7) {
          // Minggu -> mundur 2 hari ke Jumat
          target.setDate(target.getDate() - 2);
        } else if (hourWIB < 9) {
          // Senin s/d Jumat sebelum jam 09.00 WIB (bursa belum buka):
          if (dayWIB === 1) {
            // Senin sebelum 09.00 WIB -> mundur 3 hari ke Jumat
            target.setDate(target.getDate() - 3);
          } else {
            // Selasa s/d Jumat sebelum 09.00 WIB -> mundur 1 hari (kemarin)
            target.setDate(target.getDate() - 1);
          }
        }
        dateStr = Utilities.formatDate(target, "Asia/Jakarta", "yyyy-MM-dd");
      }
      const dayStr = getIndonesianDay(dateStr);
      const rows = payload.rows || [];

      if (rows.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({ status: "warning", message: "Data baris Screener kosong." }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      let wsScreener = getSheetCaseInsensitive(ss, "screener_records");
      const settingSheet = getSheetCaseInsensitive(ss, "setting_kategori");
      if (!wsScreener) {
        wsScreener = ss.insertSheet("screener_records");
        const scrHeaders = [["No", "Tanggal", "Hari", "Nama Screener", "Kategori", "Emiten", "Harga Penutupan"]];
        wsScreener.getRange("A1:G1").setValues(scrHeaders)
          .setFontWeight("bold")
          .setBackground("#00875A")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("center");
        wsScreener.setFrozenRows(1);
      }

      const categoryMap = getCategoryMapping(settingSheet);
      const category = categoryMap[screenerName] || "Lainnya";

      // Baca header kolom terkini di sheet screener_records
      let currentHeaders = [];
      const lastCol = wsScreener.getLastColumn();
      if (lastCol > 0) {
        currentHeaders = wsScreener.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      }
      const baseHeaders = ["No", "Tanggal", "Hari", "Nama Screener", "Kategori", "Emiten", "Harga Penutupan"];
      if (currentHeaders.length === 0) {
        currentHeaders = [...baseHeaders];
      }

      // Deteksi kolom metrik baru yang belum ada
      let incomingMetrics = new Set();
      rows.forEach(r => {
        if (r.metrics) {
          Object.keys(r.metrics).forEach(m => incomingMetrics.add(m));
        }
      });

      let newColsToAdd = [];
      incomingMetrics.forEach(colName => {
        if (!currentHeaders.includes(colName) && !["symbol", "emiten", "saham", "price", "harga"].includes(colName.toLowerCase())) {
          newColsToAdd.push(colName);
          currentHeaders.push(colName);
        }
      });

      if (newColsToAdd.length > 0) {
        const startNewCol = currentHeaders.length - newColsToAdd.length + 1;
        wsScreener.getRange(1, startNewCol, 1, newColsToAdd.length)
          .setValues([newColsToAdd])
          .setBackground("#00875A")
          .setFontColor("#FFFFFF")
          .setFontWeight("bold")
          .setHorizontalAlignment("center");
      }

      // Susun baris Screener
      let lastRow = wsScreener.getLastRow();
      let nextNo = lastRow > 1 ? (Number(wsScreener.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;

      let appendRows = [];
      rows.forEach(r => {
        let rowData = new Array(currentHeaders.length).fill("");
        rowData[0] = nextNo++;
        rowData[1] = dateStr;
        rowData[2] = dayStr;
        rowData[3] = screenerName;
        rowData[4] = category;
        rowData[5] = r.emiten;
        rowData[6] = r.price;

        if (r.metrics) {
          Object.keys(r.metrics).forEach(k => {
            let colIdx = currentHeaders.indexOf(k);
            if (colIdx !== -1) rowData[colIdx] = r.metrics[k];
          });
        }
        appendRows.push(rowData);
      });

      if (appendRows.length > 0) {
        wsScreener.getRange(lastRow + 1, 1, appendRows.length, currentHeaders.length).setValues(appendRows);
      }

      // Refresh tampilan Screener
      renderScreenerCalendarPart();
      renderScreenerDashboard();

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        count: appendRows.length,
        message: `Berhasil menyimpan ${appendRows.length} emiten dari Screener "${screenerName}"!`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "warning", message: "Payload tidak dikenali sebagai Screener maupun Movers." }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Stockbit Screener & Movers Web App Online! Status: OK.")
    .setMimeType(ContentService.MimeType.TEXT);
}

// =========================================================================================
// HELPER FUNCTIONS
// =========================================================================================
function getSheetCaseInsensitive(ss, targetName) {
  if (!ss || !targetName) return null;
  const direct = ss.getSheetByName(targetName);
  if (direct) return direct;
  const targetLower = targetName.toLowerCase().trim();
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName().toLowerCase().trim() === targetLower) {
      return s;
    }
  }
  return null;
}

function safeToast(msg, title = "Info", timeout = 3) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, timeout);
  } catch (_) {}
}

/**
 * Memastikan Spreadsheet selalu terkunci pada zona waktu resmi Asia/Jakarta (WIB / UTC+7)
 */
function ensureWIBTimeZone(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getSpreadsheetTimeZone() !== "Asia/Jakarta") {
      ss.setSpreadsheetTimeZone("Asia/Jakarta");
    }
  } catch (_) {}
}

function parseDateToYMD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    // Menambahkan 12 jam (tengah hari) dan format paten menggunakan zona waktu Asia/Jakarta (WIB / UTC+7)
    const safeDate = new Date(val.getTime() + (12 * 3600 * 1000));
    return Utilities.formatDate(safeDate, "Asia/Jakarta", "yyyy-MM-dd");
  }
  if (typeof val === "number") {
    const epoch = new Date(1899, 11, 30, 12, 0, 0);
    const d = new Date(epoch.getTime() + val * 86400000);
    return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.substring(0, 10);
  }
  const dmy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (dmy) {
    const d = String(dmy[1]).padStart(2, "0");
    const m = String(dmy[2]).padStart(2, "0");
    const y = dmy[3];
    return `${y}-${m}-${d}`;
  }
  return "";
}

/**
 * Format Date object ke YYYY-MM-DD murni berbasis getFullYear/getMonth/getDate lokal
 * sehingga 100% bebas dari risiko geseran offset zona waktu antar-server.
 */
function formatDateYMD(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function getIndonesianDay(dateVal) {
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  if (!dateVal) {
    const dayNum = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "u"), 10);
    return days[dayNum % 7];
  }
  if (dateVal instanceof Date) {
    const dayNum = parseInt(Utilities.formatDate(dateVal, "Asia/Jakarta", "u"), 10);
    return days[dayNum % 7];
  }
  const ymd = parseDateToYMD(dateVal);
  if (ymd) {
    const parts = ymd.split("-");
    const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
    return days[d.getDay()];
  }
  return days[new Date().getDay()];
}

function getCategoryMapping(settingSheet) {
  const map = {};
  if (!settingSheet) return map;
  const lastRow = settingSheet.getLastRow();
  if (lastRow > 1) {
    const data = settingSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(r => {
      const scrName = String(r[0] || "").trim();
      const kat = String(r[1] || "").trim();
      if (scrName && kat) map[scrName] = kat;
    });
  }
  return map;
}

function getMonthNamesList() {
  return [
    "Januari (1)", "Februari (2)", "Maret (3)", "April (4)",
    "Mei (5)", "Juni (6)", "Juli (7)", "Agustus (8)",
    "September (9)", "Oktober (10)", "November (11)", "Desember (12)"
  ];
}

function parseMonthIndex(monthStr) {
  if (!monthStr) {
    return parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "M"), 10);
  }
  const s = String(monthStr).trim().toLowerCase();

  // 1. Cek angka dalam kurung, misal "November (11)"
  const match = s.match(/\((\d+)\)/);
  if (match) return parseInt(match[1], 10);

  // 2. Cek jika hanya angka, misal "11"
  const num = parseInt(s, 10);
  if (!isNaN(num) && num >= 1 && num <= 12) return num;

  // 3. Cek nama bulan Indonesia / Inggris
  const monthMap = {
    "januari": 1, "january": 1, "jan": 1,
    "februari": 2, "february": 2, "feb": 2,
    "maret": 3, "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "mei": 5, "may": 5,
    "juni": 6, "june": 6, "jun": 6,
    "juli": 7, "july": 7, "jul": 7,
    "agustus": 8, "august": 8, "agt": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "october": 10, "okt": 10, "oct": 10,
    "november": 11, "nov": 11,
    "desember": 12, "december": 12, "des": 12, "dec": 12
  };
  for (const [name, idx] of Object.entries(monthMap)) {
    if (s.startsWith(name)) return idx;
  }

  return parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "M"), 10);
}

// =========================================================================================
// 1. SETUP TEMPLATE SEMUA SHEET (ZERO-FILE SETUP)
// =========================================================================================
function setupTemplateViews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWIBTimeZone(ss);

  // 1. Sheet: panduan_penggunaan
  setupGuideSheetStructure();

  // 2. Sheet: calendar_view (SIDE-BY-SIDE UNIFIED)
  setupCalendarSheetStructure();

  // 3. Sheet: movers_dashboard
  setupMoversDashboardStructure();

  // 4. Sheet: screener_dashboard
  setupScreenerDashboardStructure();

  // 5. Sheet: movers_records
  let wsMoversRec = getSheetCaseInsensitive(ss, "movers_records");
  if (!wsMoversRec) wsMoversRec = ss.insertSheet("movers_records");
  if (wsMoversRec.getLastRow() === 0) {
    const moversHeaders = [["Tanggal", "Kategori", "Emiten", "Price", "IDX Points", "% IDX Mov", "Value", "Volume", "Freq", "Net Foreign"]];
    wsMoversRec.getRange("A1:J1").setValues(moversHeaders)
      .setFontWeight("bold")
      .setBackground("#1e40af")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    wsMoversRec.setFrozenRows(1);
  }

  // 6. Sheet: screener_records
  let wsScreenerRec = getSheetCaseInsensitive(ss, "screener_records");
  if (!wsScreenerRec) wsScreenerRec = ss.insertSheet("screener_records");
  if (wsScreenerRec.getLastRow() === 0) {
    const scrHeaders = [["No", "Tanggal", "Hari", "Nama Screener", "Kategori", "Emiten", "Harga Penutupan"]];
    wsScreenerRec.getRange("A1:G1").setValues(scrHeaders)
      .setFontWeight("bold")
      .setBackground("#00875A")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    wsScreenerRec.setFrozenRows(1);
  }

  // 7. Sheet: setting_kategori
  let wsKat = getSheetCaseInsensitive(ss, "setting_kategori");
  if (!wsKat) wsKat = ss.insertSheet("setting_kategori");
  if (wsKat.getLastRow() === 0) {
    wsKat.getRange("A1:B1").setValues([["Nama Screener (Stockbit)", "Kategori Strategi"]])
      .setFontWeight("bold").setBackground("#00875A").setFontColor("#FFFFFF").setHorizontalAlignment("center");
    const sampleCategories = [
      ["Strong Uptrend", "MOMENTUM"],
      ["Bullish MACD Crossover", "MOMENTUM"],
      ["Breakout 52 Weeks High", "BREAKOUT"],
      ["Volume Breakout", "BREAKOUT"],
      ["Low PE Undervalued", "VALUE"],
      ["High Dividend Yield", "DIVIDEND"],
      ["Oversold RSI Rebound", "SWING"]
    ];
    wsKat.getRange(2, 1, sampleCategories.length, 2).setValues(sampleCategories);
    wsKat.setColumnWidth(1, 260);
    wsKat.setColumnWidth(2, 180);
    wsKat.setFrozenRows(1);
  }

  // Susun urutan tab sheet sesuai urutan baku
  reorderAllSheets();

  // Render konten awal
  renderUnifiedCalendarView();
  renderMoversDashboard();
  renderScreenerDashboard();
  updateCategoryDropdown();

  safeToast("Setup Awal Selesai! Urutan sheet berhasil diatur.", "Sukses", 5);
}

/**
 * Mengatur urutan tab di Google Spreadsheet persis sesuai permintaan:
 * 1. panduan_penggunaan
 * 2. calendar_view
 * 3. movers_dashboard
 * 4. screener_dashboard
 * 5. movers_records
 * 6. screener_records
 * 7. setting_kategori
 */
function reorderAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const desiredOrder = [
    "panduan_penggunaan",
    "calendar_view",
    "movers_dashboard",
    "screener_dashboard",
    "movers_records",
    "screener_records",
    "setting_kategori"
  ];

  for (let i = 0; i < desiredOrder.length; i++) {
    const sName = desiredOrder[i];
    const sheet = getSheetCaseInsensitive(ss, sName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }

  // Hapus Sheet1 default jika kosong
  const defaultSheet = getSheetCaseInsensitive(ss, "Sheet1") || getSheetCaseInsensitive(ss, "Sheet 1");
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (_) {}
  }

  // Aktifkan sheet 'panduan_penggunaan'
  const firstSheet = getSheetCaseInsensitive(ss, "panduan_penggunaan");
  if (firstSheet) ss.setActiveSheet(firstSheet);
}

// =========================================================================================
// 2. STRUKTUR & RENDER: UNIFIED CALENDAR VIEW (SIDE-BY-SIDE)
// =========================================================================================
function setupCalendarSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWIBTimeZone(ss);
  let wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  if (!wsCal) wsCal = ss.insertSheet("calendar_view", 0);

  // Pertahankan filter dan checklist yang sudah dipilih user sebelumnya (jika ada)
  let curYear = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy");
  const monthNames = getMonthNamesList();
  const curMonthIdx = parseInt(Utilities.formatDate(new Date(), "Asia/Jakarta", "M"), 10) - 1;
  let curMonth = monthNames[curMonthIdx];
  let curCategory = "SEMUA KATEGORI";
  let curChecksRow1 = [true, false, false, false, false];
  let curChecksRow2 = [false, false];

  try {
    const yVal = wsCal.getRange("E3").getValue();
    if (yVal) curYear = String(yVal).trim();
    const mVal = wsCal.getRange("H3").getValue();
    if (mVal) curMonth = String(mVal).trim();
    const catVal = wsCal.getRange("G6").getValue();
    if (catVal) curCategory = String(catVal).trim();

    // Ambil status checklist baris 6 jika sudah ada
    const c1 = wsCal.getRange("A6:E6").getValues()[0];
    if (Array.isArray(c1) && c1.length === 5 && typeof c1[0] === "boolean") {
      curChecksRow1 = c1;
    }
    // Ambil status checklist baris 8 jika sudah ada
    const c2 = wsCal.getRange("A8:B8").getValues()[0];
    if (Array.isArray(c2) && c2.length === 2 && typeof c2[0] === "boolean") {
      curChecksRow2 = c2;
    }
  } catch (_) {}

  wsCal.clear();
  wsCal.setFrozenRows(0);

  // Baris 1: Header Utama Terpadu
  wsCal.getRange("A1:K1").merge()
    .setValue("📅 CALENDAR VIEW - MOVERS & SCREENER")
    .setFontSize(15).setFontWeight("bold").setFontColor("#0f172a")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(1, 36);

  // Baris 2: Row Kosong (Spacer)
  wsCal.getRange("A2:K2").clear();
  wsCal.setRowHeight(2, 14);

  // Baris 3: Posisi Filter Bersama di CENTER (D3 & E3 untuk Tahun, G3 & H3 untuk Bulan)
  wsCal.getRange("A3:K3").clear();
  
  wsCal.getRange("D3").setValue("Tahun:").setFontWeight("bold").setHorizontalAlignment("right");
  wsCal.getRange("E3").setValue(curYear)
    .setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  wsCal.getRange("F3").setValue("|").setFontColor("#cbd5e1").setHorizontalAlignment("center");

  wsCal.getRange("G3").setValue("Bulan:").setFontWeight("bold").setHorizontalAlignment("right");
  wsCal.getRange("H3").setValue(curMonth)
    .setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Validasi Dropdown Tahun (E3) & Bulan (H3)
  const ruleTahun = SpreadsheetApp.newDataValidation()
    .requireValueInList(["2024", "2025", "2026", "2027", "2028", "2029", "2030"], true)
    .setAllowInvalid(true).build();
  wsCal.getRange("E3").setDataValidation(ruleTahun);

  const ruleBulan = SpreadsheetApp.newDataValidation()
    .requireValueInList(monthNames, true)
    .setAllowInvalid(true).build();
  wsCal.getRange("H3").setDataValidation(ruleBulan);
  wsCal.setRowHeight(3, 26);

  // Baris 4: Banner Sub-Header Sisi Kiri & Kanan
  wsCal.getRange("A4:E4").merge()
    .setValue("⚡ MOVERS CALENDAR & FILTER IRISAN")
    .setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F4").setBackground("#f1f5f9");

  wsCal.getRange("G4:K4").merge()
    .setValue("🎯 SCREENER CALENDAR")
    .setFontWeight("bold").setBackground("#00875A").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(4, 24);

  // Baris 5: Kategori Movers Baris 1 (Kiri) & Kategori Screener (Kanan) Sejajar
  const moversCatsRow1 = ["Top Gainer", "Top Loser", "Top Value", "Top Volume", "Top Frequency"];
  wsCal.getRange("A5:E5").setValues([moversCatsRow1])
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setHorizontalAlignment("center");
  wsCal.getRange("F5").setBackground("#f8fafc");

  wsCal.getRange("G5:H5").merge()
    .setValue("Kategori Screener:")
    .setFontWeight("bold").setFontSize(10).setBackground("#e2e8f0").setFontColor("#0f172a")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("I5:K5").merge()
    .setValue("💡 Pilih kategori strategi di bawah")
    .setFontStyle("italic").setFontSize(9).setFontColor("#475569")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(5, 22);

  // Baris 6: Checklist Movers Baris 1 (Kiri) & Dropdown SEMUA KATEGORI Screener (Kanan) Sejajar
  wsCal.getRange("A6:E6").insertCheckboxes()
    .setValues([curChecksRow1])
    .setHorizontalAlignment("center").setBackground("#f8fafc");
  wsCal.getRange("F6").setBackground("#f8fafc");

  wsCal.getRange("G6:H6").merge()
    .setValue(curCategory)
    .setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  wsCal.getRange("I6:K6").merge()
    .setValue("* Membaca dari sheet setting_kategori")
    .setFontStyle("italic").setFontSize(9).setFontColor("#64748b")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(6, 25);

  // Baris 7: Kategori Movers Baris 2 (Net Foreign Buy & Sell) & Info Emiten Berulang
  const moversCatsRow2 = ["Net Foreign Buy", "Net Foreign Sell"];
  wsCal.getRange("A7:B7").setValues([moversCatsRow2])
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setHorizontalAlignment("center");

  wsCal.getRange("C7:E7").merge()
    .setValue("💡 Filter Irisan: Emiten harus lolos SEMUA kategori dicentang")
    .setFontStyle("italic").setFontSize(9).setFontColor("#475569").setBackground("#f1f5f9")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F7").setBackground("#f8fafc");

  wsCal.getRange("G7:K7").merge()
    .setValue("🟢 Emiten hijau = muncul ≥ 2 kali dalam bulan ini")
    .setFontStyle("italic").setFontSize(9).setFontColor("#166534").setFontWeight("bold").setBackground("#f0fdf4")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(7, 22);

  // Baris 8: Checklist Movers Baris 2 (NFB & NFS) & Legend Irisan Screener (Bintang)
  wsCal.getRange("A8:B8").insertCheckboxes()
    .setValues([curChecksRow2])
    .setHorizontalAlignment("center").setBackground("#f8fafc");

  wsCal.getRange("C8:E8").merge()
    .setValue("* Centang NFB/NFS untuk melihat emiten dengan arus asing")
    .setFontStyle("italic").setFontSize(9).setFontColor("#64748b").setBackground("#f8fafc")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F8").setBackground("#f8fafc");

  wsCal.getRange("G8:K8").merge()
    .setValue("⭐ Bintang = irisan screener di hari tsb (2⭐ = 2 screener, 3⭐ = 3 screener, dst.)")
    .setFontStyle("italic").setFontSize(9).setFontColor("#b45309").setFontWeight("bold").setBackground("#fef3c7")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(8, 25);

  // Baris 9: Header Hari (Senin - Jumat) untuk Sisi Kiri & Kanan
  const dayHeaders = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"];
  wsCal.getRange("A9:E9").setValues([dayHeaders])
    .setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F9").setValue("").setBackground("#e2e8f0");

  wsCal.getRange("G9:K9").setValues([dayHeaders])
    .setFontWeight("bold").setBackground("#00875A").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.setRowHeight(9, 26);

  // Format Lebar Kolom
  for (let c = 1; c <= 5; c++) wsCal.setColumnWidth(c, 160); // A - E: Movers
  wsCal.setColumnWidth(6, 30);                                // F: Divider
  for (let c = 7; c <= 11; c++) wsCal.setColumnWidth(c, 160); // G - K: Screener

  // Format Tinggi Baris Grid Kalender (6 Minggu x 2 Baris = 12 Baris, Baris 10 s/d 21)
  for (let w = 0; w < 6; w++) {
    const rDate = 10 + (w * 2);
    const rContent = rDate + 1;
    wsCal.setRowHeight(rDate, 22);
    wsCal.setRowHeight(rContent, 65);
    wsCal.getRange("F" + rDate + ":F" + rContent).setBackground("#f8fafc");
  }
}

/**
 * Render KEDUA sisi kalender (Movers & Screener)
 */
function renderUnifiedCalendarView() {
  renderMoversCalendarPart();
  renderScreenerCalendarPart();
}

/**
 * Render Sisi Kiri: Movers Calendar
 */
function renderMoversCalendarPart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  const wsMovers = getSheetCaseInsensitive(ss, "movers_records");
  if (!wsCal || !wsMovers) return;

  // Cek apakah struktur calendar_view sudah memiliki 7 kategori (A7 = 'Net Foreign Buy')
  const a7Val = String(wsCal.getRange("A7").getValue() || "").trim();
  if (a7Val !== "Net Foreign Buy") {
    setupCalendarSheetStructure();
    wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  }

  const year = parseInt(wsCal.getRange("E3").getValue()) || new Date().getFullYear();
  const monthStr = String(wsCal.getRange("H3").getValue());
  const month = parseMonthIndex(monthStr);

  // Kategori aktif dari Baris 5 & 6 (A-E) dan Baris 7 & 8 (A-B)
  const catNamesRow1 = wsCal.getRange("A5:E5").getValues()[0];
  const catChecksRow1 = wsCal.getRange("A6:E6").getValues()[0];
  const catNamesRow2 = wsCal.getRange("A7:B7").getValues()[0];
  const catChecksRow2 = wsCal.getRange("A8:B8").getValues()[0];

  const activeCategories = new Set();
  for (let i = 0; i < catNamesRow1.length; i++) {
    if (catNamesRow1[i] && catChecksRow1[i] === true) {
      activeCategories.add(String(catNamesRow1[i]).trim().toLowerCase());
    }
  }
  for (let i = 0; i < catNamesRow2.length; i++) {
    if (catNamesRow2[i] && catChecksRow2[i] === true) {
      activeCategories.add(String(catNamesRow2[i]).trim().toLowerCase());
    }
  }

  const colsMovers = ["A", "B", "C", "D", "E"];

  // Jika tidak ada kategori yang dicentang, kosongkan grid Movers
  if (activeCategories.size === 0) {
    for (let w = 0; w < 6; w++) {
      const rContent = 10 + (w * 2) + 1;
      for (let c = 0; c < 5; c++) {
        wsCal.getRange(colsMovers[c] + rContent).setValue("").setBackground("#ffffff");
      }
    }
    return;
  }

  // Baca data Movers
  const lastRow = wsMovers.getLastRow();
  const data = lastRow > 1 ? wsMovers.getRange(2, 1, lastRow - 1, 3).getValues() : [];

  const dateEmitenCatMap = new Map();
  for (const row of data) {
    const dateYMD = parseDateToYMD(row[0]);
    const cat = String(row[1] || "").trim().toLowerCase();
    const emiten = String(row[2] || "").trim();

    if (!dateYMD || !emiten) continue;

    if (activeCategories.has(cat)) {
      if (!dateEmitenCatMap.has(dateYMD)) dateEmitenCatMap.set(dateYMD, new Map());
      const emMap = dateEmitenCatMap.get(dateYMD);
      if (!emMap.has(emiten)) emMap.set(emiten, new Set());
      emMap.get(emiten).add(cat);
    }
  }

  // Logika Irisan Penuh: emiten harus masuk ke seluruh kategori yang dicentang
  const requiredCount = activeCategories.size;
  const dateFinalEmitenMap = new Map();
  for (const [dateYMD, emMap] of dateEmitenCatMap.entries()) {
    const validList = [];
    for (const [sym, catsSet] of emMap.entries()) {
      if (catsSet.size === requiredCount) {
        validList.push(sym);
      }
    }
    if (validList.length > 0) {
      dateFinalEmitenMap.set(dateYMD, validList.sort());
    }
  }

  // Hitung Hari Senin Pertama & Batas Akhir Bulan (Gunakan jam 12:00 tengah hari agar bebas geseran timezone)
  const firstDayOfMonth = new Date(year, month - 1, 1, 12, 0, 0);
  const lastDayOfMonth = new Date(year, month, 0, 12, 0, 0);
  const dayOfWeek = firstDayOfMonth.getDay(); // 0 = Minggu, 1 = Senin, ..., 6 = Sabtu

  let startMonday;
  if (dayOfWeek === 1) {
    startMonday = new Date(year, month - 1, 1, 12, 0, 0);
  } else if (dayOfWeek === 0) {
    // Tanggal 1 adalah Minggu -> hari Senin perdagangan pertama adalah tanggal 2
    startMonday = new Date(year, month - 1, 2, 12, 0, 0);
  } else if (dayOfWeek === 6) {
    // Tanggal 1 adalah Sabtu -> hari Senin perdagangan pertama adalah tanggal 3
    startMonday = new Date(year, month - 1, 3, 12, 0, 0);
  } else {
    // Tanggal 1 adalah Selasa s/d Jumat (dayOfWeek 2..5) -> mundur ke hari Senin pada minggu tsb (misal 30 Nov untuk Des)
    startMonday = new Date(year, month - 1, 1 - (dayOfWeek - 1), 12, 0, 0);
  }

  // Kumpulkan daftar tanggal yang aktif ditampilkan di kalender untuk hitung frekuensi kemunculan
  const displayedDates = new Set();
  for (let w = 0; w < 6; w++) {
    const weekMonday = new Date(startMonday.getTime() + (w * 7 * 86400000));
    weekMonday.setHours(12, 0, 0, 0);
    if (weekMonday <= lastDayOfMonth) {
      for (let c = 0; c < 5; c++) {
        const curDate = new Date(weekMonday.getTime() + (c * 86400000));
        curDate.setHours(12, 0, 0, 0);
        displayedDates.add(formatDateYMD(curDate));
      }
    }
  }

  // Hitung Frekuensi Kemunculan per Emiten pada Minggu-Minggu Kalender yang Tampil
  const emitenFreqMap = new Map();
  for (const dateYMD of displayedDates) {
    if (dateFinalEmitenMap.has(dateYMD)) {
      for (const sym of dateFinalEmitenMap.get(dateYMD)) {
        emitenFreqMap.set(sym, (emitenFreqMap.get(sym) || 0) + 1);
      }
    }
  }

  // Text Style untuk Emiten Berulang (Stabilo Hijau Tebal) & Normal
  const boldGreenStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#15803d") // Hijau stabilo pekat
    .setBold(true)
    .build();

  const normalStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#0f172a")
    .setBold(false)
    .build();

  // Tulis ke Grid 6 Minggu (Baris 10 s/d 21)
  for (let w = 0; w < 6; w++) {
    const rDate = 10 + (w * 2);
    const rContent = rDate + 1;

    const weekMonday = new Date(startMonday.getTime() + (w * 7 * 86400000));
    weekMonday.setHours(12, 0, 0, 0);
    const isActiveWeek = weekMonday <= lastDayOfMonth;

    for (let c = 0; c < 5; c++) {
      const dateCellMovers = wsCal.getRange(colsMovers[c] + rDate);
      const contentCellMovers = wsCal.getRange(colsMovers[c] + rContent);

      if (isActiveWeek) {
        const curDate = new Date(weekMonday.getTime() + (c * 86400000));
        curDate.setHours(12, 0, 0, 0);
        const cellDateYMD = formatDateYMD(curDate);

        if (dateFinalEmitenMap.has(cellDateYMD)) {
          const emitenList = dateFinalEmitenMap.get(cellDateYMD);
          dateCellMovers.setValue(`${cellDateYMD} (${emitenList.length})`)
            .setBackground("#dbeafe") // Biru soft
            .setFontColor("#1e40af")
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          // Format Teks & Stabilo Hijau untuk Emiten yang Muncul Beberapa Kali di Movers
          let hasRepeatEmiten = false;
          const lines = [];
          const lineHighlightFlags = [];

          for (const sym of emitenList) {
            const freq = emitenFreqMap.get(sym) || 1;
            if (freq >= 2) {
              lines.push(`${sym} (${freq}x) 🟢`);
              lineHighlightFlags.push(true);
              hasRepeatEmiten = true;
            } else {
              lines.push(sym);
              lineHighlightFlags.push(false);
            }
          }

          const fullText = lines.join("\n");
          const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);

          let currentPos = 0;
          for (let i = 0; i < lines.length; i++) {
            const lineLen = lines[i].length;
            const endPos = currentPos + lineLen;
            if (lineHighlightFlags[i]) {
              richTextBuilder.setTextStyle(currentPos, endPos, boldGreenStyle);
            } else {
              richTextBuilder.setTextStyle(currentPos, endPos, normalStyle);
            }
            currentPos = endPos + 1; // newline '\n'
          }

          contentCellMovers.setRichTextValue(richTextBuilder.build())
            .setWrap(true).setVerticalAlignment("top").setHorizontalAlignment("center")
            .setBackground(hasRepeatEmiten ? "#f0fdf4" : "#ffffff") // Stabilo hijau lembut jika ada emiten berulang
            .setBorder(false, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        } else {
          dateCellMovers.setValue(cellDateYMD)
            .setBackground("#f8fafc").setFontColor("#64748b").setFontWeight("normal")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          contentCellMovers.setValue("").setBackground("#ffffff")
            .setBorder(false, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        }
      } else {
        dateCellMovers.setValue("").setBackground("#ffffff").setBorder(false, false, false, false, false, false);
        contentCellMovers.setValue("").setBackground("#ffffff").setBorder(false, false, false, false, false, false);
      }
    }
  }
}

/**
 * Render Sisi Kanan: Screener Calendar
 */
function renderScreenerCalendarPart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  const wsScreener = getSheetCaseInsensitive(ss, "screener_records");
  if (!wsCal || !wsScreener) return;

  // Cek apakah struktur calendar_view sudah memiliki 7 kategori (A7 = 'Net Foreign Buy')
  const a7Val = String(wsCal.getRange("A7").getValue() || "").trim();
  if (a7Val !== "Net Foreign Buy") {
    setupCalendarSheetStructure();
    wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  }

  const year = parseInt(wsCal.getRange("E3").getValue()) || new Date().getFullYear();
  const monthStr = String(wsCal.getRange("H3").getValue());
  const month = parseMonthIndex(monthStr);
  const selectedCat = String(wsCal.getRange("G6").getValue() || "SEMUA KATEGORI").trim();

  // Baca data Screener dengan deteksi kolom dinamis
  const lastRow = wsScreener.getLastRow();
  const lastCol = Math.max(wsScreener.getLastColumn(), 7);
  const data = lastRow > 1 ? wsScreener.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  let colDate = 1;
  let colScreener = 3;
  let colCat = 4;
  let colEmiten = 5;

  if (lastRow > 0) {
    const headers = wsScreener.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim().toLowerCase());
    const idxDate = headers.indexOf("tanggal");
    if (idxDate !== -1) colDate = idxDate;

    let idxScr = headers.indexOf("nama screener");
    if (idxScr === -1) idxScr = headers.indexOf("screener");
    if (idxScr !== -1) colScreener = idxScr;

    const idxCat = headers.indexOf("kategori");
    if (idxCat !== -1) colCat = idxCat;

    let idxEm = headers.indexOf("emiten");
    if (idxEm === -1) idxEm = headers.indexOf("symbol");
    if (idxEm === -1) idxEm = headers.indexOf("saham");
    if (idxEm !== -1) colEmiten = idxEm;
  }

  // 1. Kumpulkan semua screener unik & kategori per emiten per tanggal secara global
  // dateDailyScreeners: Map<dateYMD, Map<emiten, Set<screenerName>>>
  const dateDailyScreeners = new Map();
  // dateDailyCategories: Map<dateYMD, Map<emiten, Set<categoryLower>>>
  const dateDailyCategories = new Map();

  for (const row of data) {
    const rDate = parseDateToYMD(row[colDate]);
    const rScreener = String(row[colScreener] || "").trim();
    const rCat = String(row[colCat] || "Lainnya").trim();
    const rEmiten = String(row[colEmiten] || "").trim().toUpperCase();

    if (!rDate || !rEmiten) continue;

    if (!dateDailyScreeners.has(rDate)) dateDailyScreeners.set(rDate, new Map());
    const scrMap = dateDailyScreeners.get(rDate);
    if (!scrMap.has(rEmiten)) scrMap.set(rEmiten, new Set());
    if (rScreener) scrMap.get(rEmiten).add(rScreener);

    if (!dateDailyCategories.has(rDate)) dateDailyCategories.set(rDate, new Map());
    const catMap = dateDailyCategories.get(rDate);
    if (!catMap.has(rEmiten)) catMap.set(rEmiten, new Set());
    if (rCat) catMap.get(rEmiten).add(rCat.toLowerCase());
  }

  // 2. Filter emiten yang ditampilkan berdasarkan pilihan dropdown G6
  const isOnlyIrisan = selectedCat.toUpperCase().includes("IRISAN") || selectedCat.includes("≥ 2") || selectedCat.includes(">= 2");
  const isAllCategories = selectedCat === "SEMUA KATEGORI" || selectedCat === "SEMUA" || !selectedCat;

  // dateDisplayMap: Map<dateYMD, string[]> (Daftar emiten yang lolos filter)
  const dateDisplayMap = new Map();

  for (const [dateYMD, emMap] of dateDailyScreeners.entries()) {
    const validEmitens = [];
    for (const [sym, screenerSet] of emMap.entries()) {
      if (isOnlyIrisan) {
        // Mode Hanya Irisan: Wajib muncul di minimal 2 screener berbeda pada tanggal tsb
        if (screenerSet.size >= 2) {
          validEmitens.push(sym);
        }
      } else if (isAllCategories) {
        validEmitens.push(sym);
      } else {
        // Filter kategori spesifik
        const catSet = dateDailyCategories.get(dateYMD)?.get(sym);
        if (catSet && catSet.has(selectedCat.toLowerCase())) {
          validEmitens.push(sym);
        }
      }
    }
    if (validEmitens.length > 0) {
      dateDisplayMap.set(dateYMD, validEmitens.sort());
    }
  }

  // Hitung Hari Senin Pertama & Batas Akhir Bulan (Gunakan jam 12:00 tengah hari agar bebas geseran timezone)
  const firstDayOfMonth = new Date(year, month - 1, 1, 12, 0, 0);
  const lastDayOfMonth = new Date(year, month, 0, 12, 0, 0);
  const dayOfWeek = firstDayOfMonth.getDay(); // 0 = Minggu, 1 = Senin, ..., 6 = Sabtu

  let startMonday;
  if (dayOfWeek === 1) {
    startMonday = new Date(year, month - 1, 1, 12, 0, 0);
  } else if (dayOfWeek === 0) {
    // Tanggal 1 adalah Minggu -> hari Senin perdagangan pertama adalah tanggal 2
    startMonday = new Date(year, month - 1, 2, 12, 0, 0);
  } else if (dayOfWeek === 6) {
    // Tanggal 1 adalah Sabtu -> hari Senin perdagangan pertama adalah tanggal 3
    startMonday = new Date(year, month - 1, 3, 12, 0, 0);
  } else {
    // Tanggal 1 adalah Selasa s/d Jumat (dayOfWeek 2..5) -> mundur ke hari Senin pada minggu tsb (misal 30 Nov untuk Des)
    startMonday = new Date(year, month - 1, 1 - (dayOfWeek - 1), 12, 0, 0);
  }

  // Kumpulkan daftar tanggal yang aktif ditampilkan di kalender untuk hitung frekuensi kemunculan
  const displayedDates = new Set();
  for (let w = 0; w < 6; w++) {
    const weekMonday = new Date(startMonday.getTime() + (w * 7 * 86400000));
    weekMonday.setHours(12, 0, 0, 0);
    if (weekMonday <= lastDayOfMonth) {
      for (let c = 0; c < 5; c++) {
        const curDate = new Date(weekMonday.getTime() + (c * 86400000));
        curDate.setHours(12, 0, 0, 0);
        displayedDates.add(formatDateYMD(curDate));
      }
    }
  }

  // 3. Hitung Frekuensi Kemunculan per Emiten pada Minggu-Minggu Kalender yang Tampil
  const emitenFreqMap = new Map();
  for (const dateYMD of displayedDates) {
    if (dateDisplayMap.has(dateYMD)) {
      for (const sym of dateDisplayMap.get(dateYMD)) {
        emitenFreqMap.set(sym, (emitenFreqMap.get(sym) || 0) + 1);
      }
    }
  }

  // Text Style untuk Emiten Berulang (Stabilo Hijau Tebal) & Normal
  const boldGreenStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#15803d")
    .setBold(true)
    .build();

  const normalStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#0f172a")
    .setBold(false)
    .build();

  const colsScreener = ["G", "H", "I", "J", "K"];

  // Tulis ke Grid 6 Minggu (Baris 10 s/d 21)
  for (let w = 0; w < 6; w++) {
    const rDate = 10 + (w * 2);
    const rContent = rDate + 1;

    const weekMonday = new Date(startMonday.getTime() + (w * 7 * 86400000));
    weekMonday.setHours(12, 0, 0, 0);
    const isActiveWeek = weekMonday <= lastDayOfMonth;

    for (let c = 0; c < 5; c++) {
      const dateCell = wsCal.getRange(colsScreener[c] + rDate);
      const contentCell = wsCal.getRange(colsScreener[c] + rContent);

      if (isActiveWeek) {
        const curDate = new Date(weekMonday.getTime() + (c * 86400000));
        curDate.setHours(12, 0, 0, 0);
        const cellDateYMD = formatDateYMD(curDate);

        if (dateDisplayMap.has(cellDateYMD)) {
          const emitenList = dateDisplayMap.get(cellDateYMD);

          dateCell.setValue(`${cellDateYMD} (${emitenList.length})`)
            .setBackground("#dcfce7") // Hijau soft
            .setFontColor("#14532d")
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          // Format Teks & Stabilo Hijau untuk Emiten yang Muncul Beberapa Kali + Simbol Bintang Irisan Screener
          let hasRepeatEmiten = false;
          const lines = [];
          const lineHighlightFlags = [];

          for (const sym of emitenList) {
            const screenerSet = dateDailyScreeners.get(cellDateYMD)?.get(sym) || new Set();
            const screenerCount = screenerSet.size;
            const stars = screenerCount >= 2 ? ` ${screenerCount}⭐` : "";
            const freq = emitenFreqMap.get(sym) || 1;

            if (freq >= 2) {
              lines.push(`${sym}${stars} (${freq}x) 🟢`);
              lineHighlightFlags.push(true);
              hasRepeatEmiten = true;
            } else {
              lines.push(`${sym}${stars}`);
              lineHighlightFlags.push(false);
            }
          }

          const fullText = lines.join("\n");
          const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);

          let currentPos = 0;
          for (let i = 0; i < lines.length; i++) {
            const lineLen = lines[i].length;
            const endPos = currentPos + lineLen;
            if (lineHighlightFlags[i]) {
              richTextBuilder.setTextStyle(currentPos, endPos, boldGreenStyle);
            } else {
              richTextBuilder.setTextStyle(currentPos, endPos, normalStyle);
            }
            currentPos = endPos + 1; // newline '\n'
          }

          contentCell.setRichTextValue(richTextBuilder.build())
            .setWrap(true).setVerticalAlignment("top").setHorizontalAlignment("center")
            .setBackground(hasRepeatEmiten ? "#f0fdf4" : "#ffffff")
            .setBorder(false, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        } else {
          dateCell.setValue(cellDateYMD)
            .setBackground("#f8fafc").setFontColor("#64748b").setFontWeight("normal")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          contentCell.setValue("").setBackground("#ffffff")
            .setBorder(false, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        }
      } else {
        dateCell.setValue("").setBackground("#ffffff").setBorder(false, false, false, false, false, false);
        contentCell.setValue("").setBackground("#ffffff").setBorder(false, false, false, false, false, false);
      }
    }
  }
}

// =========================================================================================
// 3. STRUKTUR & RENDER: MOVERS DASHBOARD
// =========================================================================================
function setupMoversDashboardStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsDash = getSheetCaseInsensitive(ss, "movers_dashboard");
  if (!wsDash) wsDash = ss.insertSheet("movers_dashboard", 3);
  wsDash.clear();
  wsDash.setFrozenRows(0);

  wsDash.getRange("A1").setValue("📊 DASHBOARD REKAP & IRISAN MOVERS")
    .setFontSize(14).setFontWeight("bold").setFontColor("#1e40af");

  wsDash.getRange("A3").setValue("Dari Tanggal:").setFontWeight("bold").setHorizontalAlignment("right");
  wsDash.getRange("B3").setValue("2026-01-01").setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  wsDash.getRange("D3").setValue("Sampai Tanggal:").setFontWeight("bold").setHorizontalAlignment("right");
  wsDash.getRange("E3").setValue("2026-12-31").setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const dashHeaders = [["Emiten", "Total Muncul (Hari)", "Total Irisan (Kategori)", "Kategori Irisan", "Harga Terakhir"]];
  wsDash.getRange("A5:E5").setValues(dashHeaders)
    .setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsDash.setRowHeight(5, 26);

  wsDash.setColumnWidth(1, 100);
  wsDash.setColumnWidth(2, 160);
  wsDash.setColumnWidth(3, 170);
  wsDash.setColumnWidth(4, 280);
  wsDash.setColumnWidth(5, 120);
}

function renderMoversDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsDash = getSheetCaseInsensitive(ss, "movers_dashboard");
  const wsRecord = getSheetCaseInsensitive(ss, "movers_records");
  if (!wsDash || !wsRecord) return;

  // Jika tanggal dikosongkan/dihapus, otomatis memilih seluruh data tanpa batas
  const startVal = wsDash.getRange("B3").getValue();
  const endVal = wsDash.getRange("E3").getValue();
  const startDateStr = startVal ? parseDateToYMD(startVal) : "";
  const endDateStr = endVal ? parseDateToYMD(endVal) : "";

  const lastRowRec = wsRecord.getLastRow();
  const data = lastRowRec > 1 ? wsRecord.getRange(2, 1, lastRowRec - 1, 10).getValues() : [];

  const emitenMap = new Map();
  for (const row of data) {
    const rDate = parseDateToYMD(row[0]);
    if (!rDate) continue;
    if (startDateStr && rDate < startDateStr) continue;
    if (endDateStr && rDate > endDateStr) continue;

    const cat = String(row[1] || "").trim();
    const sym = String(row[2] || "").trim();
    const price = row[3];

    if (!sym) continue;

    if (!emitenMap.has(sym)) {
      emitenMap.set(sym, {
        symbol: sym,
        kategoriSet: new Set(),
        tanggalSet: new Set(),
        lastPrice: price
      });
    }
    const item = emitenMap.get(sym);
    if (cat) item.kategoriSet.add(cat);
    item.tanggalSet.add(rDate);
    if (price && price !== "-") item.lastPrice = price;
  }

  const resultRows = [];
  for (const [sym, d] of emitenMap.entries()) {
    resultRows.push([
      sym,
      d.tanggalSet.size,
      d.kategoriSet.size,
      Array.from(d.kategoriSet).join(", "),
      d.lastPrice
    ]);
  }

  resultRows.sort((a, b) => {
    if (b[2] !== a[2]) return b[2] - a[2]; // Irisan kategori terbesar
    return b[1] - a[1];                   // Frekuensi kemunculan terbanyak
  });

  const lastDashRow = wsDash.getLastRow();
  if (lastDashRow >= 6) {
    wsDash.getRange(6, 1, Math.max(lastDashRow - 5, 1), 5).clearContent().setBackground("#ffffff");
  }

  if (resultRows.length > 0) {
    const range = wsDash.getRange(6, 1, resultRows.length, 5);
    range.setValues(resultRows).setHorizontalAlignment("center");
    wsDash.getRange(6, 4, resultRows.length, 1).setHorizontalAlignment("left"); // Kolom daftar kategori rata kiri

    // Berikan stabilo hijau lembut untuk emiten yang muncul >= 2 hari
    const backgrounds = [];
    for (let i = 0; i < resultRows.length; i++) {
      const daysCount = resultRows[i][1];
      const rowBg = daysCount >= 2 ? ["#dcfce7", "#dcfce7", "#dcfce7", "#f0fdf4", "#f0fdf4"] : ["#ffffff", "#ffffff", "#ffffff", "#ffffff", "#ffffff"];
      backgrounds.push(rowBg);
    }
    range.setBackgrounds(backgrounds);
  }
}

// =========================================================================
// 4. STRUKTUR & RENDER: SCREENER DASHBOARD (PER TABEL KATEGORI)
// =========================================================================
function setupScreenerDashboardStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsDash = getSheetCaseInsensitive(ss, "screener_dashboard");
  if (!wsDash) wsDash = ss.insertSheet("screener_dashboard", 4);
  wsDash.clear();
  wsDash.setFrozenRows(0);

  wsDash.getRange("A1").setValue("📊 DASHBOARD REKAP & IRISAN SCREENER (PER KATEGORI)")
    .setFontSize(14).setFontWeight("bold").setFontColor("#00875A");

  wsDash.getRange("B3").setValue("Dari Tanggal:").setFontWeight("bold").setHorizontalAlignment("right");
  wsDash.getRange("C3").setValue("2026-01-01").setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  wsDash.getRange("D3").setValue("Sampai Tanggal:").setFontWeight("bold").setHorizontalAlignment("right");
  wsDash.getRange("E3").setValue("2026-12-31").setHorizontalAlignment("center").setFontWeight("bold")
    .setBackground("#fffbeb").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  wsDash.setColumnWidth(1, 110); // Emiten
  wsDash.setColumnWidth(2, 160); // Total Muncul (Hari)
  wsDash.setColumnWidth(3, 170); // Frekuensi Screener
  wsDash.setColumnWidth(4, 380); // Daftar Screener Terkait
}

function renderScreenerDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsDash = getSheetCaseInsensitive(ss, "screener_dashboard");
  const wsRecord = getSheetCaseInsensitive(ss, "screener_records");
  const wsKat = getSheetCaseInsensitive(ss, "setting_kategori");
  if (!wsDash || !wsRecord) return;

  // Jika tanggal dikosongkan/dihapus, otomatis memilih seluruh data tanpa batas
  const startVal = wsDash.getRange("C3").getValue();
  const endVal = wsDash.getRange("E3").getValue();
  const startDateStr = startVal ? parseDateToYMD(startVal) : "";
  const endDateStr = endVal ? parseDateToYMD(endVal) : "";

  const lastRowRec = wsRecord.getLastRow();
  const lastColRec = Math.max(wsRecord.getLastColumn(), 7);
  const data = lastRowRec > 1 ? wsRecord.getRange(2, 1, lastRowRec - 1, lastColRec).getValues() : [];

  let colDate = 1;
  let colScreener = 3;
  let colCat = 4;
  let colEmiten = 5;

  if (lastRowRec > 0) {
    const headers = wsRecord.getRange(1, 1, 1, lastColRec).getValues()[0].map(h => String(h || "").trim().toLowerCase());
    const idxDate = headers.indexOf("tanggal");
    if (idxDate !== -1) colDate = idxDate;

    let idxScr = headers.indexOf("nama screener");
    if (idxScr === -1) idxScr = headers.indexOf("screener");
    if (idxScr !== -1) colScreener = idxScr;

    const idxCat = headers.indexOf("kategori");
    if (idxCat !== -1) colCat = idxCat;

    let idxEm = headers.indexOf("emiten");
    if (idxEm === -1) idxEm = headers.indexOf("symbol");
    if (idxEm === -1) idxEm = headers.indexOf("saham");
    if (idxEm !== -1) colEmiten = idxEm;
  }

  // 1. Kumpulkan daftar urutan kategori dari setting_kategori agar teratur
  const categoryOrder = [];
  if (wsKat && wsKat.getLastRow() > 1) {
    const kData = wsKat.getRange(2, 2, wsKat.getLastRow() - 1, 1).getValues();
    kData.forEach(r => {
      const k = String(r[0] || "").trim();
      if (k && !categoryOrder.includes(k)) categoryOrder.push(k);
    });
  }

  // 2. Kelompokkan data per Kategori -> Emiten
  // Structure: categoryMap = Map<CategoryName, Map<Emiten, { dates: Set, screeners: Set }>>
  const categoryMap = new Map();

  for (const row of data) {
    const rDate = parseDateToYMD(row[colDate]);
    if (!rDate) continue;
    if (startDateStr && rDate < startDateStr) continue;
    if (endDateStr && rDate > endDateStr) continue;

    const rScreener = String(row[colScreener] || "").trim();
    const rCat = String(row[colCat] || "Lainnya").trim() || "Lainnya";
    const rEmiten = String(row[colEmiten] || "").trim().toUpperCase();

    if (!rEmiten) continue;

    if (!categoryMap.has(rCat)) {
      categoryMap.set(rCat, new Map());
    }
    const catGroup = categoryMap.get(rCat);
    if (!catGroup.has(rEmiten)) {
      catGroup.set(rEmiten, {
        symbol: rEmiten,
        dates: new Set(),
        screeners: new Set()
      });
    }
    const item = catGroup.get(rEmiten);
    item.dates.add(rDate);
    if (rScreener) item.screeners.add(rScreener);
  }

  // Bersihkan area tabel lama dari baris 5 ke bawah
  const maxRows = wsDash.getMaxRows();
  const lastDashRow = wsDash.getLastRow();
  if (lastDashRow >= 5) {
    wsDash.getRange(5, 1, Math.max(lastDashRow - 4, 1), 4).clear();
  }

  // Palet warna banner header per kategori agar estetik dan mudah dibedakan
  const themeColors = [
    { bg: "#00875A", text: "#FFFFFF" }, // Hijau Stockbit
    { bg: "#1e40af", text: "#FFFFFF" }, // Biru
    { bg: "#b45309", text: "#FFFFFF" }, // Amber/Cokelat
    { bg: "#7e22ce", text: "#FFFFFF" }, // Ungu
    { bg: "#0f766e", text: "#FFFFFF" }, // Teal
    { bg: "#be123c", text: "#FFFFFF" }, // Rose
    { bg: "#334155", text: "#FFFFFF" }  // Slate
  ];

  // Susun daftar kategori yang akan dirender (kategori setting dulu, baru sisanya)
  const allRenderCats = [...categoryOrder];
  for (const cat of categoryMap.keys()) {
    if (!allRenderCats.includes(cat)) {
      allRenderCats.push(cat);
    }
  }

  let currentRow = 5;
  let colorIdx = 0;

  for (const catName of allRenderCats) {
    const emMap = categoryMap.get(catName);
    const catRows = [];

    if (emMap && emMap.size > 0) {
      for (const [sym, d] of emMap.entries()) {
        catRows.push([
          sym,
          d.dates.size,
          d.screeners.size,
          Array.from(d.screeners).join(", ")
        ]);
      }

      catRows.sort((a, b) => {
        if (b[2] !== a[2]) return b[2] - a[2]; // Frekuensi screener terbanyak
        return b[1] - a[1];                   // Total hari terbanyak
      });
    }

    const color = themeColors[colorIdx % themeColors.length];
    colorIdx++;

    // 1. Banner Kategori
    const totalEmiten = catRows.length;
    wsDash.getRange(currentRow, 1, 1, 4).merge()
      .setValue(`🎯 KATEGORI: ${catName.toUpperCase()} (${totalEmiten} Emiten)`)
      .setFontWeight("bold").setFontSize(11)
      .setBackground(color.bg).setFontColor(color.text)
      .setHorizontalAlignment("left").setVerticalAlignment("middle");
    wsDash.setRowHeight(currentRow, 26);
    currentRow++;

    // 2. Table Column Headers
    const headers = [["Emiten", "Total Muncul (Hari)", "Frekuensi Screener", "Daftar Screener Terkait"]];
    wsDash.getRange(currentRow, 1, 1, 4).setValues(headers)
      .setFontWeight("bold").setFontSize(9)
      .setBackground("#f1f5f9").setFontColor("#334155")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    wsDash.setRowHeight(currentRow, 22);
    currentRow++;

    // 3. Table Rows
    if (catRows.length > 0) {
      const dataRange = wsDash.getRange(currentRow, 1, catRows.length, 4);
      dataRange.setValues(catRows).setHorizontalAlignment("center").setVerticalAlignment("middle");
      wsDash.getRange(currentRow, 4, catRows.length, 1).setHorizontalAlignment("left"); // Screener rata kiri

      // Stabilo hijau lembut untuk emiten yang muncul >= 2 hari
      const backgrounds = [];
      for (let i = 0; i < catRows.length; i++) {
        const days = catRows[i][1];
        const isHighlight = days >= 2;
        backgrounds.push(isHighlight 
          ? ["#dcfce7", "#dcfce7", "#f0fdf4", "#f0fdf4"] 
          : ["#ffffff", "#ffffff", "#ffffff", "#ffffff"]);
      }
      dataRange.setBackgrounds(backgrounds)
        .setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);

      for (let i = 0; i < catRows.length; i++) {
        wsDash.setRowHeight(currentRow + i, 21);
      }
      currentRow += catRows.length;
    } else {
      // Baris kosong jika belum ada emiten
      wsDash.getRange(currentRow, 1, 1, 4).merge()
        .setValue("Belum ada data emiten untuk rentang tanggal ini.")
        .setFontStyle("italic").setFontSize(9).setFontColor("#94a3b8")
        .setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setBackground("#ffffff")
        .setBorder(true, true, true, true, false, false, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
      wsDash.setRowHeight(currentRow, 22);
      currentRow++;
    }

    // Spacer antar tabel kategori
    wsDash.getRange(currentRow, 1, 1, 4).clear();
    wsDash.setRowHeight(currentRow, 14);
    currentRow++;
  }
}

// =========================================================================
// 5. UPDATE DROPDOWN KATEGORI SCREENER
// =========================================================================
function updateCategoryDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  const wsKat = getSheetCaseInsensitive(ss, "setting_kategori");
  if (!wsCal) return;

  const categories = ["SEMUA KATEGORI", "⭐ HANYA IRISAN (≥ 2 SCREENER)"];
  if (wsKat && wsKat.getLastRow() > 1) {
    const vals = wsKat.getRange(2, 2, wsKat.getLastRow() - 1, 1).getValues();
    vals.forEach(r => {
      const v = String(r[0] || "").trim();
      if (v && !categories.includes(v)) categories.push(v);
    });
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true)
    .setAllowInvalid(true)
    .build();

  wsCal.getRange("G6").setDataValidation(rule);
}

// =========================================================================
// 6. REFRESH SEMUA VIEW & DATA
// =========================================================================
function updateAllViews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWIBTimeZone(ss);
  renderUnifiedCalendarView();
  renderMoversDashboard();
  renderScreenerDashboard();
  updateCategoryDropdown();
  reorderAllSheets();
  safeToast("Seluruh tampilan & data berhasil di-refresh!", "Sukses", 3);
}

// =========================================================================
// 7. SHEET PANDUAN PENGGUNAAN
// =========================================================================
function setupGuideSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wsGuide = getSheetCaseInsensitive(ss, "panduan_penggunaan");
  if (!wsGuide) wsGuide = ss.insertSheet("panduan_penggunaan", 5);
  wsGuide.clear();

  wsGuide.getRange("A1").setValue("📖 PANDUAN PENGGUNAAN STOCKBIT SCREENER & MOVERS SAVER")
    .setFontSize(16).setFontWeight("bold").setFontColor("#0f172a");

  const guideText = [
    ["1. CARA MENGHUBUNGKAN DENGAN EXTENSION BROWSER (CHROME / FIREFOX):"],
    ["   a. Klik tombol 'Deploy' (di kanan atas) > 'New deployment' (Penyebaran baru)."],
    ["   b. Pilih jenis: 'Web app' (Aplikasi web)."],
    ["   c. Pada 'Execute as' (Jalankan sebagai): Pilih 'Me' (email akun Google Anda)."],
    ["   d. Pada 'Who has access' (Siapa yang memiliki akses): WAJIB pilih 'Anyone' (Siapa saja)."],
    ["   e. Klik 'Deploy', izinkan hak akses (Authorize Access), lalu salin 'Web App URL'."],
    ["   f. Buka Stockbit di Chrome/Firefox, klik ikon ⚙️ pada tombol ekstensi, lalu paste URL tersebut dan Simpan."],
    [""],
    ["2. CARA MENGGUNAKAN CALENDAR VIEW (SIDE-BY-SIDE):"],
    ["   a. Pengaturan Tahun (E3) & Bulan (H3) di posisi tengah atas berlaku untuk KEDUA kalender sekaligus."],
    ["   b. SISI KIRI (Movers): Centang kategori Movers yang ingin Anda pantau (Top Gainer, Top Loser, Top Value, Top Volume, Top Frequency, Net Foreign Buy, Net Foreign Sell). Hanya saham yang masuk irisan penuh dari seluruh kategori yang dicentang yang akan muncul."],
    ["   c. SISI KANAN (Screener): Pilih kategori screener yang ingin Anda pantau pada dropdown G6."],
    ["   d. Anda dapat melihat secara instan per baris hari: saham apa yang masuk Top Movers sekaligus lolos Screener hari itu!"],
    [""],
    ["3. CARA MENGGUNAKAN DASHBOARD:"],
    ["   - 'movers_dashboard' menyajikan frekuensi kemunculan saham dan irisan kategori movers dalam rentang tanggal tertentu."],
    ["   - 'screener_dashboard' menyajikan frekuensi kemunculan saham per strategi screener dalam rentang tanggal tertentu."],
    ["   - TIPS: Jika tanggal di dashboard dikosongkan/dihapus, dashboard otomatis memilih dan menampilkan SELURUH data."]
  ];

  wsGuide.getRange(3, 1, guideText.length, 1).setValues(guideText)
    .setFontSize(11).setFontColor("#334155");
  wsGuide.setColumnWidth(1, 850);
}

// =========================================================================
// 8. MODAL UPDATE SCRIPT DARI GITHUB
// =========================================================================
function showUpdateScriptModal() {
  const ui = SpreadsheetApp.getUi();
  const repoRawUrl = "https://raw.githubusercontent.com/imronade/Stockbit-Screener-Movers-Saver/main/google_apps_script.js";
  const repoUrl = "https://github.com/imronade/Stockbit-Screener-Movers-Saver";

  let latestCode = "";
  let fetchError = "";
  let githubVersion = "1.0.0";
  let isUpToDate = false;

  try {
    const response = UrlFetchApp.fetch(repoRawUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      latestCode = response.getContentText();
      const versionMatch = latestCode.match(/const\s+SCRIPT_VERSION\s*=\s*["']([^"']+)["']/);
      if (versionMatch) {
        githubVersion = versionMatch[1];
      }
      if (typeof SCRIPT_VERSION !== "undefined" && githubVersion === SCRIPT_VERSION) {
        isUpToDate = true;
      }
    } else {
      fetchError = `Gagal mengambil kode dari GitHub (HTTP ${response.getResponseCode()})`;
    }
  } catch (err) {
    fetchError = `Koneksi gagal: ${err.message || String(err)}`;
  }

  const currentVer = typeof SCRIPT_VERSION !== "undefined" ? SCRIPT_VERSION : "1.0.0";

  // Tampilkan notifikasi toast langsung di lembar spreadsheet
  if (isUpToDate) {
    safeToast(`Script sudah versi terbaru (v${currentVer})!`, "✅ Script Up to Date", 5);
  } else if (!fetchError) {
    safeToast(`Update tersedia! Versi v${githubVersion} tersedia di GitHub.`, "⚡ Update Tersedia", 5);
  }

  // Escape kode agar aman disematkan di textarea HTML
  const safeCode = latestCode
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 16px; color: #1e293b; background: #f8fafc; }
          .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
          .header h3 { margin: 0; font-size: 16px; color: #0f172a; display: flex; align-items: center; gap: 8px; }
          .badge-uptodate { background: #dcfce7; color: #15803d; font-size: 11px; padding: 4px 10px; border-radius: 9999px; font-weight: 700; border: 1px solid #86efac; }
          .badge-update-available { background: #fef3c7; color: #b45309; font-size: 11px; padding: 4px 10px; border-radius: 9999px; font-weight: 700; border: 1px solid #fde68a; }
          .hero-uptodate { background: #f0fdf4; border: 2px solid #86efac; border-radius: 10px; padding: 18px; text-align: center; margin-bottom: 14px; }
          .hero-update-avail { background: #fffbeb; border: 2px solid #fde68a; border-radius: 10px; padding: 16px; margin-bottom: 14px; }
          .info-grid { display: flex; justify-content: center; gap: 14px; background: #ffffff; padding: 10px 14px; border-radius: 8px; border: 1px solid #cbd5e1; text-align: left; font-size: 12px; margin-top: 12px; }
          .info-divider { border-right: 1px solid #e2e8f0; }
          .features-box { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 14px; font-size: 12px; color: #334155; margin-bottom: 14px; line-height: 1.5; }
          .steps-box { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 14px; font-size: 12px; line-height: 1.5; color: #334155; }
          .steps-box ol { margin: 0; padding-left: 18px; }
          .steps-box li { margin-bottom: 5px; }
          .copy-btn { width: 100%; background: #00875A; color: white; border: none; padding: 10px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; transition: background 0.2s; }
          .copy-btn:hover { background: #006644; }
          .btn-close { width: 100%; background: #00875A; color: white; border: none; padding: 10px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; text-align: center; }
          .btn-close:hover { background: #006644; }
          .code-preview { margin-top: 8px; width: 100%; height: 90px; box-sizing: border-box; font-family: monospace; font-size: 11px; background: #0f172a; color: #e2e8f0; padding: 10px; border-radius: 6px; border: 1px solid #334155; resize: none; }
          .toast-copied { display: none; background: #059669; color: white; text-align: center; border-radius: 4px; padding: 6px; font-size: 12px; margin-top: 8px; font-weight: 500; }
          .footer { margin-top: 14px; display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #64748b; }
          .footer a { color: #2563eb; text-decoration: none; font-weight: 500; }
          .footer a:hover { text-decoration: underline; }
        </style>
      </head>
      <body>
        <div class="header">
          <h3>🚀 Status Script Stockbit</h3>
          ${isUpToDate ? `
            <span class="badge-uptodate">✅ SCRIPT SUDAH TERBARU</span>
          ` : `
            <span class="badge-update-available">⚡ UPDATE TERSEDIA</span>
          `}
        </div>

        ${fetchError ? `
          <div style="background:#fee2e2; color:#b91c1c; padding:12px; border-radius:8px; font-size:12px; margin-bottom:12px; border:1px solid #fca5a5;">
            ⚠️ <b>${fetchError}</b><br>
            Tidak dapat memeriksa versi online. Anda tetap dapat mengunjungi <a href="${repoUrl}" target="_blank" style="color:#b91c1c; font-weight:bold; text-decoration:underline;">GitHub Repository</a> untuk melihat kode terbaru secara langsung.
          </div>
        ` : isUpToDate ? `
          <div class="hero-uptodate">
            <div style="font-size: 36px; margin-bottom: 6px;">🎉</div>
            <div style="font-size: 16px; font-weight: 700; color: #15803d; margin-bottom: 4px;">
              Script Google Spreadsheet Anda Sudah Versi Terbaru!
            </div>
            <div style="font-size: 12px; color: #166534;">
              Tidak ada pembaruan yang diperlukan. Script Anda sudah 100% sinkron dengan GitHub.
            </div>

            <div class="info-grid">
              <div>
                <span style="color:#64748b; font-size:11px;">Versi Terpasang:</span><br>
                <strong style="color:#0f172a; font-size:14px;">v${currentVer}</strong>
              </div>
              <div class="info-divider"></div>
              <div>
                <span style="color:#64748b; font-size:11px;">Versi Github:</span><br>
                <strong style="color:#0f172a; font-size:14px;">v${githubVersion}</strong>
              </div>
              <div class="info-divider"></div>
              <div>
                <span style="color:#64748b; font-size:11px;">Status:</span><br>
                <strong style="color:#15803d; font-size:13px;">🟢 Updated</strong>
              </div>
            </div>
          </div>

          <div class="features-box">
            <div style="font-weight: 700; color: #0f172a; margin-bottom: 6px;">✨ Perbaikan & Pembaruan Terkini (v${currentVer}):</div>
            <ul style="margin: 0; padding-left: 18px; line-height: 1.6;">
              <li><b>Menu Baru: 🛠️ Koreksi Tanggal Records:</b> Alat bantu 1-klik untuk memeriksa dan mengoreksi tanggal transaksi di sheet records yang sempat salah akibat jam bursa atau zona waktu.</li>
              <li><b>Sinkronisasi Otomatis Jam Buka Bursa (09.00 WIB):</b> Sebelum jam 09.00 WIB bursa hari ini belum buka, sehingga data otomatis dicatat sebagai data penutupan hari bursa sebelumnya (misal: sync Rabu &lt; 09.00 WIB = data Selasa; Senin &lt; 09.00 WIB = data Jumat).</li>
              <li><b>Sinkronisasi Akhir Pekan (Sabtu & Minggu):</b> Otomatis ditetapkan ke hari <b>Jumat</b> penutupan terakhir.</li>
              <li><b>Pengembalian Zona Waktu ke Semula:</b> Tidak mengubah setelan timezone Spreadsheet, parsing tanggal kembali mengikuti konfigurasi asli.</li>
              <li><b>Format Simbol Bintang Ringkas:</b> Irisan screener disajikan ringkas (<b>2⭐</b>, <b>3⭐</b>, dst.) hemat tempat & rapi di kolom kalender.</li>
            </ul>
          </div>

          <details style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:6px; padding:8px 12px; font-size:11px; margin-bottom:14px;">
            <summary style="cursor:pointer; font-weight:600; color:#475569;">
              📋 Ingin salin ulang kode? Klik di sini (Opsional untuk Reset / Re-Deploy)
            </summary>
            <div style="margin-top:10px;">
              <button class="copy-btn" id="copyBtn" onclick="copyLatestScript()" style="background:#475569;">
                📋 Salin Ulang Kode (v${githubVersion})
              </button>
              <div id="toastCopied" class="toast-copied">✅ Kode berhasil disalin ke Clipboard!</div>
              <textarea id="scriptSource" class="code-preview" readonly>${safeCode}</textarea>
            </div>
          </details>

          <button class="btn-close" onclick="google.script.host.close()">
            ✅ Selesai (Tutup Jendela)
          </button>
        ` : `
          <div class="hero-update-avail">
            <div style="font-size: 15px; font-weight: 700; color: #92400e; margin-bottom: 4px;">
              📢 Tersedia Pembaruan Script Baru di GitHub!
            </div>
            <div style="font-size: 12px; color: #b45309;">
              Versi baru <b>v${githubVersion}</b> membawa pembaruan fitur tampilan & stabilitas.
            </div>

            <div class="info-grid" style="border-color:#fde68a;">
              <div>
                <span style="color:#64748b; font-size:11px;">Versi Terpasang:</span><br>
                <strong style="color:#dc2626; font-size:14px;">v${currentVer} (Lama)</strong>
              </div>
              <div class="info-divider"></div>
              <div>
                <span style="color:#64748b; font-size:11px;">Versi Baru GitHub:</span><br>
                <strong style="color:#15803d; font-size:14px;">v${githubVersion} (Terbaru)</strong>
              </div>
              <div class="info-divider"></div>
              <div>
                <span style="color:#64748b; font-size:11px;">Status:</span><br>
                <strong style="color:#b45309; font-size:13px;">⚡ Perlu Update</strong>
              </div>
            </div>
          </div>

          <div class="steps-box">
            <div style="font-weight: 700; color: #0f172a; margin-bottom: 6px;">Langkah Update 2 Menit:</div>
            <ol>
              <li>Klik tombol hijau <strong>"📋 Salin Kode Terbaru (v${githubVersion})"</strong> di bawah.</li>
              <li>Di Google Sheets, buka <strong>Ekstensi (Extensions)</strong> &gt; <strong>Apps Script</strong>.</li>
              <li>Hapus semua isi file <code>Code.gs</code> (<strong>Ctrl+A</strong> lalu Hapus), kemudian Paste (<strong>Ctrl+V</strong>).</li>
              <li>Klik <strong>💾 Simpan (Save)</strong> (Ctrl+S). Jika menggunakan Web App, klik <strong>Deploy</strong> &gt; <strong>Manage deployments</strong> &gt; Edit (pensil) &gt; pilih <strong>New version</strong> &gt; klik <strong>Deploy</strong>.</li>
            </ol>
          </div>

          <button class="copy-btn" id="copyBtn" onclick="copyLatestScript()">
            📋 Salin Seluruh Kode Terbaru (v${githubVersion})
          </button>
          <div id="toastCopied" class="toast-copied">✅ Kode berhasil disalin ke Clipboard!</div>
          <textarea id="scriptSource" class="code-preview" style="height:110px;" readonly>${safeCode}</textarea>
        `}

        <div class="footer">
          <span>Repository: <a href="${repoUrl}" target="_blank">Stockbit-Screener-Movers-Saver</a></span>
          <button onclick="google.script.host.close()" style="background:transparent; border:none; color:#64748b; cursor:pointer; font-size:12px;">Tutup</button>
        </div>

        <script>
          function copyLatestScript() {
            var textarea = document.getElementById("scriptSource");
            textarea.select();
            textarea.setSelectionRange(0, 999999);
            try {
              navigator.clipboard.writeText(textarea.value).then(function() {
                showSuccess();
              }).catch(function() {
                document.execCommand("copy");
                showSuccess();
              });
            } catch (e) {
              document.execCommand("copy");
              showSuccess();
            }
          }

          function showSuccess() {
            var btn = document.getElementById("copyBtn");
            var toast = document.getElementById("toastCopied");
            btn.style.background = "#059669";
            btn.innerHTML = "✅ Berhasil Disalin ke Clipboard!";
            toast.style.display = "block";
            setTimeout(function() {
              toast.style.display = "none";
              btn.style.background = "${isUpToDate ? '#475569' : '#00875A'}";
              btn.innerHTML = "${isUpToDate ? `📋 Salin Ulang Kode (v${githubVersion})` : `📋 Salin Seluruh Kode Terbaru (v${githubVersion})`}";
            }, 3000);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(540)
    .setHeight(isUpToDate ? 470 : 540);
  ui.showModalDialog(htmlOutput, "Update Script Stockbit");
}

// =========================================================================
// 8. ALAT BANTU KOREKSI TANGGAL RECORDS (MOVERS & SCREENER)
// =========================================================================

/**
 * Menampilkan modal interaktif untuk melihat dan mengoreksi tanggal pada
 * sheet 'movers_records' dan 'screener_records' secara 1-klik.
 */
function showFixRecordsDateModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureWIBTimeZone(ss);
  const ui = SpreadsheetApp.getUi();
  const summary = getRecordsDateSummary();

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 16px; color: #1e293b; background: #f8fafc; font-size: 13px; }
          .header { margin-bottom: 14px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; }
          .header h3 { margin: 0 0 4px 0; font-size: 16px; color: #0f172a; display: flex; align-items: center; gap: 8px; }
          .header p { margin: 0; font-size: 11px; color: #64748b; }
          .card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 14px; margin-bottom: 12px; }
          .card-title { font-weight: 700; color: #0f172a; margin-bottom: 8px; font-size: 12px; display: flex; align-items: center; gap: 6px; }
          .quick-btns { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 6px; }
          .btn-quick { background: #eff6ff; border: 1px solid #bfdbfe; color: #1d4ed8; padding: 8px 10px; border-radius: 6px; font-weight: 600; font-size: 11px; cursor: pointer; text-align: left; transition: all 0.2s; }
          .btn-quick:hover { background: #dbeafe; border-color: #93c5fd; }
          .btn-quick span { display: block; font-size: 10px; color: #60a5fa; font-weight: normal; margin-top: 2px; }
          .custom-grid { display: grid; grid-template-columns: 1fr 1fr auto; gap: 8px; align-items: end; }
          .input-group label { display: block; font-size: 11px; color: #475569; margin-bottom: 4px; font-weight: 600; }
          .input-group input, .input-group select { width: 100%; box-sizing: border-box; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 12px; background: #fff; }
          .btn-submit { background: #00875A; color: white; border: none; padding: 7px 14px; border-radius: 6px; font-weight: 600; font-size: 12px; cursor: pointer; white-space: nowrap; height: 32px; }
          .btn-submit:hover { background: #006644; }
          .summary-table { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 6px; }
          .summary-table th { background: #f1f5f9; padding: 6px 8px; text-align: left; border: 1px solid #e2e8f0; color: #475569; font-weight: 600; }
          .summary-table td { padding: 5px 8px; border: 1px solid #e2e8f0; color: #1e293b; }
          .badge { background: #e0f2fe; color: #0369a1; padding: 2px 6px; border-radius: 4px; font-weight: 600; font-size: 10px; }
          .alert-box { padding: 8px 12px; border-radius: 6px; font-size: 11px; margin-bottom: 12px; display: none; line-height: 1.4; }
          .alert-success { background: #dcfce7; border: 1px solid #86efac; color: #166534; }
          .alert-error { background: #fee2e2; border: 1px solid #fca5a5; color: #991b1b; }
          .footer { display: flex; justify-content: flex-end; margin-top: 10px; }
          .btn-close { background: #64748b; color: white; border: none; padding: 6px 14px; border-radius: 6px; font-size: 12px; cursor: pointer; }
          .btn-close:hover { background: #475569; }
        </style>
      </head>
      <body>
        <div class="header">
          <h3>🛠️ Koreksi Tanggal Data Records</h3>
          <p>Perbaiki tanggal transaksi saham di sheet <b>movers_records</b> & <b>screener_records</b> secara otomatis.</p>
        </div>

        <div id="alertBox" class="alert-box"></div>

        <!-- 1. Tombol Cepat (1-Klik) -->
        <div class="card">
          <div class="card-title">⚡ Tombol Cepat Koreksi Tanggal</div>
          <div class="quick-btns">
            <button class="btn-quick" onclick="applyQuickFix('2026-11-21', '2026-11-20')">
              📅 21 Nov ➔ 20 Nov (Jumat)
              <span>Koreksi sync hari Sabtu ke penutupan Jumat</span>
            </button>
            <button class="btn-quick" onclick="applyQuickFix('2026-11-20', '2026-11-19')">
              📅 20 Nov ➔ 19 Nov (Kamis)
              <span>Koreksi sync Jumat pagi (&lt; 09.00 WIB) ke Kamis</span>
            </button>
          </div>
        </div>

        <!-- 2. Koreksi Kustom Tanggal Tertentu -->
        <div class="card">
          <div class="card-title">🎯 Koreksi Kustom Tanggal Tertentu</div>
          <div class="custom-grid">
            <div class="input-group">
              <label>Tanggal Lama (yang salah):</label>
              <input type="text" id="oldDateInput" placeholder="Contoh: 2026-11-20" />
            </div>
            <div class="input-group">
              <label>Tanggal Baru (yang benar):</label>
              <input type="text" id="newDateInput" placeholder="Contoh: 2026-11-19" />
            </div>
            <button class="btn-submit" id="btnCustom" onclick="applyCustomFix()">🔄 Ubah</button>
          </div>
          <div style="font-size: 10px; color: #94a3b8; margin-top: 6px;">
            *Format didukung: <code>YYYY-MM-DD</code> (2026-11-20) atau <code>DD/MM/YYYY</code> (20/11/2026). Kolom Hari otomatis disesuaikan!
          </div>
        </div>

        <!-- 3. Ringkasan Data Saat Ini -->
        <div class="card">
          <div class="card-title">📊 Tanggal yang Tersimpan Saat Ini di Spreadsheet</div>
          <div id="summaryContainer">
            ${renderSummaryTablesHtml(summary)}
          </div>
        </div>

        <div class="footer">
          <button class="btn-close" onclick="google.script.host.close()">Tutup</button>
        </div>

        <script>
          function showAlert(msg, isSuccess) {
            var b = document.getElementById("alertBox");
            b.style.display = "block";
            b.className = "alert-box " + (isSuccess ? "alert-success" : "alert-error");
            b.innerHTML = (isSuccess ? "✅ " : "⚠️ ") + msg;
          }

          function setButtonsDisabled(disabled) {
            var btns = document.querySelectorAll("button");
            btns.forEach(function(btn) { btn.disabled = disabled; });
          }

          function applyQuickFix(oldDate, newDate) {
            if (!confirm("Apakah Anda yakin ingin mengubah seluruh record bertanggal " + oldDate + " menjadi " + newDate + "?")) return;
            setButtonsDisabled(true);
            showAlert("Sedang mengoreksi tanggal data...", true);
            google.script.run
              .withSuccessHandler(onFixSuccess)
              .withFailureHandler(onFixError)
              .executeFixRecordDate(oldDate, newDate);
          }

          function applyCustomFix() {
            var oldDate = document.getElementById("oldDateInput").value.trim();
            var newDate = document.getElementById("newDateInput").value.trim();
            if (!oldDate || !newDate) {
              alert("Mohon isi kedua kolom Tanggal Lama dan Tanggal Baru!");
              return;
            }
            if (!confirm("Ubah data dari tanggal " + oldDate + " menjadi " + newDate + "?")) return;
            setButtonsDisabled(true);
            showAlert("Sedang mengoreksi tanggal data...", true);
            google.script.run
              .withSuccessHandler(onFixSuccess)
              .withFailureHandler(onFixError)
              .executeFixRecordDate(oldDate, newDate);
          }

          function onFixSuccess(res) {
            setButtonsDisabled(false);
            if (res && res.success) {
              showAlert(res.message, true);
              if (res.summaryHtml) {
                document.getElementById("summaryContainer").innerHTML = res.summaryHtml;
              }
            } else {
              showAlert((res && res.message) ? res.message : "Gagal mengoreksi tanggal.", false);
            }
          }

          function onFixError(err) {
            setButtonsDisabled(false);
            showAlert("Terjadi kesalahan: " + (err.message || String(err)), false);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(580)
    .setHeight(560);
  ui.showModalDialog(htmlOutput, "Koreksi Tanggal Records");
}

function renderSummaryTablesHtml(summary) {
  let h = '';
  if (!summary || (!summary.movers.length && !summary.screener.length)) {
    return '<div style="color:#94a3b8; font-style:italic;">Belum ada data di sheet records.</div>';
  }

  h += '<div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">';

  // Kolom Movers
  h += '<div><b style="color:#1e40af; font-size:11px;">📈 movers_records:</b>';
  if (!summary.movers.length) {
    h += '<div style="color:#94a3b8; font-size:10px; margin-top:4px;">Kosong</div>';
  } else {
    h += '<table class="summary-table"><tr><th>Tanggal</th><th>Jumlah</th></tr>';
    summary.movers.forEach(function(m) {
      h += '<tr><td><code>' + m.ymd + '</code> (' + m.dmy + ')</td><td><span class="badge">' + m.count + ' baris</span></td></tr>';
    });
    h += '</table>';
  }
  h += '</div>';

  // Kolom Screener
  h += '<div><b style="color:#00875A; font-size:11px;">🔍 screener_records:</b>';
  if (!summary.screener.length) {
    h += '<div style="color:#94a3b8; font-size:10px; margin-top:4px;">Kosong</div>';
  } else {
    h += '<table class="summary-table"><tr><th>Tanggal</th><th>Hari</th><th>Jumlah</th></tr>';
    summary.screener.forEach(function(s) {
      h += '<tr><td><code>' + s.ymd + '</code></td><td>' + s.day + '</td><td><span class="badge">' + s.count + ' baris</span></td></tr>';
    });
    h += '</table>';
  }
  h += '</div>';

  h += '</div>';
  return h;
}

function getRecordsDateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = { movers: [], screener: [] };

  // 1. Movers Records
  const wsMovers = getSheetCaseInsensitive(ss, "movers_records");
  if (wsMovers && wsMovers.getLastRow() > 1) {
    const data = wsMovers.getRange(2, 1, wsMovers.getLastRow() - 1, 1).getValues();
    const map = new Map();
    data.forEach(function(r) {
      const ymd = parseDateToYMD(r[0]);
      if (ymd) {
        if (!map.has(ymd)) {
          const parts = ymd.split("-");
          map.set(ymd, { ymd: ymd, dmy: parts[2] + "/" + parts[1] + "/" + parts[0], count: 0 });
        }
        map.get(ymd).count++;
      }
    });
    summary.movers = Array.from(map.values()).sort(function(a, b) { return b.ymd.localeCompare(a.ymd); });
  }

  // 2. Screener Records
  const wsScreener = getSheetCaseInsensitive(ss, "screener_records");
  if (wsScreener && wsScreener.getLastRow() > 1) {
    const lastCol = wsScreener.getLastColumn();
    const headers = wsScreener.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    let colDate = headers.indexOf("tanggal");
    let colDay = headers.indexOf("hari");
    if (colDate === -1) colDate = 1;
    if (colDay === -1) colDay = 2;

    const data = wsScreener.getRange(2, 1, wsScreener.getLastRow() - 1, Math.max(colDate, colDay) + 1).getValues();
    const map = new Map();
    data.forEach(function(r) {
      const ymd = parseDateToYMD(r[colDate]);
      if (ymd) {
        if (!map.has(ymd)) {
          map.set(ymd, { ymd: ymd, day: String(r[colDay] || getIndonesianDay(ymd)), count: 0 });
        }
        map.get(ymd).count++;
      }
    });
    summary.screener = Array.from(map.values()).sort(function(a, b) { return b.ymd.localeCompare(a.ymd); });
  }

  return summary;
}

function executeFixRecordDate(oldDateStr, newDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldYMD = parseDateToYMD(oldDateStr);
  const newYMD = parseDateToYMD(newDateStr);

  if (!oldYMD || !newYMD) {
    return { success: false, message: "Format tanggal tidak valid. Gunakan format YYYY-MM-DD atau DD/MM/YYYY." };
  }

  const parts = newYMD.split("-");
  const newDMY = `${parts[2]}/${parts[1]}/${parts[0]}`;
  const newDay = getIndonesianDay(newYMD);

  let moversCount = 0;
  let screenerCount = 0;

  // 1. Perbarui movers_records (Kolom 1: Tanggal)
  const wsMovers = getSheetCaseInsensitive(ss, "movers_records");
  if (wsMovers && wsMovers.getLastRow() > 1) {
    const lastRow = wsMovers.getLastRow();
    const range = wsMovers.getRange(2, 1, lastRow - 1, 1);
    const vals = range.getValues();
    let modified = false;

    for (let i = 0; i < vals.length; i++) {
      if (parseDateToYMD(vals[i][0]) === oldYMD) {
        vals[i][0] = newDMY;
        moversCount++;
        modified = true;
      }
    }
    if (modified) {
      range.setValues(vals);
    }
  }

  // 2. Perbarui screener_records (Kolom Tanggal & Kolom Hari)
  const wsScreener = getSheetCaseInsensitive(ss, "screener_records");
  if (wsScreener && wsScreener.getLastRow() > 1) {
    const lastRow = wsScreener.getLastRow();
    const lastCol = wsScreener.getLastColumn();
    const headers = wsScreener.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    let colDate = headers.indexOf("tanggal");
    let colDay = headers.indexOf("hari");
    if (colDate === -1) colDate = 1;
    if (colDay === -1) colDay = 2;

    const maxCol = Math.max(colDate, colDay) + 1;
    const range = wsScreener.getRange(2, 1, lastRow - 1, maxCol);
    const vals = range.getValues();
    let modified = false;

    for (let i = 0; i < vals.length; i++) {
      if (parseDateToYMD(vals[i][colDate]) === oldYMD) {
        vals[i][colDate] = newYMD;
        vals[i][colDay] = newDay;
        screenerCount++;
        modified = true;
      }
    }
    if (modified) {
      range.setValues(vals);
    }
  }

  // 3. Refresh seluruh kalender dan dashboard
  renderUnifiedCalendarView();
  renderMoversDashboard();
  renderScreenerDashboard();

  const total = moversCount + screenerCount;
  if (total === 0) {
    return {
      success: true,
      message: `Tidak ditemukan baris bertanggal ${oldYMD} di kedua sheet records.`,
      summaryHtml: renderSummaryTablesHtml(getRecordsDateSummary())
    };
  }

  return {
    success: true,
    message: `Berhasil mengoreksi ${moversCount} baris Movers dan ${screenerCount} baris Screener dari ${oldYMD} menjadi ${newYMD} (${newDay})!`,
    summaryHtml: renderSummaryTablesHtml(getRecordsDateSummary())
  };
}

/**
 * Fungsi bantuan langsung (Quick Fix) yang bisa dijalankan langsung dari editor Apps Script
 */
function fixDate21To20() {
  const res = executeFixRecordDate("2026-11-21", "2026-11-20");
  safeToast(res.message, "Koreksi Tanggal", 5);
}

function fixDate20To19() {
  const res = executeFixRecordDate("2026-11-20", "2026-11-19");
  safeToast(res.message, "Koreksi Tanggal", 5);
}
