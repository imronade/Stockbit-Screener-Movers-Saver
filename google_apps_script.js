/**
 * =========================================================================================
 * STOCKBIT SCREENER & MOVERS SAVER - ALL-IN-ONE GOOGLE APPS SCRIPT
 * =========================================================================================
 * Versi: 1.0 (Unified Side-by-Side Calendar & Centralized Multi-Tool Dashboard)
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

/**
 * Menu Kustom di Google Sheets saat file dibuka
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Stockbit Tools")
    .addItem("✨ Setup Awal / Reset", "setupTemplateViews")
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
      const dateStr = payload.date || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
      const dayStr = getIndonesianDay(new Date());
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

function parseDateToYMD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd");
  }
  if (typeof val === "number") {
    const epoch = new Date(1899, 11, 30);
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

function getIndonesianDay(date) {
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  return days[date.getDay()];
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
  let month = new Date().getMonth() + 1;
  const match = String(monthStr).match(/\((\d+)\)/);
  if (match) {
    month = parseInt(match[1]);
  } else if (!isNaN(parseInt(monthStr))) {
    month = parseInt(monthStr);
  }
  return month;
}

// =========================================================================================
// 1. SETUP TEMPLATE SEMUA SHEET (ZERO-FILE SETUP)
// =========================================================================================
function setupTemplateViews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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
  let wsCal = getSheetCaseInsensitive(ss, "calendar_view");
  if (!wsCal) wsCal = ss.insertSheet("calendar_view", 0);

  // Pertahankan filter dan checklist yang sudah dipilih user sebelumnya (jika ada)
  let curYear = String(new Date().getFullYear());
  const monthNames = getMonthNamesList();
  let curMonth = monthNames[new Date().getMonth()];
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

  // Baris 7: Kategori Movers Baris 2 (Net Foreign Buy & Sell) & Info Screener
  const moversCatsRow2 = ["Net Foreign Buy", "Net Foreign Sell"];
  wsCal.getRange("A7:B7").setValues([moversCatsRow2])
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setHorizontalAlignment("center");

  wsCal.getRange("C7:E7").merge()
    .setValue("💡 Filter Irisan: Emiten harus lolos SEMUA kategori dicentang")
    .setFontStyle("italic").setFontSize(9).setFontColor("#475569").setBackground("#f1f5f9")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F7").setBackground("#f8fafc");

  wsCal.getRange("G7:K7").merge()
    .setValue("💡 Kalender Screener menyajikan emiten berdasarkan strategi")
    .setFontStyle("italic").setFontSize(9).setFontColor("#475569").setBackground("#f1f5f9")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  wsCal.setRowHeight(7, 22);

  // Baris 8: Checklist Movers Baris 2 (NFB & NFS) & Legend Emiten Berulang
  wsCal.getRange("A8:B8").insertCheckboxes()
    .setValues([curChecksRow2])
    .setHorizontalAlignment("center").setBackground("#f8fafc");

  wsCal.getRange("C8:E8").merge()
    .setValue("* Centang NFB/NFS untuk melihat emiten dengan arus asing")
    .setFontStyle("italic").setFontSize(9).setFontColor("#64748b").setBackground("#f8fafc")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  wsCal.getRange("F8").setBackground("#f8fafc");

  wsCal.getRange("G8:K8").merge()
    .setValue("🟢 Emiten hijau = muncul ≥ 2 kali dalam bulan ini")
    .setFontStyle("italic").setFontSize(9).setFontColor("#166534").setFontWeight("bold").setBackground("#f0fdf4")
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

  // Hitung Frekuensi Kemunculan per Emiten di Bulan Ini (Untuk Deteksi Saham Berulang)
  const emitenFreqMap = new Map();
  for (const [dateYMD, emList] of dateFinalEmitenMap.entries()) {
    for (const sym of emList) {
      emitenFreqMap.set(sym, (emitenFreqMap.get(sym) || 0) + 1);
    }
  }

  // Hitung Hari Senin Pertama
  const firstDayOfMonth = new Date(year, month - 1, 1);
  let dayOfWeek = firstDayOfMonth.getDay();
  let startMonday;
  if (dayOfWeek === 6) {
    startMonday = new Date(year, month - 1, 3);
  } else if (dayOfWeek === 0) {
    startMonday = new Date(year, month - 1, 2);
  } else {
    startMonday = new Date(year, month - 1, 1 - (dayOfWeek - 1));
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

    for (let c = 0; c < 5; c++) {
      const curDate = new Date(startMonday);
      curDate.setDate(startMonday.getDate() + (w * 7) + c);

      const cellDateYMD = parseDateToYMD(curDate);
      const isSameMonth = (curDate.getMonth() + 1) === month;

      const dateCell = wsCal.getRange(colsMovers[c] + rDate);
      const contentCell = wsCal.getRange(colsMovers[c] + rContent);

      if (isSameMonth) {
        if (dateFinalEmitenMap.has(cellDateYMD)) {
          const emitenList = dateFinalEmitenMap.get(cellDateYMD);
          dateCell.setValue(`${cellDateYMD} (${emitenList.length})`)
            .setBackground("#dbeafe") // Biru soft
            .setFontColor("#1e40af")
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          // Format Teks & Stabilo Hijau untuk Emiten yang Muncul Beberapa Kali
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

          contentCell.setRichTextValue(richTextBuilder.build())
            .setWrap(true).setVerticalAlignment("top").setHorizontalAlignment("center")
            .setBackground(hasRepeatEmiten ? "#f0fdf4" : "#ffffff") // Stabilo hijau lembut jika ada emiten berulang
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

  // Baca data Screener
  const lastRow = wsScreener.getLastRow();
  const data = lastRow > 1 ? wsScreener.getRange(2, 1, lastRow - 1, 7).getValues() : [];

  const dateMap = new Map();
  for (const row of data) {
    const rDate = parseDateToYMD(row[1]);
    const rScreener = String(row[3] || "").trim();
    const rCat = String(row[4] || "").trim();
    const rEmiten = String(row[5] || "").trim();

    if (!rDate || !rEmiten) continue;

    // Filter Kategori Screener
    if (selectedCat !== "SEMUA KATEGORI" && selectedCat !== "SEMUA") {
      if (rCat.toLowerCase() !== selectedCat.toLowerCase()) continue;
    }

    if (!dateMap.has(rDate)) dateMap.set(rDate, new Map());
    const emitenMap = dateMap.get(rDate);
    if (!emitenMap.has(rEmiten)) emitenMap.set(rEmiten, new Set());
    if (rScreener) emitenMap.get(rEmiten).add(rScreener);
  }

  // Hitung Frekuensi Kemunculan per Emiten di Bulan Ini (Untuk Deteksi Saham Berulang)
  const emitenFreqMap = new Map();
  for (const [dateYMD, emMap] of dateMap.entries()) {
    for (const sym of emMap.keys()) {
      emitenFreqMap.set(sym, (emitenFreqMap.get(sym) || 0) + 1);
    }
  }

  // Hitung Hari Senin Pertama
  const firstDayOfMonth = new Date(year, month - 1, 1);
  let dayOfWeek = firstDayOfMonth.getDay();
  let startMonday;
  if (dayOfWeek === 6) {
    startMonday = new Date(year, month - 1, 3);
  } else if (dayOfWeek === 0) {
    startMonday = new Date(year, month - 1, 2);
  } else {
    startMonday = new Date(year, month - 1, 1 - (dayOfWeek - 1));
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

    for (let c = 0; c < 5; c++) {
      const curDate = new Date(startMonday);
      curDate.setDate(startMonday.getDate() + (w * 7) + c);

      const cellDateYMD = parseDateToYMD(curDate);
      const isSameMonth = (curDate.getMonth() + 1) === month;

      const dateCell = wsCal.getRange(colsScreener[c] + rDate);
      const contentCell = wsCal.getRange(colsScreener[c] + rContent);

      if (isSameMonth) {
        if (dateMap.has(cellDateYMD)) {
          const emitenMap = dateMap.get(cellDateYMD);
          const emitenList = Array.from(emitenMap.keys()).sort();

          dateCell.setValue(`${cellDateYMD} (${emitenList.length})`)
            .setBackground("#dcfce7") // Hijau soft
            .setFontColor("#14532d")
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBorder(true, true, false, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

          // Format Teks & Stabilo Hijau untuk Emiten yang Muncul Beberapa Kali
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
            currentPos = endPos + 1;
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
  const data = lastRowRec > 1 ? wsRecord.getRange(2, 1, lastRowRec - 1, 7).getValues() : [];

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
    const rDate = parseDateToYMD(row[1]);
    if (!rDate) continue;
    if (startDateStr && rDate < startDateStr) continue;
    if (endDateStr && rDate > endDateStr) continue;

    const rScreener = String(row[3] || "").trim();
    const rCat = String(row[4] || "Lainnya").trim() || "Lainnya";
    const rEmiten = String(row[5] || "").trim();

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
  if (!wsCal || !wsKat) return;

  const categories = new Set(["SEMUA KATEGORI"]);
  const lastRow = wsKat.getLastRow();
  if (lastRow > 1) {
    const vals = wsKat.getRange(2, 2, lastRow - 1, 1).getValues();
    vals.forEach(r => {
      const v = String(r[0] || "").trim();
      if (v) categories.add(v);
    });
  }

  const catList = Array.from(categories);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(catList, true)
    .setAllowInvalid(true)
    .build();

  wsCal.getRange("G6").setDataValidation(rule);
}

// =========================================================================
// 6. REFRESH SEMUA VIEW & DATA
// =========================================================================
function updateAllViews() {
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

  try {
    const response = UrlFetchApp.fetch(repoRawUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      latestCode = response.getContentText();
    } else {
      fetchError = `Gagal mengambil kode dari GitHub (HTTP ${response.getResponseCode()})`;
    }
  } catch (err) {
    fetchError = `Koneksi gagal: ${err.message || String(err)}`;
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
          .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
          .header h3 { margin: 0; font-size: 16px; color: #0f172a; display: flex; align-items: center; gap: 8px; }
          .badge { background: #dcfce7; color: #15803d; font-size: 11px; padding: 3px 8px; border-radius: 9999px; font-weight: 600; }
          .steps-box { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 12px; font-size: 12px; line-height: 1.5; color: #334155; }
          .steps-box ol { margin: 0; padding-left: 18px; }
          .steps-box li { margin-bottom: 4px; }
          .steps-box strong { color: #0f172a; }
          .copy-btn { width: 100%; background: #00875A; color: white; border: none; padding: 10px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; transition: background 0.2s; }
          .copy-btn:hover { background: #006644; }
          .copy-btn:active { transform: scale(0.99); }
          .code-preview { margin-top: 10px; width: 100%; height: 140px; box-sizing: border-box; font-family: monospace; font-size: 11px; background: #0f172a; color: #e2e8f0; padding: 10px; border-radius: 6px; border: 1px solid #334155; resize: none; }
          .footer { margin-top: 12px; display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #64748b; }
          .footer a { color: #2563eb; text-decoration: none; font-weight: 500; }
          .footer a:hover { text-decoration: underline; }
          .toast-copied { display: none; background: #059669; color: white; text-align: center; border-radius: 4px; padding: 6px; font-size: 12px; margin-top: 8px; font-weight: 500; }
        </style>
      </head>
      <body>
        <div class="header">
          <h3>🚀 Update Script Stockbit</h3>
          <span class="badge">Official GitHub</span>
        </div>

        <div class="steps-box">
          <ol>
            <li>Klik tombol hijau <strong>"📋 Salin Kode Terbaru"</strong> di bawah.</li>
            <li>Di Google Sheets, buka <strong>Extensions</strong> &gt; <strong>Apps Script</strong>.</li>
            <li>Hapus semua isi <code>Code.gs</code> (<strong>Ctrl+A</strong> lalu Hapus), kemudian Paste (<strong>Ctrl+V</strong>).</li>
            <li>Klik <strong>Save (Ctrl+S)</strong>, lalu <strong>Deploy</strong> &gt; <strong>Manage deployments</strong> &gt; Edit (pensil) &gt; pilih <strong>New version</strong> &gt; klik <strong>Deploy</strong>.</li>
          </ol>
        </div>

        ${fetchError ? `
          <div style="background:#fee2e2; color:#b91c1c; padding:8px 12px; border-radius:6px; font-size:12px; margin-bottom:10px;">
            ⚠️ ${fetchError}. <br>Silakan buka <a href="${repoUrl}" target="_blank" style="color:#b91c1c; font-weight:bold;">GitHub Repo</a> untuk salin manual.
          </div>
        ` : `
          <button class="copy-btn" id="copyBtn" onclick="copyLatestScript()">
            📋 Salin Kode Terbaru (${(latestCode.length / 1024).toFixed(1)} KB)
          </button>
          <div id="toastCopied" class="toast-copied">✅ Kode berhasil disalin ke Clipboard!</div>
          <textarea id="scriptSource" class="code-preview" readonly>${safeCode}</textarea>
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
            btn.innerHTML = "✅ Berhasil Disalin!";
            toast.style.display = "block";
            setTimeout(function() {
              toast.style.display = "none";
              btn.style.background = "#00875A";
              btn.innerHTML = "📋 Salin Kode Terbaru";
            }, 3000);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(520)
    .setHeight(420);
  ui.showModalDialog(htmlOutput, "Update Script Stockbit");
}
