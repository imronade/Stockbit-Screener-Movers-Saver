/**
 * Stockbit Screener & Movers Saver - Unified Content Script
 * Mendukung Google Chrome & Mozilla Firefox (Manifest V3)
 * Menggabungkan Ekspor/Sync Stockbit Screener & Stockbit Movers ke 1 Google Spreadsheet & Excel (.xlsx)
 */

(function () {
  'use strict';
  console.log('[Stockbit Screener & Movers Saver] Script dimuat.');

  // =========================================================================
  // 1. HELPER TOAST & NOTIFIKASI
  // =========================================================================
  function showToast(message, duration = 3500) {
    const existing = document.querySelector('.sb-toast');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = 'sb-toast';
    toast.innerHTML = `<span>${message}</span>`;
    document.body.appendChild(toast);

    setTimeout(() => {
      toast.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(12px)';
      setTimeout(() => toast.remove(), 300);
    }, duration);
  }

  // =========================================================================
  // 2. PENYIMPANAN URL WEB APP GOOGLE APPS SCRIPT (SHARED STORAGE)
  // =========================================================================
  const STORAGE_KEY = 'sb_gsheet_webhook_url';

  async function getSavedGSheetUrl() {
    return new Promise((resolve) => {
      try {
        const storageApi = (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) ? chrome.storage.local : null;
        if (storageApi) {
          storageApi.get([STORAGE_KEY, 'sb_mover_gsheet_url', 'stockbit_gsheets_webhook'], (result) => {
            if (result && result[STORAGE_KEY]) {
              resolve(result[STORAGE_KEY]);
            } else if (result && result.sb_mover_gsheet_url) {
              resolve(result.sb_mover_gsheet_url);
            } else if (result && result.stockbit_gsheets_webhook) {
              resolve(result.stockbit_gsheets_webhook);
            } else {
              resolve(
                localStorage.getItem(STORAGE_KEY) ||
                localStorage.getItem('sb_mover_gsheet_url') ||
                localStorage.getItem('stockbit_gsheets_webhook') ||
                ''
              );
            }
          });
        } else {
          resolve(
            localStorage.getItem(STORAGE_KEY) ||
            localStorage.getItem('sb_mover_gsheet_url') ||
            localStorage.getItem('stockbit_gsheets_webhook') ||
            ''
          );
        }
      } catch (_) {
        resolve('');
      }
    });
  }

  async function saveGSheetUrl(url) {
    try {
      localStorage.setItem(STORAGE_KEY, url);
      const storageApi = (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) ? chrome.storage.local : null;
      if (storageApi) {
        await storageApi.set({ [STORAGE_KEY]: url });
      }
    } catch (_) {}
  }

  // =========================================================================
  // 3. MODAL SETTINGS (TERPADU UNTUK SCREENER & MOVERS)
  // =========================================================================
  async function openSettingsModal(notice) {
    const existing = document.querySelector('.sb-modal-overlay');
    if (existing) existing.remove();

    const currentUrl = await getSavedGSheetUrl();

    const overlay = document.createElement('div');
    overlay.className = 'sb-modal-overlay';
    overlay.innerHTML = `
      <div class="sb-modal-box">
        <div class="sb-modal-header">
          <h3>⚙️ Pengaturan Web App Google Spreadsheet</h3>
          <button class="sb-modal-close" id="sb-modal-close-btn">&times;</button>
        </div>
        <div class="sb-modal-body">
          ${notice ? `<p style="color: #e53e3e; font-weight: 600;">⚠️ ${notice}</p>` : ''}
          <p>Masukkan <strong>Web App URL</strong> yang Anda peroleh dari Google Spreadsheet (menu <em>Deploy &gt; Web app</em>). URL ini berlaku untuk sinkronisasi <strong>Screener</strong> maupun <strong>Movers</strong>:</p>
          <label for="sb-gsheet-input-url">Google Apps Script Web App URL:</label>
          <input type="text" id="sb-gsheet-input-url" placeholder="https://script.google.com/macros/s/.../exec" value="${currentUrl}" />
          <small>Pastikan berawalan <code>https://script.google.com/macros/s/</code> dan berakhiran <code>/exec</code></small>
        </div>
        <div class="sb-modal-footer">
          <button class="sb-btn-secondary" id="sb-modal-cancel-btn">Batal</button>
          <button class="sb-btn-primary" id="sb-modal-save-btn">💾 Simpan URL</button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    const closeBtn = overlay.querySelector('#sb-modal-close-btn');
    const cancelBtn = overlay.querySelector('#sb-modal-cancel-btn');
    const saveBtn = overlay.querySelector('#sb-modal-save-btn');
    const inputEl = overlay.querySelector('#sb-gsheet-input-url');

    function closeModal() {
      overlay.remove();
    }

    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal();
    });

    saveBtn.addEventListener('click', async () => {
      const newUrl = (inputEl.value || '').trim();
      if (!newUrl) {
        await saveGSheetUrl('');
        showToast('ℹ️ URL Google Spreadsheet telah dikosongkan.');
        closeModal();
        return;
      }

      if (!newUrl.startsWith('https://script.google.com/macros/s/')) {
        alert('Harap masukkan URL yang valid diawali dengan: https://script.google.com/macros/s/.../exec');
        return;
      }

      await saveGSheetUrl(newUrl);
      showToast('✅ URL Google Spreadsheet berhasil disimpan!');
      closeModal();
    });
  }

  // =========================================================================
  // 4. MODUL: STOCKBIT SCREENER SAVER
  // =========================================================================
  function getScreenerTitle() {
    const headings = document.querySelectorAll('h1, h2, h3, h4');
    for (let el of headings) {
      let t = el.innerText.trim();
      if (t && !t.includes('Stockbit') && !t.includes('Screener Rules') && t.length < 100) {
        return t;
      }
    }

    const candidates = document.querySelectorAll('span, div');
    for (let el of candidates) {
      if (el.children.length === 0 && el.innerText) {
        const text = el.innerText.trim();
        const next = el.nextElementSibling;
        if (next && (next.innerText.includes('Symbols') || next.querySelector('svg'))) {
          return text;
        }
      }
    }

    return 'Stockbit Screener';
  }

  function cleanFileName(name) {
    return name.replace(/[\\/:*?"<>|]/g, '').trim();
  }

  function cleanSheetName(name) {
    let clean = name.replace(/[\\/*?[\]:]/g, '').trim();
    if (clean.length > 31) clean = clean.substring(0, 31).trim();
    return clean || 'Screener';
  }

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

  function extractScreenerTableData() {
    const tables = document.querySelectorAll('table');
    if (!tables.length) return null;

    let rowsData = [];
    tables.forEach((table) => {
      const rows = table.querySelectorAll('tr');
      rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('th, td');
        if (!cells.length) return;

        let rowCells = [];
        cells.forEach((cell) => {
          let text = cell.innerText
            .replace(/\u00A0/g, ' ')
            .replace(/\n/g, ' ')
            .trim();
          rowCells.push(text);
        });

        if (!rowsData[rowIndex]) {
          rowsData[rowIndex] = rowCells;
        } else {
          rowsData[rowIndex] = rowsData[rowIndex].concat(rowCells);
        }
      });
    });

    return rowsData.filter((r) => r.some((c) => c !== ''));
  }

  function downloadScreenerExcel() {
    if (typeof XLSX === 'undefined') {
      showToast('⚠️ Library Excel sedang dimuat, silakan coba sesaat lagi.');
      return;
    }

    const cleanRows = extractScreenerTableData();
    if (!cleanRows || !cleanRows.length) {
      showToast('⚠️ Tabel Screener belum selesai dimuat atau tidak ditemukan.');
      return;
    }

    const rawTitle = getScreenerTitle();
    const today = new Date().toISOString().slice(0, 10);
    const fileName = `${cleanFileName(rawTitle)}_${today}.xlsx`;
    const sheetName = cleanSheetName(rawTitle);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(cleanRows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast(`🎉 Sukses download Excel: ${fileName}`);
  }

  async function syncScreenerToGoogleSheets() {
    const rawUrl = await getSavedGSheetUrl();
    const webhookUrl = (rawUrl || '').trim();
    if (!webhookUrl) {
      openSettingsModal('Harap isi Web App URL terlebih dahulu.');
      return;
    }

    const cleanRows = extractScreenerTableData();
    if (!cleanRows || cleanRows.length <= 1) {
      showToast('⚠️ Tabel Screener tidak ditemukan atau belum ada data saham.');
      return;
    }

    const headers = cleanRows[0].map((h) => h.trim());
    const symbolIdx = headers.findIndex((h) => /symbol|emiten|saham/i.test(h));
    const priceIdx = headers.findIndex((h) => /price|harga|last/i.test(h));

    let rowsList = [];
    for (let i = 1; i < cleanRows.length; i++) {
      const row = cleanRows[i];
      let emiten = symbolIdx !== -1 ? row[symbolIdx] : row[0];
      if (emiten) emiten = emiten.split(' ')[0].trim();

      let price = priceIdx !== -1 ? (parseFloat(row[priceIdx].replace(/,/g, '')) || row[priceIdx]) : '';

      let metrics = {};
      headers.forEach((colName, cIdx) => {
        if (cIdx !== symbolIdx && colName) {
          metrics[colName] = row[cIdx] || '';
        }
      });

      rowsList.push({
        emiten: emiten || '',
        price: price || '',
        metrics: metrics
      });
    }

    const screenerTitle = getScreenerTitle();
    const payload = {
      type: 'screener',
      action: 'appendScreener',
      screenerName: screenerTitle,
      date: new Date().toISOString().slice(0, 10),
      rows: rowsList
    };

    showToast(`⏳ Mengirim ${rowsList.length} saham "${screenerTitle}" ke Google Sheets...`, 4000);

    const runtimeApi = (typeof chrome !== 'undefined' && chrome.runtime) ? chrome.runtime : (typeof browser !== 'undefined' ? browser.runtime : null);
    if (!runtimeApi || !runtimeApi.sendMessage) {
      showToast('⚠️ Ekstensi baru direload. Silakan tekan F5 pada halaman Stockbit.');
      return;
    }

    runtimeApi.sendMessage(
      {
        action: 'syncToGoogleSheets',
        url: webhookUrl,
        payload: payload
      },
      (res) => {
        if (chrome.runtime && chrome.runtime.lastError) {
          showToast(`❌ Gagal: ${chrome.runtime.lastError.message}`);
          return;
        }
        if (res && res.success) {
          showToast(`✅ Sukses! Data "${screenerTitle}" (${rowsList.length} emiten) tersimpan di Google Sheets.`);
        } else {
          showToast(`❌ Gagal sync: ${(res && res.error) || 'Terjadi kesalahan'}`);
        }
      }
    );
  }

  function injectScreenerFAB() {
    if (!window.location.href.includes('screener')) {
      const container = document.getElementById('sb-unified-container');
      if (container) container.remove();
      return;
    }

    if (document.getElementById('sb-unified-container')) return;

    const container = document.createElement('div');
    container.id = 'sb-unified-container';

    const menu = document.createElement('div');
    menu.id = 'sb-unified-menu';

    // 1. Download Excel
    const itemExcel = document.createElement('button');
    itemExcel.className = 'sb-menu-item';
    itemExcel.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00ab6b" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Download Screener (.xlsx)</span>
    `;
    itemExcel.onclick = () => {
      menu.style.display = 'none';
      downloadScreenerExcel();
    };

    // 2. Sync to Sheets
    const itemSync = document.createElement('button');
    itemSync.className = 'sb-menu-item sync-btn';
    itemSync.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#1a73e8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
      </svg>
      <span>Sync Screener ke Sheets</span>
    `;
    itemSync.onclick = () => {
      menu.style.display = 'none';
      syncScreenerToGoogleSheets();
    };

    const divider = document.createElement('div');
    divider.className = 'sb-menu-divider';

    // 3. Setting Web App
    const itemSetting = document.createElement('button');
    itemSetting.className = 'sb-menu-item settings-btn';
    itemSetting.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="3"></circle>
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
      </svg>
      <span>Pengaturan Web App URL</span>
    `;
    itemSetting.onclick = () => {
      menu.style.display = 'none';
      openSettingsModal();
    };

    menu.appendChild(itemExcel);
    menu.appendChild(itemSync);
    menu.appendChild(divider);
    menu.appendChild(itemSetting);

    // Floating Button
    const mainBtn = document.createElement('button');
    mainBtn.id = 'sb-main-btn';
    mainBtn.title = 'Stockbit Screener Saver';
    mainBtn.innerHTML = `
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
    `;

    mainBtn.onclick = (e) => {
      e.stopPropagation();
      menu.style.display = menu.style.display === 'flex' ? 'none' : 'flex';
    };

    document.addEventListener('click', (e) => {
      if (!container.contains(e.target)) {
        menu.style.display = 'none';
      }
    });

    container.appendChild(menu);
    container.appendChild(mainBtn);
    document.body.appendChild(container);
  }

  // =========================================================================
  // 5. MODUL: STOCKBIT MOVERS SAVER
  // =========================================================================
  function getTodayString() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${d}/${m}/${y}`;
  }

  function getTodayCompact() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}${m}${d}`;
  }

  function findMoversHeader() {
    const all = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, div, span, p'));
    return all.find((el) => {
      const txt = (el.textContent || '').trim();
      return txt === 'Movers' && el.children.length === 0;
    });
  }

  function getMoversPanelContainer() {
    const header = findMoversHeader();
    if (!header) return null;

    let p = header.parentElement;
    for (let i = 0; i < 7; i++) {
      if (!p || p === document.body) break;
      const rect = p.getBoundingClientRect();
      if (rect.width >= 200 && rect.height >= 400 && rect.left >= window.innerWidth * 0.45) {
        return p;
      }
      p = p.parentElement;
    }
    return header.parentElement;
  }

  function getMoversTabElements() {
    const panel = getMoversPanelContainer() || document.body;
    const possibleTabNames = [
      'Top Gainer',
      'Top Loser',
      'Top Value',
      'Top Volume',
      'Top Freq',
      'Top Frequency',
      'Net Foreign Buy',
      'Net Foreign Sell'
    ];

    const tabMap = new Map();
    const rightElements = Array.from(panel.querySelectorAll('*')).filter((el) => {
      const rect = el.getBoundingClientRect();
      return rect.left >= window.innerWidth * 0.45 && rect.width > 0 && rect.height > 0;
    });

    for (const name of possibleTabNames) {
      const found = rightElements.find((el) => {
        return el.children.length === 0 && (el.textContent || '').trim().toLowerCase() === name.toLowerCase();
      });
      if (found) {
        let stdName = name;
        if (name === 'Top Freq') stdName = 'Top Frequency';
        tabMap.set(stdName, found);
      }
    }
    return tabMap;
  }

  function getActiveMoversTabName() {
    const tabs = getMoversTabElements();
    for (const [name, el] of tabs.entries()) {
      const style = window.getComputedStyle(el);
      const isGreen = style.color.includes('0, 176') || style.color.includes('0, 150') || style.color.includes('rgb(0,');
      const isBold = style.fontWeight === 'bold' || parseInt(style.fontWeight || '400') >= 600;
      const hasActiveBorder = style.borderBottomColor && !style.borderBottomColor.includes('transparent') && style.borderBottomWidth !== '0px';
      if (isGreen || isBold || hasActiveBorder) {
        return name;
      }
    }
    return 'Top Gainer';
  }

  async function performMoversExtraction(onProgress, customCategory = null) {
    const activeTab = customCategory || getActiveMoversTabName();
    const panel = getMoversPanelContainer();
    if (!panel) return { activeTab, items: [] };

    const scrollableElements = Array.from(panel.querySelectorAll('*')).filter((el) => {
      const rect = el.getBoundingClientRect();
      const style = window.getComputedStyle(el);
      const isRightSide = rect.left >= window.innerWidth * 0.45;
      const isVisible = rect.width > 150 && rect.height > 200;
      const canScroll = (style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
      return isRightSide && isVisible && canScroll;
    });

    let scrollEl = scrollableElements[scrollableElements.length - 1] || null;
    const collectedData = new Map();

    const excludeTexts = new Set([
      'BUMN', 'ISSI', 'WBSA', 'IHSG', 'LQ45', 'IDX30', 'KOMPAS100', 'JII70',
      'STREAM', 'PROFILE', 'WATCHLIST', 'PORTFOLIO', 'ORDERBOOK', 'CHARTBIT', 'NEWS', 'FINANCIALS', 'VALUATION',
      'SYMBOL', 'PRICE', 'CHANGE', 'VALUE', 'VOLUME', 'FREQ', 'LAST', 'MORE'
    ]);

    function scanVisibleRows() {
      const allPanelElements = Array.from(panel.querySelectorAll('*')).filter((el) => {
        const r = el.getBoundingClientRect();
        return r.left >= window.innerWidth * 0.45 && r.width > 0 && r.height > 0;
      });

      const stockCodeElements = allPanelElements.filter((el) => {
        if (el.children.length > 0) return false;
        const txt = (el.textContent || '').trim();
        return /^[A-Z]{4}$/.test(txt) && !excludeTexts.has(txt.toUpperCase());
      });

      for (const symEl of stockCodeElements) {
        const symbol = symEl.textContent.trim();
        if (collectedData.has(symbol)) {
          const existing = collectedData.get(symbol);
          if (existing && (existing.price !== '-' || existing.value !== '-' || existing.volume !== '-')) {
            continue;
          }
        }

        let rowCandidate = symEl.parentElement;
        let matchedRow = null;
        const panelRect = panel.getBoundingClientRect();
        const minRowWidth = Math.max(160, panelRect.width * 0.55);

        for (let i = 0; i < 8; i++) {
          if (!rowCandidate || rowCandidate === panel || rowCandidate === document.body) break;
          const rect = rowCandidate.getBoundingClientRect();
          const txt = rowCandidate.innerText || '';

          const isRowWidth = rect.width >= minRowWidth;
          const isRowHeight = rect.height >= 20 && rect.height <= 95;
          const hasFinancialData = /\d+[.,]?\d*%/.test(txt) || /\d+[.,]?\d*\s*[BKM]/i.test(txt) || (txt.match(/\d+/g) || []).length >= 2;
          const codeMatches = txt.match(/\b[A-Z]{4}\b/g) || [];
          const validCodes = codeMatches.filter((c) => !excludeTexts.has(c));

          if (isRowWidth && isRowHeight && hasFinancialData && validCodes.length <= 1) {
            matchedRow = rowCandidate;
            break;
          }
          rowCandidate = rowCandidate.parentElement;
        }

        if (!matchedRow) {
          rowCandidate = symEl.parentElement;
          for (let i = 0; i < 7; i++) {
            if (!rowCandidate || rowCandidate === panel || rowCandidate === document.body) break;
            const txt = rowCandidate.innerText || '';
            const rect = rowCandidate.getBoundingClientRect();
            if (rect.height <= 95 && (/\d+[.,]?\d*%/.test(txt) || /\d+[.,]?\d*\s*[BKM]/i.test(txt))) {
              matchedRow = rowCandidate;
              break;
            }
            rowCandidate = rowCandidate.parentElement;
          }
        }

        if (!matchedRow) continue;

        const rawText = matchedRow.innerText || '';
        const rawTokens = rawText
          .replace(/[\r\n\t]+/g, ' ')
          .split(/\s+/)
          .map((t) => t.trim())
          .filter((t) => t.length > 0 && t !== symbol && !excludeTexts.has(t.toUpperCase()));

        const leafTokens = Array.from(matchedRow.querySelectorAll('*'))
          .filter((e) => e.children.length === 0 && e.textContent && e.textContent.trim().length > 0)
          .map((e) => e.textContent.trim())
          .filter((t) => t !== symbol && !excludeTexts.has(t.toUpperCase()));

        let tokens = leafTokens.length > rawTokens.length ? leafTokens : rawTokens;

        let price = '-';
        let idxPoints = '-';
        let idxMov = '-';
        let value = '-';
        let volume = '-';
        let freq = '-';
        let netForeign = '-';

        const cleanTokens = [];
        for (const tok of tokens) {
          const combo = tok.match(/^([0-9.,]+)\s*\(([-+]?[0-9.,]+%)\)$/);
          if (combo) {
            cleanTokens.push(combo[1]);
            cleanTokens.push(combo[2]);
            continue;
          }
          const combo2 = tok.match(/^([0-9.,]+)\s*([-+][0-9.,]+%)$/);
          if (combo2) {
            cleanTokens.push(combo2[1]);
            cleanTokens.push(combo2[2]);
            continue;
          }
          cleanTokens.push(tok);
        }

        const kmbTokens = [];
        const pctTokens = [];
        const numTokens = [];
        const signTokens = [];

        for (const tok of cleanTokens) {
          if (/^[+-]?[0-9.,]+[BKM]$/i.test(tok)) {
            kmbTokens.push(tok);
          } else if (tok.includes('%')) {
            pctTokens.push(tok.replace(/[()]/g, ''));
          } else if (/^[+-][0-9.,]+$/.test(tok)) {
            signTokens.push(tok);
          } else if (/^[0-9.,]+$/.test(tok)) {
            numTokens.push(tok);
          }
        }

        if (numTokens.length > 0) {
          price = numTokens[0];
        }

        if (activeTab === 'Top Gainer' || activeTab === 'Top Loser') {
          if (signTokens.length > 0) idxPoints = signTokens[0];
          if (pctTokens.length >= 2) {
            idxMov = pctTokens[1];
          } else if (pctTokens.length === 1) {
            const p0 = pctTokens[0] || '';
            if (p0.startsWith('+') || p0.startsWith('-')) idxMov = p0;
          }

          if (kmbTokens.length >= 4) {
            value = kmbTokens[0];
            volume = kmbTokens[1];
            freq = kmbTokens[2];
            netForeign = kmbTokens[3];
          } else if (kmbTokens.length === 3) {
            value = kmbTokens[0];
            volume = kmbTokens[1];
            freq = kmbTokens[2];
          } else if (kmbTokens.length === 2) {
            value = kmbTokens[0];
            volume = kmbTokens[1];
            if (numTokens.length > 1) freq = numTokens[1];
          } else if (kmbTokens.length === 1) {
            value = kmbTokens[0];
          }
        } else if (activeTab === 'Top Value') {
          if (kmbTokens.length >= 4) {
            value = kmbTokens[0];
            volume = kmbTokens[1];
            freq = kmbTokens[2];
            netForeign = kmbTokens[3];
          } else if (kmbTokens.length === 3) {
            value = kmbTokens[0];
            volume = kmbTokens[1];
            freq = kmbTokens[2];
          }
        } else if (activeTab === 'Top Volume') {
          if (kmbTokens.length >= 4) {
            volume = kmbTokens[0];
            value = kmbTokens[1];
            freq = kmbTokens[2];
            netForeign = kmbTokens[3];
          } else if (kmbTokens.length === 3) {
            volume = kmbTokens[0];
            value = kmbTokens[1];
            freq = kmbTokens[2];
          }
        } else if (activeTab === 'Top Frequency') {
          if (kmbTokens.length >= 4) {
            freq = kmbTokens[0];
            value = kmbTokens[1];
            volume = kmbTokens[2];
            netForeign = kmbTokens[3];
          } else if (kmbTokens.length === 3) {
            freq = kmbTokens[0];
            value = kmbTokens[1];
            volume = kmbTokens[2];
          }
        } else if (activeTab === 'Net Foreign Buy' || activeTab === 'Net Foreign Sell') {
          if (kmbTokens.length >= 4) {
            netForeign = kmbTokens[0];
            value = kmbTokens[1];
            volume = kmbTokens[2];
            freq = kmbTokens[3];
          } else if (kmbTokens.length === 3) {
            netForeign = kmbTokens[0];
            value = kmbTokens[1];
            volume = kmbTokens[2];
          }
        }

        const hasValidData = price !== '-' || value !== '-' || volume !== '-' || idxPoints !== '-' || idxMov !== '-';

        if (hasValidData || !collectedData.has(symbol)) {
          collectedData.set(symbol, {
            tanggal: getTodayString(),
            kategori: activeTab,
            symbol,
            price,
            idxPoints,
            idxMov,
            value,
            volume,
            freq,
            netForeign
          });
        }
      }
    }

    if (scrollEl) {
      const maxSteps = 45;
      let prevScrollTop = -1;
      let stuck = 0;

      for (let s = 0; s < maxSteps; s++) {
        scanVisibleRows();
        if (onProgress) onProgress(`[${activeTab}] Mengambil ${collectedData.size} saham...`);

        if (scrollEl.scrollTop === prevScrollTop && s > 0) {
          stuck++;
          if (stuck >= 2) break;
        } else {
          stuck = 0;
        }

        prevScrollTop = scrollEl.scrollTop;
        scrollEl.scrollTop += 320;
        await new Promise((r) => setTimeout(r, 200));
      }
      scrollEl.scrollTop = 0;
    } else {
      scanVisibleRows();
    }

    return {
      activeTab,
      items: Array.from(collectedData.values())
    };
  }

  async function crawlAllMoversTabs(onProgress) {
    const tabMap = getMoversTabElements();
    const tabList = Array.from(tabMap.keys());
    const allRecords = [];

    if (tabList.length === 0) {
      const res = await performMoversExtraction(onProgress);
      return res.items;
    }

    for (let i = 0; i < tabList.length; i++) {
      const tabName = tabList[i];
      const el = tabMap.get(tabName);

      if (onProgress) onProgress(`Membuka tab (${i + 1}/${tabList.length}): ${tabName}...`);

      if (el) {
        el.click();
        await new Promise((r) => setTimeout(r, 1000));
      }

      const res = await performMoversExtraction(onProgress, tabName);
      for (const item of res.items) {
        allRecords.push(item);
      }
    }

    return allRecords;
  }

  function generateMoversDashboardData(allRecords) {
    const emitenMap = new Map();

    for (const row of allRecords) {
      const sym = row.symbol;
      if (!emitenMap.has(sym)) {
        emitenMap.set(sym, {
          symbol: sym,
          kategoriSet: new Set(),
          tanggalSet: new Set(),
          lastPrice: row.price
        });
      }
      const item = emitenMap.get(sym);
      item.kategoriSet.add(row.kategori);
      item.tanggalSet.add(row.tanggal);
      if (row.price && row.price !== '-') item.lastPrice = row.price;
    }

    const dashboardRows = [];
    for (const [sym, data] of emitenMap.entries()) {
      dashboardRows.push({
        Emiten: sym,
        'Total Muncul (Hari)': data.tanggalSet.size,
        'Total Irisan (Kategori)': data.kategoriSet.size,
        'Kategori Irisan': Array.from(data.kategoriSet).join(', '),
        'Harga Terakhir': data.lastPrice
      });
    }

    dashboardRows.sort((a, b) => {
      if (b['Total Irisan (Kategori)'] !== a['Total Irisan (Kategori)']) {
        return b['Total Irisan (Kategori)'] - a['Total Irisan (Kategori)'];
      }
      return b['Total Muncul (Hari)'] - a['Total Muncul (Hari)'];
    });

    return dashboardRows;
  }

  function exportMoversWorkbookXLSX(allRecords) {
    if (typeof XLSX === 'undefined') {
      showToast('⚠️ SheetJS belum terpasang atau gagal dimuat.');
      return;
    }

    const wb = XLSX.utils.book_new();

    const recordHeaders = ['Tanggal', 'Kategori', 'Emiten', 'Price', 'IDX Points', '% IDX Mov', 'Value', 'Volume', 'Freq', 'Net Foreign'];
    const recordRows = allRecords.map((r) => [
      r.tanggal,
      r.kategori,
      r.symbol,
      r.price,
      r.idxPoints,
      r.idxMov,
      r.value,
      r.volume,
      r.freq,
      r.netForeign
    ]);
    const wsRecord = XLSX.utils.aoa_to_sheet([recordHeaders, ...recordRows]);
    XLSX.utils.book_append_sheet(wb, wsRecord, 'movers_records');

    const dashboardData = generateMoversDashboardData(allRecords);
    const dashHeaders = ['Emiten', 'Total Muncul (Hari)', 'Total Irisan (Kategori)', 'Kategori Irisan', 'Harga Terakhir'];
    const dashRows = dashboardData.map((d) => [
      d['Emiten'],
      d['Total Muncul (Hari)'],
      d['Total Irisan (Kategori)'],
      d['Kategori Irisan'],
      d['Harga Terakhir']
    ]);
    const wsDash = XLSX.utils.aoa_to_sheet([dashHeaders, ...dashRows]);
    const dashSheetName = `dashboard_${getTodayCompact()}`;
    XLSX.utils.book_append_sheet(wb, wsDash, dashSheetName);

    const filename = `Stockbit_Movers_Dashboard_${getTodayCompact()}.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  function injectMoversButtons() {
    if (document.getElementById('sb-mover-btn-group')) return;

    const header = findMoversHeader();
    if (!header) return;

    const container = document.createElement('div');
    container.id = 'sb-mover-btn-group';
    container.className = 'sb-mover-btn-container';

    // 1. Tombol Download Excel (.xlsx)
    const btnAllXLSX = document.createElement('button');
    btnAllXLSX.className = 'sb-mover-download-btn sb-mover-all-btn';
    btnAllXLSX.innerHTML = `<span>📊 Download Movers (.xlsx)</span>`;
    btnAllXLSX.title = "Otomatis crawl seluruh tab ke Excel";

    btnAllXLSX.addEventListener('click', async (e) => {
      e.stopPropagation();
      e.preventDefault();
      try {
        btnAllXLSX.disabled = true;
        btnSync.disabled = true;
        btnAllXLSX.innerHTML = `<span>⏳ Proses...</span>`;
        showToast('Memulai auto-crawl semua tab Movers...');

        const allRecords = await crawlAllMoversTabs((msg) => {
          btnAllXLSX.innerHTML = `<span>⏳ ${msg}</span>`;
        });

        if (!allRecords || allRecords.length === 0) {
          showToast('⚠️ Gagal mengumpulkan data Movers.');
          return;
        }

        showToast('Menyusun file Excel (.xlsx)...');
        exportMoversWorkbookXLSX(allRecords);
        showToast(`🎉 Sukses! ${allRecords.length} baris Movers tersimpan ke Excel.`);
      } catch (err) {
        console.error(err);
        showToast('❌ Terjadi kesalahan saat crawl Movers.');
      } finally {
        btnAllXLSX.disabled = false;
        btnSync.disabled = false;
        btnAllXLSX.innerHTML = `<span>📊 Download Movers (.xlsx)</span>`;
      }
    });

    // 2. Tombol Sync ke Google Sheets
    const btnSync = document.createElement('button');
    btnSync.className = 'sb-mover-download-btn sb-mover-sync-btn';
    btnSync.innerHTML = `<span>☁️ Sync Movers ke Sheets</span>`;
    btnSync.title = 'Kirim seluruh data tab movers hari ini langsung ke Google Spreadsheet';

    btnSync.addEventListener('click', async (e) => {
      e.stopPropagation();
      e.preventDefault();

      let rawUrl = await getSavedGSheetUrl();
      let savedUrl = (rawUrl || '').trim();
      if (!savedUrl) {
        showToast('⚙️ Harap isi URL Web App Google Sheets terlebih dahulu.');
        openSettingsModal();
        return;
      }

      try {
        btnSync.disabled = true;
        btnAllXLSX.disabled = true;
        btnSync.innerHTML = `<span>⏳ Crawl data...</span>`;
        showToast('Memulai crawl tab untuk sinkronisasi ke Google Sheets...');

        const allRecords = await crawlAllMoversTabs((msg) => {
          btnSync.innerHTML = `<span>⏳ ${msg}</span>`;
        });

        if (!allRecords || allRecords.length === 0) {
          showToast('⚠️ Tidak ada data movers yang berhasil dikumpulkan.');
          return;
        }

        btnSync.innerHTML = `<span>⏳ Mengirim...</span>`;
        showToast(`Mengirim ${allRecords.length} baris Movers ke Google Sheets...`);

        const payload = {
          type: 'movers',
          action: 'appendMovers',
          records: allRecords
        };

        const runtimeApi = (typeof chrome !== 'undefined' && chrome.runtime) ? chrome.runtime : (typeof browser !== 'undefined' ? browser.runtime : null);
        if (!runtimeApi || !runtimeApi.sendMessage) {
          showToast('⚠️ Ekstensi baru direload. Silakan tekan F5 pada halaman Stockbit.');
          return;
        }

        const bgResponse = await new Promise((resolve) => {
          try {
            runtimeApi.sendMessage(
              {
                action: 'syncToGoogleSheets',
                url: savedUrl,
                payload: payload
              },
              (res) => {
                if (chrome.runtime && chrome.runtime.lastError) {
                  const errMsg = chrome.runtime.lastError.message || '';
                  resolve({ success: false, error: errMsg });
                } else {
                  resolve(res || { success: true });
                }
              }
            );
          } catch (sendErr) {
            resolve({
              success: false,
              error: 'Koneksi ekstensi terputus. Harap Refresh halaman Stockbit (F5).'
            });
          }
        });

        if (bgResponse && bgResponse.success) {
          const count = bgResponse.data && bgResponse.data.count ? bgResponse.data.count : allRecords.length;
          showToast(`🎉 Sukses Sync! ${count} baris Movers telah masuk ke Google Sheets.`);
        } else {
          const detail = bgResponse && bgResponse.error ? bgResponse.error : 'Gagal mengirim data.';
          showToast(`❌ ${detail}`);
        }
      } catch (err) {
        console.error('[Stockbit Movers Sync Error]', err);
        const errDetail = err && err.message ? err.message : String(err);
        showToast(`❌ Gagal sync: ${errDetail.slice(0, 80)}`);
      } finally {
        btnSync.disabled = false;
        btnAllXLSX.disabled = false;
        btnSync.innerHTML = `<span>☁️ Sync Movers ke Sheets</span>`;
      }
    });

    // 3. Tombol Settings
    const btnSettings = document.createElement('button');
    btnSettings.className = 'sb-mover-download-btn sb-mover-settings-btn';
    btnSettings.innerHTML = `<span>⚙️</span>`;
    btnSettings.title = 'Pengaturan URL Web App Google Sheets';
    btnSettings.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      openSettingsModal();
    });

    container.appendChild(btnAllXLSX);
    container.appendChild(btnSync);
    container.appendChild(btnSettings);

    header.parentElement.insertBefore(container, header.nextSibling);
  }

  // =========================================================================
  // 6. OBSERVER & AUTO-INJECTION CYCLE
  // =========================================================================
  function runInjectionChecks() {
    injectScreenerFAB();
    injectMoversButtons();
  }

  runInjectionChecks();

  const observer = new MutationObserver(() => {
    runInjectionChecks();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });

  setInterval(runInjectionChecks, 2500);
})();
