// Background Service Worker untuk Chrome & Firefox (Manifest V3)
// Mengeksekusi HTTP POST ke Google Apps Script tanpa terhalang Content Security Policy (CSP)

const runtimeApi = (typeof browser !== "undefined" && browser.runtime) ? browser.runtime : chrome.runtime;

runtimeApi.onInstalled.addListener(() => {
  console.log("[Stockbit Saver] Extension terpasang/diperbarui.");
});

// Listener pesan dari content script
runtimeApi.onMessage.addListener((request, sender, sendResponse) => {
  if (request && request.action === "syncToGoogleSheets") {
    const { url, payload } = request;

    if (!url || typeof url !== "string" || !url.startsWith("http")) {
      sendResponse({ success: false, error: "URL Web App belum diatur atau tidak valid." });
      return false;
    }

    // Menggunakan text/plain agar tidak memicu preflight OPTIONS yang sering gagal di GAS
    fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(payload),
      redirect: "follow"
    })
      .then(async (response) => {
        let responseText = "";
        let json = null;
        try {
          responseText = await response.text();
          json = JSON.parse(responseText);
        } catch (_) {}

        if (response.ok) {
          sendResponse({
            success: true,
            status: response.status,
            data: json,
            raw: responseText
          });
        } else {
          sendResponse({
            success: false,
            error: `Server Google Apps Script mengembalikan HTTP ${response.status}: ${responseText.slice(0, 120)}`
          });
        }
      })
      .catch((err) => {
        console.warn("[Stockbit Saver Background Sync Error]", err);
        sendResponse({
          success: false,
          error: err.message || String(err)
        });
      });

    // Wajib return true untuk async message handling di Manifest V3
    return true;
  }
});
