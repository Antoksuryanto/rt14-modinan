/**
 * ============================================================
 *  SIMPADU RT — Backend Google Apps Script
 *  Menyimpan data aplikasi (index.html) ke Google Sheets
 *  sehingga semua warga melihat data yang SAMA.
 *
 *  CARA DEPLOY (sekali saja):
 *  1. Buka https://script.google.com  →  New project
 *  2. Tempel seluruh isi file ini ke Code.gs
 *  3. (Opsional) Isi TOKEN di bawah dengan kode rahasia Anda
 *  4. Deploy → New deployment → Web app
 *       - Execute as      : Me
 *       - Who has access  : Anyone   ← WAJIB, agar semua warga bisa akses
 *  5. Salin URL Web App (mis. https://script.google.com/macros/s/XXXX/exec)
 *     → tempel ke konstanta GAS_URL di index.html
 * ============================================================
 */

var SHEET_NAME = "SIMPADU_RT";
var TOKEN = ""; // opsional: isi kode rahasia, mis. "rt14-modinan-2026"

function doGet(e) {
  return handleRequest(e, "GET");
}

function doPost(e) {
  return handleRequest(e, "POST");
}

function handleRequest(e, method) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    if (TOKEN && e.parameter.token !== TOKEN) {
      return out.setContent(JSON.stringify({ ok: false, error: "Token tidak valid" }));
    }
    var action = e.parameter.action;
    if (action === "get") {
      out.setContent(JSON.stringify({ ok: true, data: loadAll() }));
    } else if (action === "save") {
      var payload = JSON.parse(e.postData.contents);
      if (!payload || !payload.data) throw new Error("Payload tidak lengkap");
      saveAll(payload.data);
      out.setContent(JSON.stringify({ ok: true, saved: new Date().toISOString() }));
    } else if (action === "ping") {
      out.setContent(JSON.stringify({ ok: true, pong: true, time: new Date().toISOString() }));
    } else {
      out.setContent(JSON.stringify({ ok: false, error: "Aksi tidak dikenal: " + action }));
    }
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, error: String(err) }));
  }
  return out;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["key", "value", "updated"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function loadAll() {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var data = {};
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim();
    if (!key) continue;
    try {
      data[key] = JSON.parse(values[i][1]);
    } catch (err) {
      data[key] = null;
    }
  }
  return data;
}

function saveAll(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    var rowByKey = {};
    for (var i = 1; i < values.length; i++) {
      var k = String(values[i][0] || "").trim();
      if (k) rowByKey[k] = i + 1;
    }
    var now = new Date().toISOString();
    var keys = Object.keys(data);
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      var json = JSON.stringify(data[key]);
      if (rowByKey[key]) {
        sh.getRange(rowByKey[key], 2).setValue(json);
        sh.getRange(rowByKey[key], 3).setValue(now);
      } else {
        sh.appendRow([key, json, now]);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/* ====== Utilitas admin (jalankan manual dari editor GAS) ====== */

function resetData() {
  // Hapus semua data di sheet — HATI-HATI, tidak bisa dibatalkan!
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  return "Data direset";
}

function testConnection() {
  return JSON.stringify({ ok: true, data: loadAll() });
}