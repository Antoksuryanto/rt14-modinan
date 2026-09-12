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

/* ===== KONFIGURASI SHEET PER KATEGORI (Level 2: Sheets jadi DB utama) =====
 * Setiap modul = sheet tersendiri, kolom = field, baris = 1 record.
 * Frontend (index.html) tetap kirim JSON per modul — GAS konversi ke tabel.
 * Kolom pertama tiap sheet = "id" (untuk array) atau "key" (untuk scalar/object).
 */
var MODULES = ["profil","users","warga","pengurus","periodeList","pengumuman","agenda","arisan","danaSosial","transaksi","iuran","pengaduan","berita","galeri","arsip","notifikasi"];
var SCALAR_KEYS = ["nextId","nextTid"]; // nilai scalar, bukan array

/* Sheet & kolom header per modul (urutan = urutan kolom) */
var MODULE_COLUMNS = {
  profil: ["key","value"],
  users: ["id","u","p","role","nama","jab"],
  warga: ["id","nama","alamat","hp","tinggal","status","jk","anggota"],
  pengurus: ["id","jab","nama","hp"],
  periodeList: ["id","mulai","selesai","thn","ket","pengurus"],
  pengumuman: ["id","judul","kat","tgl","penulis","isi"],
  agenda: ["id","judul","kat","tgl","jam","tempat","desc"],
  arisan: ["id","bulan","penerima","keterangan","jumlah","selanjutnya"],
  danaSosial: ["id","nama","jumlah","tanggal"],
  transaksi: ["id","tgl","jenis","kat","ket","jumlah","bukti"],
  iuran: ["id","kkId","bulan","status","tgl"],
  pengaduan: ["id","pelapor","kat","judul","desc","tgl","status","catatan"],
  berita: ["id","judul","kat","tgl","foto","excerpt","isi"],
  galeri: ["id","foto","judul","kat","tgl"],
  arsip: ["id","thn","acara"],
  notifikasi: ["id","tipe","judul","pesan","waktu","read"]
};

function doGet(e) {
  pastikanBackupTerkonfigurasi_();
  return handleRequest(e, "GET");
}

function doPost(e) {
  pastikanBackupTerkonfigurasi_();
  return handleRequest(e, "POST");
}

/* Setelah deploy (sekali authorize), backup harian otomatis langsung aktif.
 * Self-install: cek flag — jika belum ada trigger backupOtomatis, pasang. */
function pastikanBackupTerkonfigurasi_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty("trig_backup_ok") === "1") return;
    var ada = false;
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === "backupOtomatis" && t.getEventType() === ScriptApp.EventType.CLOCK) ada = true;
    });
    if (!ada) {
      ScriptApp.newTrigger("backupOtomatis")
        .timeBased().everyDays(1).atHour(2).create();
    }
    props.setProperty("trig_backup_ok", "1");
  } catch (err) {
    // Izin trigger belum di-authorize — trigger menyusul saat Run manual pasangTriggerBackup()
  }
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
      out.setContent(JSON.stringify({ ok: true, data: loadAll(), rev: loadRev() }));
    } else if (action === "save") {
      var payload = JSON.parse(e.postData.contents);
      if (!payload || !payload.data) throw new Error("Payload tidak lengkap");
      saveAll(payload.data, payload.rev || null);
      out.setContent(JSON.stringify({ ok: true, saved: new Date().toISOString() }));
    } else if (action === "ping") {
      out.setContent(JSON.stringify({ ok: true, pong: true, time: new Date().toISOString() }));
    } else if (action === "backup") {
      var url = buatBackupSpreadsheet();
      out.setContent(JSON.stringify({ ok: true, url: url }));
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
    sh.appendRow(["key", "value", "updated", "rev"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ====== LEVEL 2: SHEET PER KATEGORI (Sheets jadi DB utama) ======
 * Setiap modul punya sheet sendiri (warga, users, pengurus, ...).
 * Baris 1 = header kolom (field), baris 2+ = 1 record per baris.
 * Frontend tetap kirim JSON per modul — GAS konversi ke tabel.
 * Anda bisa inject/isi data langsung di Google Sheets.
 */

function getSheetByKey_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = namaSheet_(key);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var cols = MODULE_COLUMNS[key] || ["key", "value"];
    sh.appendRow(cols);
    sh.setFrozenRows(1);
  }
  return sh;
}

function loadAll() {
  var data = {};
  var props = PropertiesService.getScriptProperties();
  // 1) Baca dari sheet per kategori (format baru)
  MODULES.forEach(function(m) {
    var sh = ssGetSheetByName_(m);
    if (!sh) return;
    if (m === "profil") {
      data[m] = bacaObject_(sh);
    } else {
      data[m] = bacaSheet_(sh, m);
    }
    // Deteksi perubahan eksternal (inject manual): jika hash sheet != hash terakhir saveAll
    var curHash = hashKonten_(sh.getDataRange().getValues());
    var lastSavedHash = Number(props.getProperty("hash_" + m) || 0);
    if (curHash !== lastSavedHash && lastSavedHash !== 0) {
      // Sheet diubah manual → bump rev supaya frontend pull data server
      props.setProperty("rev_" + m, String(Date.now()));
    }
  });
  SCALAR_KEYS.forEach(function(k) {
    var sh = ssGetSheetByName_(k);
    if (!sh) return;
    data[k] = bacaScalar_(sh);
    var curHash = hashKonten_(sh.getDataRange().getValues());
    var lastSavedHash = Number(props.getProperty("hash_" + k) || 0);
    if (curHash !== lastSavedHash && lastSavedHash !== 0) {
      props.setProperty("rev_" + k, String(Date.now()));
    }
  });
  // 2) Migrasi satu kali dari format lama (blob JSON) jika belum pernah
  if (props.getProperty("migrated_tabel") !== "1") {
    var legacy = getSheet_();
    var values = legacy.getDataRange().getValues();
    var adaLegacy = false;
    for (var i = 1; i < values.length; i++) {
      var key = String(values[i][0] || "").trim();
      if (!key) continue;
      adaLegacy = true;
      if (data[key] === undefined) {
        try { data[key] = JSON.parse(values[i][1]); } catch (err) { data[key] = null; }
      }
    }
    if (adaLegacy) {
      migrasiKeTabel_();
      props.setProperty("migrated_tabel", "1");
    }
  }
  // 3) Isi sheet kosong dari data lokal (jika sheet ada tapi kosong — mis. arisan/profil)
  //    Ini menutup kasus: sheet dibuat tapi data tidak pernah tertulis (race hash lama)
  MODULES.forEach(function(m) {
    var sh = ssGetSheetByName_(m);
    if (!sh) return;
    var isi = sh.getDataRange().getValues();
    var kosong = isi.length < 2 || (isi.length === 1 && isi[0].join("").trim() === "");
    if (kosong && data[m] !== undefined && data[m] !== null) {
      var isiData = (m === "profil" && typeof data[m] === "object" && !Array.isArray(data[m]))
        ? Object.keys(data[m]).length > 0
        : Array.isArray(data[m]) && data[m].length > 0;
      if (isiData) {
        tulisSheet_(m, data[m]);
        props.setProperty("hash_" + m, String(hashKonten_(sh.getDataRange().getValues())));
        props.setProperty("rev_" + m, String(Date.now()));
      }
    }
  });
  return data;
}

/* Baca sheet format key|value → object (untuk profil) */
function bacaObject_(sh) {
  var values = sh.getDataRange().getValues();
  var obj = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || "").trim();
    if (!k) continue;
    var v = values[i][1];
    if (v === "" || v === null || v === undefined) { obj[k] = ""; continue; }
    if (v instanceof Date) {
      obj[k] = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
      continue;
    }
    if (typeof v === "string" && (v.charAt(0) === "[" || v.charAt(0) === "{")) {
      try { obj[k] = JSON.parse(v); } catch (e) { obj[k] = v; }
    } else {
      obj[k] = v;
    }
  }
  return obj;
}

function ssGetSheetByName_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet_(name));
}

function bacaSheet_(sh, m) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var arr = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0] && !row[1]) continue; // baris kosong
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      var v = row[j];
      if (v === "" || v === null || v === undefined) { obj[headers[j]] = ""; continue; }
      if (v instanceof Date) {
        // Google Sheets auto-convert string tanggal → Date. Kembalikan ke string.
        obj[headers[j]] = Utilities.formatDate(v, Session.getScriptTimeZone(), headers[j] === "bulan" ? "yyyy-MM" : "yyyy-MM-dd");
        continue;
      }
      if (typeof v === "string" && (v.charAt(0) === "[" || v.charAt(0) === "{")) {
        try { obj[headers[j]] = JSON.parse(v); } catch (e) { obj[headers[j]] = v; }
      } else {
        obj[headers[j]] = v;
      }
    }
    arr.push(obj);
  }
  return arr;
}

function bacaScalar_(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return null;
  var v = values[1][1];
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "string" && !isNaN(Number(v))) return Number(v);
  return v;
}

function loadRev() {
  var rev = {};
  var props = PropertiesService.getScriptProperties();
  MODULES.forEach(function(m) {
    rev[m] = Number(props.getProperty("rev_" + m) || 0);
  });
  SCALAR_KEYS.forEach(function(k) {
    rev[k] = Number(props.getProperty("rev_" + k) || 0);
  });
  return rev;
}

/* Hash sederhana dari seluruh isi sheet (untuk deteksi perubahan) */
function hashKonten_(values) {
  var s = "";
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      s += String(values[i][j]) + "|";
    }
    s += "\n";
  }
  var h = 0;
  for (var k = 0; k < s.length; k++) {
    h = ((h << 5) - h + s.charCodeAt(k)) | 0;
  }
  return Math.abs(h);
}

function saveAll(data, rev) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var now = new Date().toISOString();
    var props = PropertiesService.getScriptProperties();
    var keys = Object.keys(data);
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      var sh = ssGetSheetByName_(key);
      if (!sh) {
        // Sheet belum ada — tulis langsung
        if (MODULES.indexOf(key) >= 0) tulisSheet_(key, data[key]);
        else if (SCALAR_KEYS.indexOf(key) >= 0) tulisScalar_(key, data[key]);
        var shNew = ssGetSheetByName_(key);
        if (shNew) {
          props.setProperty("hash_" + key, String(hashKonten_(shNew.getDataRange().getValues())));
          props.setProperty("rev_" + key, String(Date.now()));
        }
        continue;
      }
      var curHash = hashKonten_(sh.getDataRange().getValues());
      var lastHash = Number(props.getProperty("hash_" + key) || 0);
      var isi = sh.getDataRange().getValues();
      var kosong = isi.length < 2 || (isi.length === 1 && isi[0].join("").trim() === "");
      if (curHash === lastHash || lastHash === 0 || kosong) {
        // Sheet tidak diubah manual sejak terakhir ditulis server → aman timpa
        // lastHash===0: sheet baru / hash belum pernah diset → timpa juga
        // kosong: sheet cuma header → timpa juga (perbaiki race hash lama)
        if (MODULES.indexOf(key) >= 0) tulisSheet_(key, data[key]);
        else if (SCALAR_KEYS.indexOf(key) >= 0) tulisScalar_(key, data[key]);
        var sh2 = ssGetSheetByName_(key);
        if (sh2) {
          props.setProperty("hash_" + key, String(hashKonten_(sh2.getDataRange().getValues())));
          props.setProperty("rev_" + key, String(Date.now()));
        }
      }
      // else: sheet diubah manual (inject) → server hormati, jangan timpa
    }
  } finally {
    lock.releaseLock();
  }
}

function tulisSheet_(m, val) {
  var sh = getSheetByKey_(m);
  var cols = MODULE_COLUMNS[m] || ["key", "value"];
  // Hapus isi lama, tulis ulang header + data
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  else if (last === 1) sh.getRange(1, 1, 1, cols.length).clearContent();
  sh.appendRow(cols);
  if (m === "profil") {
    // profil = object → tulis key|value
    if (val && typeof val === "object" && !Array.isArray(val)) {
      var rows = Object.keys(val).map(function(k) {
        var v = val[k];
        if (v === null || v === undefined) return [k, ""];
        if (typeof v === "object") return [k, JSON.stringify(v)];
        return [k, String(v)];
      });
      if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    sh.setFrozenRows(1);
    return;
  }
  if (!Array.isArray(val)) return;
  var rows = [];
  val.forEach(function(item) {
    if (!item || typeof item !== "object") return;
    var row = cols.map(function(c) {
      var v = item[c];
      if (v === null || v === undefined) return "";
      if (typeof v === "object") return JSON.stringify(v);
      return String(v);
    });
    rows.push(row);
  });
  // Paksa kolom "bulan" jadi teks SEBELUM setValues — cegah Google Sheets auto-convert "2026-03" → Date
  var idxBulan = cols.indexOf("bulan");
  if (idxBulan >= 0 && rows.length) {
    sh.getRange(2, idxBulan + 1, rows.length, 1).setNumberFormat("@");
  }
  if (rows.length) sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
  sh.setFrozenRows(1);
}

function tulisScalar_(k, val) {
  var sh = getSheetByKey_(k);
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  else if (last === 1) sh.getRange(1, 1, 1, 2).clearContent();
  sh.appendRow(["key", "value"]);
  sh.appendRow([k, val === null || val === undefined ? "" : String(val)]);
  sh.setFrozenRows(1);
}

/* Migrasi: konversi blob JSON lama (sheet SIMPADU_RT) → sheet per kategori */
function migrasiKeTabel_() {
  var data = {};
  var legacy = getSheet_();
  var values = legacy.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim();
    if (!key) continue;
    try { data[key] = JSON.parse(values[i][1]); } catch (err) { data[key] = null; }
  }
  var keys = Object.keys(data);
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  for (var j = 0; j < keys.length; j++) {
    var k = keys[j];
    if (MODULES.indexOf(k) >= 0) {
      tulisSheet_(k, data[k]);
      var sh = ssGetSheetByName_(k);
      if (sh) {
        props.setProperty("hash_" + k, String(hashKonten_(sh.getDataRange().getValues())));
        props.setProperty("rev_" + k, String(now));
      }
    } else if (SCALAR_KEYS.indexOf(k) >= 0) {
      tulisScalar_(k, data[k]);
      var sh2 = ssGetSheetByName_(k);
      if (sh2) {
        props.setProperty("hash_" + k, String(hashKonten_(sh2.getDataRange().getValues())));
        props.setProperty("rev_" + k, String(now));
      }
    }
  }
}

/* ====== Utilitas admin (jalankan manual dari editor GAS) ====== */

function resetData() {
  // Hapus semua data di sheet — HATI-HATI, tidak bisa dibatalkan!
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  MODULES.forEach(function(m) {
    var sh = ss.getSheetByName(namaSheet_(m));
    if (sh) {
      var last = sh.getLastRow();
      if (last > 0) sh.deleteRows(1, last);
    }
  });
  SCALAR_KEYS.forEach(function(k) {
    var sh = ss.getSheetByName(namaSheet_(k));
    if (sh) {
      var last = sh.getLastRow();
      if (last > 0) sh.deleteRows(1, last);
    }
  });
  var legacy = ss.getSheetByName(SHEET_NAME);
  if (legacy) {
    var last = legacy.getLastRow();
    if (last > 1) legacy.deleteRows(2, last - 1);
  }
  return "Data direset";
}

/* Reset semua hash_* di ScriptProperties agar saveAll bisa menulis ulang.
 * Jalankan dari editor GAS: pulihkanData_()  */
function pulihkanData_() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  var count = 0;
  keys.forEach(function(k) {
    if (k.indexOf("hash_") === 0) {
      props.deleteProperty(k);
      count++;
    }
  });
  return "Reset " + count + " hash propertie(s). Sekarang frontend bisa push data.";
}

function testConnection() {
  return JSON.stringify({ ok: true, data: loadAll() });
}

/* ====== BACKUP SPREADSHEET (tampilan seperti database / Laragon) ======
 * Membuat file Google Spreadsheet terpisah berisi seluruh data,
 * satu tab per kategori (warga, transaksi, pengumuman, dll)
 * dengan judul + kolom sesuai field — mudah dibaca & dicari.
 *
 * Jalankan manual: buatBackupSpreadsheet()
 * Atau pasang trigger harian: pasangTriggerBackup()
 */

var BACKUP_PREFIX = "SIMPADU_RT_Backup";

var MODULE_LABELS = {
  profil: "Profil RT",
  users: "Pengguna & Hak Akses",
  warga: "Data Warga (Kepala Keluarga)",
  pengurus: "Struktur Pengurus",
  periodeList: "Arsip Periode Kepengurusan",
  pengumuman: "Pengumuman & Informasi",
  agenda: "Agenda & Kegiatan",
  arisan: "Jadwal Arisan",
  danaSosial: "Dana Sosial",
  transaksi: "Transaksi Kas",
  iuran: "Iuran Warga",
  pengaduan: "Pengaduan Warga",
  berita: "Berita & Artikel",
  galeri: "Galeri Foto",
  arsip: "Arsip Berita Tahunan",
  notifikasi: "Notifikasi"
};

function buatBackupSpreadsheet() {
  var data = loadAll();
  var tz = Session.getScriptTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd_HHmm");
  var backup = SpreadsheetApp.create(BACKUP_PREFIX + "_" + stamp);
  tulisBackup_(backup, data);
  return backup.getUrl();
}

function perbaruiBackupSpreadsheet() {
  // Perbarui file backup terbaru (tanpa bikin file baru)
  var files = DriveApp.searchFiles('title contains "' + BACKUP_PREFIX + '" and mimeType="' + MimeType.GOOGLE_SHEETS + '"');
  var backup = null;
  while (files.hasNext()) backup = SpreadsheetApp.open(files.next());
  if (!backup) return buatBackupSpreadsheet();
  tulisBackup_(backup, loadAll());
  return backup.getUrl();
}

function tulisBackup_(backup, data) {
  var sheets = backup.getSheets();
  var info = sheets[0];
  info.setName("INFO");
  info.clear();
  for (var i = 1; i < sheets.length; i++) backup.deleteSheet(sheets[i]);

  info.appendRow(["SIMPADU RT — BACKUP DATA"]);
  info.appendRow(["Dibuat", new Date().toISOString()]);
  info.appendRow(["Jumlah Kategori", Object.keys(MODULE_LABELS).length]);
  info.appendRow([]);
  info.appendRow(["Kategori", "Judul", "Jumlah Data"]);

  var mods = Object.keys(MODULE_LABELS);
  for (var j = 0; j < mods.length; j++) {
    var m = mods[j];
    var val = data[m];
    var count = Array.isArray(val) ? val.length : (val && typeof val === "object" ? 1 : 1);
    info.appendRow([m, MODULE_LABELS[m], count]);
    if (Array.isArray(val) && val.length) tulisArraySheet_(backup, m, MODULE_LABELS[m], val);
    else if (val && typeof val === "object") tulisObjectSheet_(backup, m, MODULE_LABELS[m], val);
    else tulisScalarSheet_(backup, m, MODULE_LABELS[m], val);
  }

  info.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  info.getRange(5, 1, 1, 3).setFontWeight("bold").setBackground("#e8f0e4");
  info.setFrozenRows(5);
  try { info.autoResizeColumns(1, 3); } catch (e) {}
}

function namaSheet_(m) {
  return String(m).replace(/[\\\/\?\*\[\]:]/g, "_").substring(0, 100) || "Sheet";
}

function tulisArraySheet_(backup, m, label, arr) {
  var sh = backup.insertSheet(namaSheet_(m));
  sh.appendRow([label]);
  sh.appendRow([]);
  var headers = [];
  arr.forEach(function(item) {
    if (item && typeof item === "object") Object.keys(item).forEach(function(k) { if (headers.indexOf(k) < 0) headers.push(k); });
  });
  if (!headers.length) headers = ["value"];
  sh.appendRow(headers);
  var rows = arr.map(function(item) {
    return headers.map(function(h) {
      var v = item ? item[h] : "";
      if (v === null || v === undefined) return "";
      if (typeof v === "object") return JSON.stringify(v);
      return String(v);
    });
  });
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  sh.getRange(3, 1, 1, headers.length).setFontWeight("bold").setBackground("#e8f0e4");
  sh.setFrozenRows(3);
  try { sh.autoResizeColumns(1, headers.length); } catch (e) {}
}

function tulisObjectSheet_(backup, m, label, obj) {
  var sh = backup.insertSheet(namaSheet_(m));
  sh.appendRow([label]);
  sh.appendRow([]);
  sh.appendRow(["Field", "Nilai"]);
  var rows = Object.keys(obj).map(function(k) {
    var v = obj[k];
    if (v === null || v === undefined) return [k, ""];
    if (typeof v === "object") return [k, JSON.stringify(v)];
    return [k, String(v)];
  });
  if (rows.length) sh.getRange(4, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  sh.getRange(3, 1, 1, 2).setFontWeight("bold").setBackground("#e8f0e4");
  sh.setFrozenRows(3);
  try { sh.autoResizeColumns(1, 2); } catch (e) {}
}

function tulisScalarSheet_(backup, m, label, val) {
  var sh = backup.insertSheet(namaSheet_(m));
  sh.appendRow([label]);
  sh.appendRow([]);
  sh.appendRow(["Nilai"]);
  sh.appendRow([val === null || val === undefined ? "" : String(val)]);
  sh.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  sh.getRange(3, 1).setFontWeight("bold").setBackground("#e8f0e4");
  sh.setFrozenRows(3);
  try { sh.autoResizeColumns(1, 1); } catch (e) {}
}

function pasangTriggerBackup() {
  // Backup otomatis setiap hari jam 02.00
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === "backupOtomatis") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("backupOtomatis").timeBased().everyDays(1).atHour(2).create();
  return "Trigger backup harian dipasang (setiap hari 02.00)";
}

function backupOtomatis() {
  return buatBackupSpreadsheet();
}