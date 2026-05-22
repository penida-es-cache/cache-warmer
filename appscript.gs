const DEFAULT_HEADER = ["run_id","started_at","finished_at","country","url","status","cf_cache","ls_cache","cf_ray","response_ms","error","message"];
const TZ = 'Asia/Makassar';
const SHEETS_TO_DELETE = 5;   // jumlah sheet terlama yang dihapus per cleanup
const MAX_CLEANUP_RETRIES = 3; // maks berapa kali retry cleanup jika masih kena limit

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return respondJSON_({ ok: false, error: 'No payload' });
    }
    const data = JSON.parse(e.postData.contents);

    // Ambil rows mentah
    let rows = data && data.rows;
    if (!rows) return respondJSON_({ ok: false, error: 'No rows' });

    // NORMALISASI:
    // - pastikan 2D array (array of arrays)
    // - ubah null/undefined jadi '' supaya setValues tidak error
    // - samakan panjang kolom (pad) agar "persegi panjang"
    rows = normalizeRows_(rows);

    // Tentukan nama sheet (opsional dari client, kalau tidak ada pakai timestamp WITA)
    const sheetName = (data.sheetName && String(data.sheetName).trim()) || generateSheetName_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ===== CLEANUP: hapus sheet terlama dulu sebelum insert =====
    deleteOldestSheets_(ss, SHEETS_TO_DELETE);

    // Buat/ambil sheet tujuan
    const sh = ensureSheet_(sheetName);

    // header sekali saja
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, DEFAULT_HEADER.length).setValues([DEFAULT_HEADER]);
    }

    // tulis rows dengan retry (jika masih kena limit, hapus lagi)
    let inserted = false;
    let lastErr = null;
    for (var attempt = 0; attempt < MAX_CLEANUP_RETRIES; attempt++) {
      try {
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        inserted = true;
        break;
      } catch (writeErr) {
        lastErr = writeErr;
        var errMsg = String(writeErr);
        // Hanya retry jika error karena cell limit
        if (errMsg.indexOf('number of cells') === -1 && errMsg.indexOf('limit') === -1) {
          throw writeErr; // error lain, langsung lempar
        }
        Logger.log('Cell limit hit, cleanup attempt ' + (attempt + 1) + '/' + MAX_CLEANUP_RETRIES);
        deleteOldestSheets_(ss, SHEETS_TO_DELETE);
      }
    }

    if (!inserted) {
      return respondJSON_({ ok: false, error: 'Still over cell limit after ' + MAX_CLEANUP_RETRIES + ' cleanups: ' + String(lastErr) });
    }

    return respondJSON_({ ok: true, inserted: rows.length, sheet: sh.getName() });
  } catch (err) {
    return respondJSON_({ ok: false, error: String(err) });
  }
}

/**
 * Hapus N sheet tab paling lama berdasarkan tanggal di nama sheet.
 * Nama sheet yang diharapkan: YYYY-MM-DD_HH-mm-ss_WITA
 * Sheet yang tidak bisa di-parse tanggalnya akan DILEWATI (tidak dihapus).
 * Sheet pertama (index 0) TIDAK akan dihapus jika itu satu-satunya yang tersisa
 * (Google Sheets wajib punya minimal 1 sheet).
 */
function deleteOldestSheets_(ss, count) {
  var allSheets = ss.getSheets();

  // Hanya proses sheet yang namanya mengandung tanggal (format: YYYY-MM-DD_HH-mm-ss_WITA)
  var datedSheets = [];
  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    var parsed = parseSheetDate_(sheet.getName());
    if (parsed) {
      datedSheets.push({ sheet: sheet, date: parsed, name: sheet.getName() });
    }
  }

  // Urutkan dari tanggal terlama ke terbaru
  datedSheets.sort(function(a, b) {
    return a.date.getTime() - b.date.getTime();
  });

  // Pastikan kita tidak menghapus SEMUA sheet (Google Sheets butuh min 1)
  var totalSheets = ss.getSheets().length;
  var toDelete = Math.min(count, datedSheets.length, totalSheets - 1);

  var deleted = [];
  for (var j = 0; j < toDelete; j++) {
    try {
      Logger.log('Deleting oldest sheet: ' + datedSheets[j].name);
      ss.deleteSheet(datedSheets[j].sheet);
      deleted.push(datedSheets[j].name);
    } catch (delErr) {
      Logger.log('Failed to delete sheet "' + datedSheets[j].name + '": ' + String(delErr));
    }
  }

  if (deleted.length > 0) {
    // Flush perubahan supaya batas sel ter-update
    SpreadsheetApp.flush();
    Logger.log('Deleted ' + deleted.length + ' oldest sheet(s): ' + deleted.join(', '));
  }

  return deleted;
}

/**
 * Parse nama sheet ke Date object.
 * Format yang diharapkan: "YYYY-MM-DD_HH-mm-ss_WITA"
 * Return null jika format tidak cocok.
 */
function parseSheetDate_(name) {
  // Regex untuk format: 2025-05-22_15-30-45_WITA
  var match = String(name).match(/^(\d{4})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})_WITA/);
  if (!match) return null;

  var year  = parseInt(match[1], 10);
  var month = parseInt(match[2], 10) - 1; // JS bulan 0-indexed
  var day   = parseInt(match[3], 10);
  var hour  = parseInt(match[4], 10);
  var min   = parseInt(match[5], 10);
  var sec   = parseInt(match[6], 10);

  return new Date(year, month, day, hour, min, sec);
}

// Ubah input apapun jadi matrix yang rapi untuk setValues
function normalizeRows_(rows) {
  // Jika kirim satu baris 1D (mis. [1,2,3]) bungkus jadi [[1,2,3]]
  if (Array.isArray(rows) && !Array.isArray(rows[0])) rows = [rows];

  // Jadikan tiap baris array; ubah null/undefined → '' (string kosong)
  let normalized = rows.map(r => {
    if (!Array.isArray(r)) r = [r];
    return r.map(v => (v === null || v === undefined) ? '' : v);
  });

  // Samakan panjang kolom (pad) → persegi panjang
  const maxCols = Math.max(DEFAULT_HEADER.length, ...normalized.map(r => r.length));
  normalized = normalized.map(r => (r.length < maxCols)
    ? r.concat(Array(maxCols - r.length).fill(''))
    : r.slice(0, maxCols)
  );

  return normalized;
}

function generateSheetName_() {
  var now = new Date();
  var stamp = Utilities.formatDate(now, TZ, "yyyy-MM-dd_HH-mm-ss") + "_WITA";
  return sanitizeSheetName_(stamp);
}

function ensureSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clean = sanitizeSheetName_(name);
  var base = clean;
  var i = 1;
  var existing = ss.getSheetByName(clean);
  while (existing) {
    i += 1;
    clean = truncateSheetName_(base + "-" + i);
    existing = ss.getSheetByName(clean);
  }
  return existing || ss.insertSheet(clean);
}

function sanitizeSheetName_(name) {
  var clean = String(name).replace(/[:\\\/\?\*\[\]]/g, " ");
  return truncateSheetName_(clean.trim() || "untitled");
}

function truncateSheetName_(name) {
  return name.length > 100 ? name.slice(0, 100) : name;
}

function respondJSON_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
