/**
 * =========================================================
 * FILE: DailyReport.gs
 * FUNGSI: CRUD Daily HR Report (sheet "Data_Daily_Report").
 *
 * Aturan data: SATU baris per tanggal (upsert by tanggal).
 * Saat menyimpan, jika tanggal sudah ada → overwrite konten.
 *
 * Struktur kolom sheet (3 kolom):
 *   A: Tanggal | B: Konten Laporan | C: Waktu Update
 * =========================================================
 */

const SHEET_DAILY_REPORT = "Data_Daily_Report";

// ── Init Sheet ────────────────────────────────────────────
function initSheetDailyReport() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_DAILY_REPORT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DAILY_REPORT);
    sheet.appendRow(["Tanggal", "Konten Laporan", "Waktu Update"]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 500); // Lebarkan kolom konten agar mudah dibaca di Sheet
  }
  return sheet;
}

// ── Helpers ───────────────────────────────────────────────
function _formatTglDR(val) {
  if (!val) return "";
  return val instanceof Date
    ? Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(val);
}

// ── Read ──────────────────────────────────────────────────
function getDailyReports() {
  try {
    const sheet = initSheetDailyReport();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      result.push({
        tanggal:     _formatTglDR(r[0]),
        konten:      r[1] || "",
        waktuUpdate: r[2] ? String(r[2]) : ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("getDailyReports gagal: " + e.message);
  }
}

// ── Create / Update (upsert by tanggal) ───────────────────
function simpanDailyReport(report) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk, coba lagi."); }

  try {
    const sheet = initSheetDailyReport();
    const rows  = sheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < rows.length; i++) {
      if (_formatTglDR(rows[i][0]) === report.tanggal) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[
          report.konten      || "",
          report.waktuUpdate || new Date()
        ]]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      sheet.appendRow([
        report.tanggal,
        report.konten      || "",
        report.waktuUpdate || new Date()
      ]);
    }
    return "Daily HR Report berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanDailyReport gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusDailyReport(tanggal) {
  try {
    const sheet = initSheetDailyReport();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (_formatTglDR(rows[i][0]) === tanggal) {
        sheet.deleteRow(i + 1);
        return "Daily HR Report berhasil dihapus.";
      }
    }
    throw new Error("Laporan tanggal " + tanggal + " tidak ditemukan.");
  } catch (e) {
    throw new Error("hapusDailyReport gagal: " + e.message);
  }
}