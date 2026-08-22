/**
 * =========================================================
 * FILE: Code.gs
 * FUNGSI: Entry point Web App (doGet) + fungsi koneksi DB.
 *
 * ATURAN PENTING — SATU SCOPE GLOBAL:
 * Semua file .gs dalam satu project Apps Script berbagi SATU
 * scope global. Jangan pernah mendefinisikan dua fungsi dengan
 * nama yang sama di file berbeda — yang "menang" ditentukan
 * urutan loading yang TIDAK dijamin stabil oleh Apps Script.
 *
 * File ini HANYA berisi:
 *   1. Konstanta SPREADSHEET_ID & getDB() → dipakai SEMUA file .gs lain.
 *   2. doGet()                            → entry point Web App.
 * Semua CRUD ada di file modular masing-masing modul.
 * =========================================================
 */

// ── Konfigurasi ──────────────────────────────────────────
// Kosongkan string di bawah jika script menempel langsung pada
// spreadsheet (bound script). Isi dengan ID spreadsheet tujuan
// jika script berdiri sendiri (standalone script) atau Anda
// ingin menulis ke spreadsheet LAIN selain yang aktif.
const SPREADSHEET_ID = "";

// ── Koneksi Database ─────────────────────────────────────
/**
 * Mengembalikan objek Spreadsheet yang menjadi database utama.
 * Dipakai oleh SEMUA file .gs lain — jangan duplikasi fungsi ini.
 */
function getDB() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.length > 20) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── Entry Point Web App ───────────────────────────────────
/**
 * Dipanggil otomatis oleh Apps Script saat URL Web App diakses.
 * Menyajikan Index.html sebagai antarmuka utama.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('HR LP — Labor Cost & Productivity')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}