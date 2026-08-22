/**
 * =========================================================
 * FILE: KaryawanSakitTerlambat.gs
 * FUNGSI: CRUD data Karyawan Sakit (sheet "Data_Karyawan_Sakit")
 *         dan Karyawan Terlambat (sheet "Data_Karyawan_Terlambat").
 *
 * Struktur kolom — Data_Karyawan_Sakit (11 kolom):
 *   A: ID | B: NIK | C: Nama | D: Departemen | E: Jabatan
 *   F: Tanggal Mulai | G: Tanggal Selesai | H: Total Hari
 *   I: Jenis | J: Keterangan | K: Waktu Update
 *
 * Struktur kolom — Data_Karyawan_Terlambat (12 kolom):
 *   A: ID | B: NIK | C: Nama | D: Departemen | E: Jabatan
 *   F: Tanggal | G: Jam Masuk | H: Jam Standar | I: Jumlah Menit
 *   J: Alasan | K: Keterangan | L: Waktu Update
 * =========================================================
 */

const SHEET_KARYAWAN_SAKIT     = "Data_Karyawan_Sakit";
const SHEET_KARYAWAN_TERLAMBAT = "Data_Karyawan_Terlambat";

// ── Helpers ───────────────────────────────────────────────
function _formatTanggal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const m = str.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  return m ? m[0] : str;
}

function _formatJam(val, defaultVal) {
  if (!val) return defaultVal;
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  const str = String(val).trim();
  if (/^([01]\d|2[0-3]):[0-5]\d$/.test(str)) return str;
  const m = str.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
  return m ? m[1].padStart(2, "0") + ":" + m[2] : (str || defaultVal);
}

// =========================================================
// MODUL: KARYAWAN SAKIT
// =========================================================

function initSheetKaryawanSakit() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN_SAKIT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN_SAKIT);
    sheet.appendRow([
      "ID", "NIK", "Nama", "Departemen", "Jabatan",
      "Tanggal Mulai", "Tanggal Selesai", "Total Hari",
      "Jenis", "Keterangan", "Waktu Update"
    ]);
    sheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getKaryawanSakit() {
  try {
    const sheet = initSheetKaryawanSakit();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      result.push({
        id:            String(r[0]),
        nik:           String(r[1] || ""),
        nama:          String(r[2] || "-"),
        dept:          String(r[3] || "-"),
        jabatan:       String(r[4] || "-"),
        tanggalMulai:  _formatTanggal(r[5]),
        tanggalSelesai:_formatTanggal(r[6]),
        totalHari:     parseInt(r[7]) || 1,
        jenis:         String(r[8] || "Surat Dokter"),
        keterangan:    String(r[9] || "")
      });
    }
    return result;
  } catch (e) {
    throw new Error("getKaryawanSakit gagal: " + e.message);
  }
}

function simpanKaryawanSakit(item) {
  try {
    const sheet = initSheetKaryawanSakit();
    const rows  = sheet.getDataRange().getValues();
    const waktu = new Date();
    const values = [
      item.nik            || "",
      item.nama           || "-",
      item.dept           || "-",
      item.jabatan        || "-",
      item.tanggalMulai   || "",
      item.tanggalSelesai || "",
      item.totalHari      || 1,
      item.jenis          || "Surat Dokter",
      item.keterangan     || "",
      waktu
    ];
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(item.id)) {
        sheet.getRange(i + 1, 2, 1, values.length).setValues([values]);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sheet.appendRow([item.id || ("KS_" + Date.now())].concat(values));
    }
    return "Data Karyawan Sakit berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanKaryawanSakit gagal: " + e.message);
  }
}

function hapusKaryawanSakit(id) {
  try {
    const sheet = initSheetKaryawanSakit();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data Karyawan Sakit berhasil dihapus.";
      }
    }
    throw new Error("ID tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusKaryawanSakit gagal: " + e.message);
  }
}

// =========================================================
// MODUL: KARYAWAN TERLAMBAT
// =========================================================

function initSheetKaryawanTerlambat() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN_TERLAMBAT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN_TERLAMBAT);
    sheet.appendRow([
      "ID", "NIK", "Nama", "Departemen", "Jabatan",
      "Tanggal", "Jam Masuk", "Jam Standar", "Jumlah Menit",
      "Alasan", "Keterangan", "Waktu Update"
    ]);
    sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getKaryawanTerlambat() {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      result.push({
        id:          String(r[0]),
        nik:         String(r[1] || ""),
        nama:        String(r[2] || "-"),
        dept:        String(r[3] || "-"),
        jabatan:     String(r[4] || "-"),
        tanggal:     _formatTanggal(r[5]),
        jamMasuk:    _formatJam(r[6], "08:30"),
        jamStandar:  _formatJam(r[7], "08:00"),
        jumlahMenit: parseInt(r[8]) || 0,
        alasan:      String(r[9] || ""),
        keterangan:  String(r[10] || "")
      });
    }
    return result;
  } catch (e) {
    throw new Error("getKaryawanTerlambat gagal: " + e.message);
  }
}

function simpanKaryawanTerlambat(item) {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const rows  = sheet.getDataRange().getValues();
    const waktu = new Date();
    const values = [
      item.nik         || "",
      item.nama        || "-",
      item.dept        || "-",
      item.jabatan     || "-",
      item.tanggal     || "",
      item.jamMasuk    || "08:30",
      item.jamStandar  || "08:00",
      item.jumlahMenit || 0,
      item.alasan      || "",
      item.keterangan  || "",
      waktu
    ];
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(item.id)) {
        sheet.getRange(i + 1, 2, 1, values.length).setValues([values]);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sheet.appendRow([item.id || ("KT_" + Date.now())].concat(values));
    }
    return "Data Karyawan Terlambat berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanKaryawanTerlambat gagal: " + e.message);
  }
}

function hapusKaryawanTerlambat(id) {
  try {
    const sheet = initSheetKaryawanTerlambat();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data Karyawan Terlambat berhasil dihapus.";
      }
    }
    throw new Error("ID tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusKaryawanTerlambat gagal: " + e.message);
  }
}