/**
 * =========================================================
 * FILE: Karyawan.gs
 * FUNGSI: CRUD Master Data Karyawan (sheet "Data_Karyawan").
 *
 * Struktur kolom sheet (9 kolom):
 *   A: ID | B: NIK | C: Nama | D: Departemen | E: Jabatan
 *   F: Target Jam | G: Mandatory | H: Status | I: Waktu Update
 * =========================================================
 */

const SHEET_KARYAWAN = "Data_Karyawan";
const KARYAWAN_HEADERS = [
  "ID", "NIK", "Nama", "Departemen", "Jabatan",
  "Target Jam", "Mandatory", "Status", "Waktu Update"
];

// ── Init Sheet ────────────────────────────────────────────
function initSheetKaryawan() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_KARYAWAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN);
    sheet.appendRow(KARYAWAN_HEADERS);
    sheet.getRange(1, 1, 1, KARYAWAN_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Read ──────────────────────────────────────────────────
/**
 * Mengembalikan array semua data karyawan ke frontend.
 */
function getKaryawan() {
  try {
    const sheet = initSheetKaryawan();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[1]) continue; // skip baris tanpa NIK
      result.push({
        id:        r[0],
        nik:       String(r[1]),
        nama:      r[2]  || "",
        dept:      r[3]  || "",
        jabatan:   r[4]  || "",
        targetJam: parseFloat(r[5]) || 0,
        mandatory: r[6]  || "Belum",
        status:    r[7]  || "Aktif"
      });
    }
    return result;
  } catch (e) {
    throw new Error("getKaryawan gagal: " + e.message);
  }
}

// ── Create / Update ───────────────────────────────────────
/**
 * Menyimpan (tambah baru) atau memperbarui (edit) satu data karyawan.
 * Kunci pencarian: karyawan.id (bukan NIK — NIK bisa berganti).
 */
function simpanKaryawan(karyawan) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk, coba lagi."); }

  try {
    const sheet      = initSheetKaryawan();
    const rows       = sheet.getDataRange().getValues();
    const waktuUpdate = new Date();
    let updated      = false;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(karyawan.id)) {
        sheet.getRange(i + 1, 2, 1, 8).setValues([[
          karyawan.nik,
          karyawan.nama      || "",
          karyawan.dept      || "",
          karyawan.jabatan   || "",
          karyawan.targetJam || 0,
          karyawan.mandatory || "Belum",
          karyawan.status    || "Aktif",
          waktuUpdate
        ]]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      sheet.appendRow([
        karyawan.id,
        karyawan.nik,
        karyawan.nama      || "",
        karyawan.dept      || "",
        karyawan.jabatan   || "",
        karyawan.targetJam || 0,
        karyawan.mandatory || "Belum",
        karyawan.status    || "Aktif",
        waktuUpdate
      ]);
    }

    return "Data karyawan berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanKaryawan gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Delete ────────────────────────────────────────────────
/**
 * Menghapus satu baris karyawan berdasarkan ID.
 */
function hapusKaryawan(id) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk, coba lagi."); }

  try {
    const sheet = initSheetKaryawan();
    const rows  = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data karyawan berhasil dihapus.";
      }
    }
    throw new Error("ID karyawan tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusKaryawan gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}