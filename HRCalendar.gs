/**
 * =========================================================
 * FILE: HRCalendar.gs
 * FUNGSI: CRUD modul HR Calendar (sheet "HR_Calendar").
 *         Lampiran disimpan ke folder "HR_Calendar_Lampiran"
 *         di Google Drive.
 *
 * Struktur kolom sheet (17 kolom):
 *   A: ID | B: Tanggal | C: Nama Kegiatan | D: Departemen
 *   E: PIC | F: Lokasi | G: Status | H: Keterangan
 *   I: Jam Mulai | J: Jam Selesai | K: Jenis Kegiatan
 *   L: Prioritas | M: Catatan | N: Dibuat Oleh
 *   O: Lampiran | P: Tanggal Dibuat | Q: Waktu Update
 * =========================================================
 */

const SHEET_HR_CALENDAR    = "HR_Calendar";
const FOLDER_CALENDAR      = "HR_Calendar_Lampiran";
const HR_CALENDAR_HEADERS  = [
  "ID", "Tanggal", "Nama Kegiatan", "Departemen", "PIC",
  "Lokasi", "Status", "Keterangan", "Jam Mulai", "Jam Selesai",
  "Jenis Kegiatan", "Prioritas", "Catatan", "Dibuat Oleh",
  "Lampiran", "Tanggal Dibuat", "Waktu Update"
];

// ── Init Sheet ────────────────────────────────────────────
/**
 * Memastikan sheet ada dengan header 17 kolom.
 * Jika header lama (kolom kurang), diperbaiki otomatis.
 */
function initSheetHRCalendar() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_HR_CALENDAR);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HR_CALENDAR);
    sheet.appendRow(HR_CALENDAR_HEADERS);
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    return sheet;
  }

  const curWidth   = Math.max(sheet.getLastColumn(), 1);
  const curHeaders = sheet.getRange(1, 1, 1, Math.min(curWidth, HR_CALENDAR_HEADERS.length)).getValues()[0];
  const needsFix   = curWidth < HR_CALENDAR_HEADERS.length
    || HR_CALENDAR_HEADERS.some((h, idx) => curHeaders[idx] !== h);

  if (needsFix) {
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length).setValues([HR_CALENDAR_HEADERS]);
    sheet.getRange(1, 1, 1, HR_CALENDAR_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Read ──────────────────────────────────────────────────
function getHRCalendarEvents() {
  try {
    const sheet = initSheetHRCalendar();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const tz = Session.getScriptTimeZone();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const tgl = r[1] instanceof Date
                  ? Utilities.formatDate(r[1], tz, "yyyy-MM-dd")
                  : String(r[1] || "").trim();
      result.push({
        id:             String(r[0]),
        tanggal:        tgl,
        namaKegiatan:   String(r[2]  || ""),
        dept:           String(r[3]  || ""),
        pic:            String(r[4]  || ""),
        lokasi:         String(r[5]  || ""),
        status:         String(r[6]  || "Planning"),
        keterangan:     String(r[7]  || ""),
        jamMulai:       String(r[8]  || ""),
        jamSelesai:     String(r[9]  || ""),
        jenisKegiatan:  String(r[10] || ""),
        prioritas:      String(r[11] || ""),
        catatan:        String(r[12] || ""),
        dibuatOleh:     String(r[13] || ""),
        lampiran:       String(r[14] || ""),
        tanggalDibuat:  r[15] ? String(r[15]) : "",
        waktuUpdate:    r[16] ? String(r[16]) : ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("getHRCalendarEvents gagal: " + e.message);
  }
}

// ── Create / Update ───────────────────────────────────────
function simpanHRCalendarEvent(ev) {
  if (!ev || typeof ev !== "object") throw new Error("Data kegiatan tidak valid.");
  if (!ev.tanggal || !ev.namaKegiatan || !ev.dept || !ev.pic) {
    throw new Error("Tanggal, Nama Kegiatan, Departemen, dan PIC wajib diisi.");
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Sistem sedang sibuk, coba lagi."); }

  try {
    const sheet    = initSheetHRCalendar();
    const rows     = sheet.getDataRange().getValues();
    const evIdStr  = String(ev.id || "").trim();
    const rowValues = [
      String(ev.tanggal        || ""),
      String(ev.namaKegiatan   || ""),
      String(ev.dept           || ""),
      String(ev.pic            || ""),
      String(ev.lokasi         || ""),
      String(ev.status         || "Planning"),
      String(ev.keterangan     || ""),
      String(ev.jamMulai       || ""),
      String(ev.jamSelesai     || ""),
      String(ev.jenisKegiatan  || ""),
      String(ev.prioritas      || ""),
      String(ev.catatan        || ""),
      String(ev.dibuatOleh     || ""),
      String(ev.lampiran       || ""),
      String(ev.tanggalDibuat  || ""),
      String(ev.waktuUpdate    || "")
    ];

    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === evIdStr) {
        sheet.getRange(i + 1, 2, 1, rowValues.length).setValues([rowValues]);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sheet.appendRow([evIdStr].concat(rowValues));
    }
    return "Kegiatan HR Calendar berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanHRCalendarEvent gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusHRCalendarEvent(id) {
  if (!id) throw new Error("ID kegiatan tidak boleh kosong.");

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Sistem sedang sibuk, coba lagi."); }

  try {
    const sheet = initSheetHRCalendar();
    const rows  = sheet.getDataRange().getValues();
    const idStr = String(id).trim();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === idStr) {
        sheet.deleteRow(i + 1);
        return "Kegiatan HR Calendar berhasil dihapus.";
      }
    }
    throw new Error("ID kegiatan tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusHRCalendarEvent gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Upload Lampiran ke Drive ──────────────────────────────
/**
 * Upload lampiran kegiatan ke Drive.
 * Mengembalikan URL publik file yang diunggah.
 */
function uploadFileLampiranKalender(base64Data, fileName, mimeType) {
  if (!base64Data) throw new Error("Data file kosong.");
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_CALENDAR);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_CALENDAR);
    const bytes   = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName || "lampiran");
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    throw new Error("Gagal mengunggah lampiran kalender: " + e.message);
  }
}