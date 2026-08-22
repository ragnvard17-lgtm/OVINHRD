/**
 * =========================================================
 * FILE: Training.gs
 * FUNGSI: CRUD riwayat pelatihan (sheet "Data_Training").
 *         Mendukung:
 *         - simpanTraining()      → edit satu record
 *         - simpanBatchTraining() → batch input banyak peserta sekaligus
 *
 * Struktur kolom sheet (18 kolom):
 *   A: ID | B: NIK | C: Tanggal | D: Topik Pelatihan | E: Durasi (Jam)
 *   F: Waktu Input | G: Event ID | H: Kategori | I: Trainer Name
 *   J: Trainer ID | K: Trainer Dept | L: Trainer Position
 *   M: Trainer Type | N: Lokasi | O: Deskripsi | P: Materi
 *   Q: Sertifikat | R: Created At
 * =========================================================
 */

const SHEET_TRAINING = "Data_Training";
const TRAINING_HEADERS = [
  "ID", "NIK", "Tanggal", "Topik Pelatihan", "Durasi (Jam)", "Waktu Input",
  "Event ID", "Kategori", "Trainer Name", "Trainer ID", "Trainer Dept",
  "Trainer Position", "Trainer Type", "Lokasi", "Deskripsi", "Materi",
  "Sertifikat", "Created At"
];

// ── Init Sheet ────────────────────────────────────────────
/**
 * Memastikan sheet ada dengan header 18 kolom yang benar.
 * Jika sheet sudah ada dengan header lama (kolom lebih sedikit),
 * header diperbaiki otomatis TANPA menghapus data yang sudah ada.
 */
function initSheetTraining() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_TRAINING);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRAINING);
    sheet.appendRow(TRAINING_HEADERS);
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Perbaiki header jika kurang atau tidak sesuai
  const currentWidth = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, Math.min(currentWidth, TRAINING_HEADERS.length))
                              .getValues()[0];
  const needsFix = currentWidth < TRAINING_HEADERS.length
    || TRAINING_HEADERS.some((h, idx) => currentHeaders[idx] !== h);

  if (needsFix) {
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length).setValues([TRAINING_HEADERS]);
    sheet.getRange(1, 1, 1, TRAINING_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Helpers ───────────────────────────────────────────────
function formatTanggalTraining_(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val).trim();
}

function buildTrainingRow_(t, waktuInput) {
  return [
    t.id,
    t.nik,
    t.tanggal                          || "",
    t.topik                            || "",
    parseFloat(t.durasi)               || 0,
    waktuInput                         || new Date(),
    t.eventId                          || "",
    t.kategori                         || "",
    t.trainerName  || t.trainer        || "",
    t.trainerId                        || "",
    t.trainerDept                      || "",
    t.trainerPos                       || "",
    t.trainerType                      || "",
    t.lokasi                           || "",
    t.deskripsi                        || "",
    t.materi                           || "",
    t.sertifikat                       || "",
    t.createdAt                        || new Date().toISOString()
  ];
}

// ── Read ──────────────────────────────────────────────────
function getTraining() {
  try {
    const sheet = initSheetTraining();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const tz = Session.getScriptTimeZone();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      result.push({
        id:          r[0],
        nik:         String(r[1]),
        tanggal:     formatTanggalTraining_(r[2]),
        topik:       r[3]  || "",
        durasi:      parseFloat(r[4]) || 0,
        eventId:     r[6]  || "",
        kategori:    r[7]  || "",
        trainerName: r[8]  || "",
        trainer:     r[8]  || "",
        trainerId:   r[9]  || "",
        trainerDept: r[10] || "",
        trainerPos:  r[11] || "",
        trainerType: r[12] || "",
        lokasi:      r[13] || "",
        deskripsi:   r[14] || "",
        materi:      r[15] || "",
        sertifikat:  r[16] || "",
        createdAt:   r[17] || ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("getTraining gagal: " + e.message);
  }
}

// ── Create / Update (satu record) ─────────────────────────
/**
 * Menyimpan atau memperbarui SATU riwayat training.
 * Digunakan saat user mengedit record individual dari tabel.
 */
function simpanTraining(training) {
  try {
    const sheet = initSheetTraining();
    const rows  = sheet.getDataRange().getValues();
    const waktu = new Date();
    const row   = buildTrainingRow_(training, waktu);
    let updated = false;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(training.id)) {
        sheet.getRange(i + 1, 2, 1, row.length - 1).setValues([row.slice(1)]);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sheet.appendRow(row);
    }
    return "Riwayat pelatihan berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanTraining gagal: " + e.message);
  }
}

// ── Batch Create (banyak peserta sekaligus) ───────────────
/**
 * Menyimpan BANYAK riwayat training dalam SATU eksekusi (Batch Training Input).
 * Juga menyinkronkan perubahan status Mandatory Assignment (opsional).
 *
 * payload = {
 *   trainingItems: [ {id, nik, tanggal, topik, durasi, ...}, ... ],
 *   mandatoryUpdates: [ {id, nik, training, dueDate, status}, ... ]  // opsional
 * }
 */
function simpanBatchTraining(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk memproses batch training lain, coba lagi."); }

  try {
    // ── Validasi payload ──────────────────────────────────
    const items   = (payload && payload.trainingItems)
                    ? payload.trainingItems
                    : (Array.isArray(payload) ? payload : []);
    const mandaUpdates = (payload && payload.mandatoryUpdates) ? payload.mandatoryUpdates : [];

    if (!items.length) {
      throw new Error("Tidak ada data peserta (trainingItems kosong). Pilih minimal 1 peserta.");
    }
    items.forEach(function (t, idx) {
      if (!t || !t.id || !t.nik || !t.tanggal || !t.topik) {
        throw new Error("Data peserta ke-" + (idx + 1) + " tidak lengkap (id/nik/tanggal/topik wajib).");
      }
    });

    // ── 1. Tulis semua riwayat training ───────────────────
    const sheet   = initSheetTraining();
    const rows    = sheet.getDataRange().getValues();
    const waktu   = new Date();

    // Buat map id → nomor baris untuk update existing
    const idToRow = {};
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) idToRow[String(rows[i][0])] = i + 1;
    }

    const toAppend = [];
    let added = 0, updated = 0;

    items.forEach(function (t) {
      const row = buildTrainingRow_(t, waktu);
      if (idToRow[String(t.id)]) {
        sheet.getRange(idToRow[String(t.id)], 2, 1, row.length - 1).setValues([row.slice(1)]);
        updated++;
      } else {
        toAppend.push(row);
        added++;
      }
    });

    if (toAppend.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    }
    SpreadsheetApp.flush(); // commit sebelum lock dilepas

    // ── 2. Sinkron Mandatory Assignment (opsional) ────────
    // Dibungkus try/catch terpisah: riwayat training sudah tersimpan,
    // jika sinkron mandatory gagal → training tetap dianggap sukses.
    let mandaAdded = 0, mandaUpdated = 0, mandaSyncError = null;
    if (mandaUpdates.length > 0) {
      try {
        const maSheet = initSheetMandatoryAssignment();
        const maRows  = maSheet.getDataRange().getValues();
        const maIdMap = {};
        for (let i = 1; i < maRows.length; i++) {
          if (maRows[i][0]) maIdMap[String(maRows[i][0])] = i + 1;
        }
        const maToAppend = [];
        const maWaktu    = new Date();

        mandaUpdates.forEach(function (a) {
          const rowVal = [a.nik, a.training, a.dueDate, a.status || "Belum", maWaktu];
          if (maIdMap[String(a.id)]) {
            maSheet.getRange(maIdMap[String(a.id)], 2, 1, rowVal.length).setValues([rowVal]);
            mandaUpdated++;
          } else {
            maToAppend.push([a.id].concat(rowVal));
            mandaAdded++;
          }
        });

        if (maToAppend.length > 0) {
          const maStart = maSheet.getLastRow() + 1;
          maSheet.getRange(maStart, 1, maToAppend.length, maToAppend[0].length).setValues(maToAppend);
        }
        SpreadsheetApp.flush();
      } catch (maErr) {
        mandaSyncError = maErr.message;
      }
    }

    return {
      added:             added,
      updated:           updated,
      mandatoryAdded:    mandaAdded,
      mandatoryUpdated:  mandaUpdated,
      mandatorySyncError: mandaSyncError
    };
  } catch (e) {
    throw new Error("simpanBatchTraining gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusTraining(id) {
  try {
    const sheet = initSheetTraining();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Riwayat pelatihan berhasil dihapus.";
      }
    }
    throw new Error("ID Training tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusTraining gagal: " + e.message);
  }
}