/**
 * =========================================================
 * FILE: MandatoryTraining.gs
 * FUNGSI: CRUD Assignment Mandatory Training
 *         (sheet "Data_Mandatory_Assignment").
 *
 * PERBAIKAN vs versi lama:
 *   - initSheetMandatoryAssignment() kini memakai getDB()
 *     (konsisten dengan seluruh file .gs lain) — bukan lagi
 *     SpreadsheetApp.getActiveSpreadsheet() yang bisa menyebabkan
 *     data tersebar di dua spreadsheet berbeda jika SPREADSHEET_ID
 *     diisi di Code.gs.
 *
 * Struktur kolom sheet (11 kolom):
 *   A: ID | B: NIK | C: Training | D: Due Date | E: Status
 *   F: Durasi | G: Trainer Type | H: Trainer ID | I: Trainer Name
 *   J: External Instansi | K: Waktu Update
 * =========================================================
 */

const SHEET_MANDATORY_ASSIGNMENT = "Data_Mandatory_Assignment";
const MANDATORY_HEADERS = [
  "ID", "NIK", "Training", "Due Date", "Status",
  "Durasi", "Trainer Type", "Trainer ID", "Trainer Name",
  "External Instansi", "Waktu Update"
];

// ── Init Sheet ────────────────────────────────────────────
function initSheetMandatoryAssignment() {
  // PENTING: gunakan getDB() agar konsisten dengan file .gs lain.
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_MANDATORY_ASSIGNMENT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MANDATORY_ASSIGNMENT);
    sheet.appendRow(MANDATORY_HEADERS);
    sheet.getRange(1, 1, 1, MANDATORY_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Helpers ───────────────────────────────────────────────
function formatDueDateMandatory_(val) {
  if (!val) return "";
  try {
    return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (e) {
    return String(val);
  }
}

// ── Read ──────────────────────────────────────────────────
function getMandatoryAssignments() {
  try {
    const sheet = initSheetMandatoryAssignment();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[1]) continue; // skip baris tanpa NIK
      result.push({
        id:               r[0],
        nik:              String(r[1]),
        training:         r[2]  || "",
        dueDate:          formatDueDateMandatory_(r[3]),
        status:           r[4]  || "Belum",
        durasi:           parseFloat(r[5]) || 2.0,
        trainerType:      r[6]  || "Internal",
        trainerId:        String(r[7] || ""),
        trainerName:      r[8]  || "",
        externalInstansi: r[9]  || ""
      });
    }
    return result;
  } catch (e) {
    throw new Error("getMandatoryAssignments gagal: " + e.message);
  }
}

// ── Create / Update (satu record) ─────────────────────────
function simpanMandatoryAssignment(assignment) {
  try {
    const sheet  = initSheetMandatoryAssignment();
    const rows   = sheet.getDataRange().getValues();
    const waktu  = new Date();
    let updated  = false;

    const values = [
      assignment.nik,
      assignment.training         || "",
      assignment.dueDate          || "",
      assignment.status           || "Belum",
      parseFloat(assignment.durasi) || 2.0,
      assignment.trainerType      || "Internal",
      assignment.trainerId        || "",
      assignment.trainerName      || "",
      assignment.externalInstansi || "",
      waktu
    ];

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(assignment.id)) {
        sheet.getRange(i + 1, 2, 1, values.length).setValues([values]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      sheet.appendRow([assignment.id].concat(values));
    }
    return "Assignment Mandatory Training berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanMandatoryAssignment gagal: " + e.message);
  }
}

// ── Batch Create / Update ─────────────────────────────────
function simpanBatchMandatoryAssignments(assignments) {
  try {
    if (!Array.isArray(assignments) || assignments.length === 0) {
      return "Tidak ada data untuk disimpan.";
    }
    const sheet = initSheetMandatoryAssignment();
    const rows  = sheet.getDataRange().getValues();
    const waktu = new Date();
    let count   = 0;

    // Buat map id → nomor baris untuk update efisien
    const idToRow = {};
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) idToRow[String(rows[i][0])] = i + 1;
    }

    const toAppend = [];
    assignments.forEach(function (a) {
      const values = [
        a.nik,
        a.training         || "",
        a.dueDate          || "",
        a.status           || "Belum",
        parseFloat(a.durasi) || 2.0,
        a.trainerType      || "Internal",
        a.trainerId        || "",
        a.trainerName      || "",
        a.externalInstansi || "",
        waktu
      ];
      if (idToRow[String(a.id)]) {
        sheet.getRange(idToRow[String(a.id)], 2, 1, values.length).setValues([values]);
      } else {
        toAppend.push([a.id].concat(values));
      }
      count++;
    });

    if (toAppend.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    }

    return count + " Assignment Mandatory Training berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanBatchMandatoryAssignments gagal: " + e.message);
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusMandatoryAssignment(id) {
  try {
    const sheet = initSheetMandatoryAssignment();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Assignment Mandatory Training berhasil dihapus.";
      }
    }
    throw new Error("ID Assignment tidak ditemukan: " + id);
  } catch (e) {
    throw new Error("hapusMandatoryAssignment gagal: " + e.message);
  }
}