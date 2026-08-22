/**
 * =========================================================
 * FILE: Perizinan.gs
 * FUNGSI: CRUD Data Perizinan Hotel (sheet "HRLP_Perizinan").
 *         Dokumen fisik disimpan ke Google Drive folder
 *         "HR_Perizinan_Dokumen".
 *
 * Hak Akses:
 *   - Mode Biasa : baca data perizinan tanpa URL dokumen
 *   - Mode Admin : baca + tulis + hapus + akses dokumen Drive
 *
 * Struktur kolom sheet (16 kolom):
 *   A: ID Perizinan | B: Nama Perizinan | C: Instansi
 *   D: Nomor Dokumen | E: Tanggal Mulai | F: Tanggal Akhir
 *   G: Due Date | H: PIC | I: Perkiraan Biaya (Rp)
 *   J: Keterangan | K: Nama File | L: File ID | M: File URL
 *   N: Tanggal Input | O: Last Updated | P: Updated By
 * =========================================================
 */

const SHEET_PERIZINAN    = "HRLP_Perizinan";
const FOLDER_PERIZINAN   = "HR_Perizinan_Dokumen";
const PERIZINAN_HEADERS  = [
  "ID Perizinan", "Nama Perizinan", "Instansi", "Nomor Dokumen",
  "Tanggal Mulai", "Tanggal Akhir", "Due Date", "PIC",
  "Perkiraan Biaya (Rp)", "Keterangan",
  "Nama File", "File ID", "File URL",
  "Tanggal Input", "Last Updated", "Updated By"
];

// ── Drive Helpers ─────────────────────────────────────────
function _getPerizinanFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_PERIZINAN);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_PERIZINAN);
}

/**
 * Upload file ke Drive, kembalikan { fileId, fileUrl, fileName }.
 * Hanya Admin yang diizinkan.
 */
function uploadFilePerizinanDrive(base64Data, fileName, mimeType, isAdmin) {
  if (!isAdmin) throw new Error("Akses ditolak: hanya Admin yang dapat mengunggah dokumen perizinan.");
  if (!base64Data) throw new Error("Data file kosong.");
  try {
    const folder = _getPerizinanFolder();
    const bytes  = Utilities.base64Decode(base64Data);
    const blob   = Utilities.newBlob(bytes, mimeType || "application/octet-stream", fileName || "dokumen_perizinan");
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (e) {
    throw new Error("Gagal mengunggah dokumen perizinan: " + e.message);
  }
}

// ── Date Helpers ──────────────────────────────────────────
function _fmtDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

// ── Init Sheet ────────────────────────────────────────────
function initSheetPerizinan_() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_PERIZINAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PERIZINAN);
    sheet.appendRow(PERIZINAN_HEADERS);
    sheet.getRange(1, 1, 1, PERIZINAN_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Parse multi-file dari satu sel ───────────────────────
function _parseFiles(fileNameStr, fileIdStr, fileUrlStr) {
  if (!fileUrlStr) return [];
  // Coba JSON array terlebih dulu
  if (fileUrlStr.trim().startsWith("[")) {
    try { return JSON.parse(fileUrlStr); } catch (e) { /* lanjut ke fallback */ }
  }
  // Fallback: pipe-separated
  const names = fileNameStr ? fileNameStr.split(" | ") : [];
  const ids   = fileIdStr   ? fileIdStr.split(" | ")   : [];
  const urls  = fileUrlStr  ? fileUrlStr.split(" | ")  : [];
  const len   = Math.max(names.length, ids.length, urls.length);
  const files = [];
  for (let k = 0; k < len; k++) {
    if (urls[k] || names[k]) {
      files.push({ fileName: names[k] || "Dokumen", fileId: ids[k] || ("FILE-" + k), fileUrl: urls[k] || "" });
    }
  }
  return files;
}

// ── Read ──────────────────────────────────────────────────
function getPerizinan(isAdmin) {
  try {
    const sheet = initSheetPerizinan_();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    // Cari indeks kolom "Perkiraan Biaya" secara dinamis (backward compat)
    const headers  = rows[0].map(h => String(h || "").trim());
    const biayaCol = headers.indexOf("Perkiraan Biaya (Rp)");

    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;

      let biaya = 0, ket = "", fName = "", fId = "", fUrl = "", tInput = "", lUpd = "", upBy = "";
      if (biayaCol > -1) {
        biaya  = parseFloat(r[biayaCol])     || 0;
        ket    = String(r[biayaCol + 1] || "");
        fName  = String(r[biayaCol + 2] || "");
        fId    = String(r[biayaCol + 3] || "");
        fUrl   = String(r[biayaCol + 4] || "");
        tInput = String(r[biayaCol + 5] || "");
        lUpd   = String(r[biayaCol + 6] || "");
        upBy   = String(r[biayaCol + 7] || "");
      } else {
        // Sheet lama tanpa kolom biaya (backward compat)
        ket    = String(r[8]  || "");
        fName  = String(r[9]  || "");
        fId    = String(r[10] || "");
        fUrl   = String(r[11] || "");
        tInput = String(r[12] || "");
        lUpd   = String(r[13] || "");
        upBy   = String(r[14] || "");
      }

      let files = _parseFiles(fName, fId, fUrl);

      // Non-admin tidak menerima URL/file dokumen
      if (!isAdmin) { fName = ""; fId = ""; fUrl = ""; files = []; }

      result.push({
        id:            String(r[0]),
        namaPerizinan: String(r[1] || ""),
        instansi:      String(r[2] || ""),
        nomorDokumen:  String(r[3] || ""),
        tanggalMulai:  _fmtDate(r[4]),
        tanggalAkhir:  r[5] && String(r[5]).toLowerCase().includes("seumur") ? "Seumur Hidup" : _fmtDate(r[5]),
        dueDate:       _fmtDate(r[6]),
        pic:           String(r[7] || ""),
        biaya:         biaya,
        keterangan:    ket,
        fileName:      fName,
        fileId:        fId,
        fileUrl:       fUrl,
        files:         files,
        tanggalInput:  tInput,
        lastUpdated:   lUpd,
        updatedBy:     upBy
      });
    }
    return result;
  } catch (e) {
    Logger.log("getPerizinan error: " + e.message);
    return [];
  }
}

// ── Generate ID baru ──────────────────────────────────────
function generateIdPerizinan_(sheet) {
  const rows = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || "");
    if (id.startsWith("PER-")) {
      const n = parseInt(id.slice(4), 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  return "PER-" + String(maxNum + 1).padStart(6, "0");
}

// ── Save ─────────────────────────────────────────────────
function simpanPerizinan(p, isAdmin) {
  if (!isAdmin) throw new Error("Akses ditolak: hanya Admin yang dapat menyimpan data perizinan.");

  try {
    const sheet   = initSheetPerizinan_();
    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h || "").trim());
    const biayaCol = headers.indexOf("Perkiraan Biaya (Rp)");

    // Tambahkan kolom Biaya jika belum ada (migrasi sheet lama)
    if (biayaCol === -1) {
      sheet.insertColumnAfter(8);
      sheet.getRange(1, 9).setValue("Perkiraan Biaya (Rp)");
    }

    // Cari baris existing berdasarkan ID
    let rowIndex   = -1;
    let exFName    = "", exFId = "", exFUrl = "";
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(p.id)) {
        rowIndex = i + 1;
        const bc = headers.indexOf("Perkiraan Biaya (Rp)");
        if (bc > -1) {
          exFName = String(rows[i][bc + 2] || "");
          exFId   = String(rows[i][bc + 3] || "");
          exFUrl  = String(rows[i][bc + 4] || "");
        }
        // Pertahankan tanggal existing jika tidak diisi
        if (!p.tanggalMulai && rows[i][4]) p.tanggalMulai = _fmtDate(rows[i][4]);
        if (!p.tanggalAkhir && rows[i][5]) p.tanggalAkhir = String(rows[i][5]);
        if (!p.dueDate      && rows[i][6]) p.dueDate      = _fmtDate(rows[i][6]);
        break;
      }
    }

    // Gabungkan file info
    const filesArr = p.files || [];
    let fNameStr, fIdStr, fUrlStr;
    if (filesArr.length > 0) {
      fNameStr = filesArr.map(f => f.fileName || "").join(" | ");
      fIdStr   = filesArr.map(f => f.fileId   || "").join(" | ");
      fUrlStr  = JSON.stringify(filesArr.map(f => ({ fileName: f.fileName || "", fileId: f.fileId || "", fileUrl: f.fileUrl || "" })));
    } else if (p.fileUrl) {
      fNameStr = p.fileName || "";
      fIdStr   = p.fileId   || "";
      fUrlStr  = p.fileUrl  || "";
    } else {
      // Pertahankan dokumen lama jika tidak ada perubahan
      fNameStr = exFName;
      fIdStr   = exFId;
      fUrlStr  = exFUrl;
    }

    const rowValues = [
      p.id, p.namaPerizinan || "", p.instansi || "", p.nomorDokumen || "",
      p.tanggalMulai || "", p.tanggalAkhir || "", p.dueDate || "",
      p.pic || "", parseFloat(p.biaya) || 0, p.keterangan || "",
      fNameStr, fIdStr, fUrlStr,
      p.tanggalInput || "", p.lastUpdated || "", p.updatedBy || ""
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    return "Data perizinan berhasil disimpan.";
  } catch (e) {
    throw new Error("simpanPerizinan gagal: " + e.message);
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusPerizinan(id, isAdmin) {
  if (!isAdmin) throw new Error("Akses ditolak: hanya Admin yang dapat menghapus data perizinan.");
  if (!id) throw new Error("ID perizinan tidak valid.");

  try {
    const sheet = initSheetPerizinan_();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Data perizinan berhasil dihapus.";
      }
    }
    throw new Error("Record dengan ID " + id + " tidak ditemukan.");
  } catch (e) {
    throw new Error("hapusPerizinan gagal: " + e.message);
  }
}

// ── Import Batch dari Excel ───────────────────────────────
function importPerizinanExcel(importDataList, isAdmin) {
  if (!isAdmin) return { success: false, message: "Akses ditolak. Fitur hanya tersedia untuk Admin." };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("Server sedang sibuk, coba lagi.");

  try {
    const sheet = initSheetPerizinan_();
    const rows  = sheet.getDataRange().getValues();

    // Kumpulkan nomor dokumen existing untuk cek duplikat
    const existing = {};
    for (let i = 1; i < rows.length; i++) {
      const n = String(rows[i][3] || "").trim().toLowerCase();
      if (n) existing[n] = true;
    }

    const result = { success: true, total: importDataList.length, imported: 0, duplicate: 0, invalid: 0, errors: [] };
    let idNum    = parseInt(generateIdPerizinan_(sheet).slice(4), 10);
    const nowStr = new Date().toISOString().slice(0, 10);
    const toApp  = [];

    importDataList.forEach(function (p, idx) {
      // Validasi
      const errs = [];
      if (!p.namaPerizinan)  errs.push("Nama Perizinan kosong");
      if (!p.instansi)       errs.push("Instansi kosong");
      if (!p.nomorDokumen)   errs.push("Nomor Dokumen kosong");
      if (!p.tanggalMulai)   errs.push("Tanggal Mulai kosong");
      if (!p.dueDate)        errs.push("Due Date kosong");
      if (!p.pic)            errs.push("PIC kosong");
      if (errs.length > 0) {
        result.invalid++;
        result.errors.push({ row: idx + 2, nama: p.namaPerizinan || "Tanpa Nama", masalah: errs.join(", ") });
        return;
      }

      const docKey = String(p.nomorDokumen || "").trim().toLowerCase();
      if (docKey && existing[docKey]) { result.duplicate++; return; }

      const newId = "PER-" + String(idNum++).padStart(6, "0");
      toApp.push([
        newId, p.namaPerizinan, p.instansi, p.nomorDokumen,
        p.tanggalMulai, p.tanggalAkhir || "", p.dueDate, p.pic,
        parseFloat(p.biaya) || 0, p.keterangan || "",
        p.fileName || "", p.fileId || "", p.fileUrl || "",
        nowStr, nowStr, "Admin Import"
      ]);
      if (docKey) existing[docKey] = true;
      result.imported++;
    });

    if (toApp.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toApp.length, toApp[0].length).setValues(toApp);
    }
    return result;
  } catch (e) {
    throw new Error("importPerizinanExcel gagal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Akses Dokumen (Admin only) ────────────────────────────
function getDokumenPerizinan(id, isAdmin) {
  if (!isAdmin) throw new Error("Akses ditolak: dokumen perizinan hanya dapat diakses dalam Mode Admin.");
  const all  = getPerizinan(true);
  const item = all.find(x => String(x.id) === String(id));
  if (!item) throw new Error("Data perizinan tidak ditemukan.");
  return item.files && item.files.length > 0
    ? item.files
    : (item.fileUrl ? [{ fileName: item.fileName, fileId: item.fileId, fileUrl: item.fileUrl }] : []);
}

function viewDokumenPerizinan(id, isAdmin)   { return getDokumenPerizinan(id, isAdmin); }
function getFilePerizinan(id, isAdmin)        { return getDokumenPerizinan(id, isAdmin); }
function uploadDokumenPerizinan(b64, fn, mt, isAdmin) { return uploadFilePerizinanDrive(b64, fn, mt, isAdmin); }
function deleteDokumenPerizinan(id, isAdmin)  { return hapusPerizinan(id, isAdmin); }