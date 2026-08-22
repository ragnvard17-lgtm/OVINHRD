/**
 * =========================================================
 * FILE: CorporateKnowledge.gs
 * FUNGSI: Backend Corporate Knowledge (CK) — pusat penyimpanan
 *         dokumen SOP, kebijakan, formulir, dsb.
 *
 *   Sheet metadata kategori  : "HRLP_CK_Kategori"
 *   Sheet metadata dokumen   : "HRLP_CK_Dokumen"
 *   Folder file fisik Drive  : "HR_Corporate_Knowledge_Dokumen"
 *
 * Fungsi yang dipanggil frontend:
 *   getCKData(isAdmin)                   → baca kategori + dokumen
 *   simpanCKKategori(kat, isAdmin)       → tambah/edit kategori
 *   hapusCKKategori(id, isAdmin)         → hapus kategori
 *   simpanCKDokumen(payload, isAdmin)    → tambah/edit metadata dokumen
 *   perbaruiVersiCKDokumen(payload, isAdmin) → rilis versi baru
 *   hapusCKDokumen(id, isAdmin)          → hapus dokumen + file Drive
 *   unggahCKDokumen(payload, isAdmin)    → upload file baru ke Drive
 *   getCKFileData(fileId, asPdf)         → baca file sebagai base64
 * =========================================================
 */

const SHEET_CK_KATEGORI   = "HRLP_CK_Kategori";
const SHEET_CK_DOKUMEN    = "HRLP_CK_Dokumen";
const FOLDER_CK           = "HR_Corporate_Knowledge_Dokumen";

const CK_KATEGORI_HEADERS = ["ID", "Nama", "Icon", "Warna", "Status", "Deskripsi", "Waktu Update"];
const CK_DOKUMEN_HEADERS  = [
  "ID", "Nomor Dokumen", "Nama Dokumen", "Kategori ID", "Deskripsi", "Status",
  "Versi Aktif", "Nama File", "Tipe File", "Ukuran File (Bytes)",
  "File ID Drive", "File URL", "Diunggah Oleh", "Dibuat Pada", "Diupdate Pada",
  "Riwayat Versi (JSON)"
];

// 17 kategori bawaan — sinkron dengan window.DEFAULT_CK_CATEGORIES di Index.html
const CK_DEFAULT_CATEGORIES = [
  { id: "cat-sop-operasional",       nama: "SOP Operasional",          icon: "fa-clipboard-list",       warna: "sky",     deskripsi: "Standar Operasional Prosedur seluruh divisi operasional hotel" },
  { id: "cat-kebijakan-regulasi",    nama: "Kebijakan & Regulasi",     icon: "fa-scale-balanced",       warna: "indigo",  deskripsi: "Kebijakan resmi manajemen dan kepatuhan regulasi eksternal" },
  { id: "cat-formulir-template",     nama: "Formulir & Template",      icon: "fa-file-lines",           warna: "emerald", deskripsi: "Template dokumen, form permohonan, dan formulir standar" },
  { id: "cat-santika-management-system", nama: "Santika Management System", icon: "fa-hotel",           warna: "blue",    deskripsi: "Sistem manajemen mutu dan standar operasional Santika" },
  { id: "cat-indonesian-home",       nama: "Indonesian Home",          icon: "fa-house-chimney-user",   warna: "amber",   deskripsi: "Konsep keramahtamahan khas budaya Indonesia (Hospitality Values)" },
  { id: "cat-hr-policy",             nama: "HR Policy",                icon: "fa-user-shield",          warna: "purple",  deskripsi: "Ketentuan SDM, etika kerja, hak & kewajiban karyawan" },
  { id: "cat-training-material",     nama: "Training Material",        icon: "fa-graduation-cap",       warna: "teal",    deskripsi: "Modul pelatihan, materi orientasi, dan pengembangan kompetensi" },
  { id: "cat-health-safety",         nama: "Health & Safety",          icon: "fa-heart-pulse",          warna: "rose",    deskripsi: "Keselamatan, kesehatan kerja, sanitasi, dan mitigasi darurat" },
  { id: "cat-peraturan-perusahaan",  nama: "Peraturan Perusahaan",     icon: "fa-book-bookmark",        warna: "violet",  deskripsi: "Buku peraturan perusahaan dan pedoman disiplin kerja" },
  { id: "cat-engineering",           nama: "Engineering",              icon: "fa-wrench",               warna: "orange",  deskripsi: "Pedoman pemeliharaan teknis gedung, genset, chiller, dan ME" },
  { id: "cat-front-office",          nama: "Front Office",             icon: "fa-bell-concierge",       warna: "cyan",    deskripsi: "SOP layanan tamu, reservasi, check-in/out, dan handling keluhan" },
  { id: "cat-housekeeping",          nama: "Housekeeping",             icon: "fa-broom",                warna: "lime",    deskripsi: "Standar kebersihan kamar, public area, laundry, dan linen" },
  { id: "cat-food-beverage",         nama: "Food & Beverage",          icon: "fa-utensils",             warna: "amber",   deskripsi: "Standar penyajian makanan, resep standar, sanitasi FB, dan bar" },
  { id: "cat-finance",               nama: "Finance",                  icon: "fa-coins",                warna: "emerald", deskripsi: "Prosedur keuangan, petty cash, purchasing, invoicing, dan audit" },
  { id: "cat-sales-marketing",       nama: "Sales & Marketing",        icon: "fa-bullhorn",             warna: "pink",    deskripsi: "Materi promosi, rate structure, paket event, dan marketing collateral" },
  { id: "cat-hom",                   nama: "HOM",                      icon: "fa-briefcase",            warna: "slate",   deskripsi: "Head of Maintenance / Hotel Operations Management guidelines" },
  { id: "cat-lainnya",               nama: "Lainnya",                  icon: "fa-folder-open",          warna: "gray",    deskripsi: "Dokumen referensi dan pengetahuan umum perusahaan" }
];

// ── Init Sheets ───────────────────────────────────────────
function initSheetCKKategori() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_CK_KATEGORI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CK_KATEGORI);
    sheet.appendRow(CK_KATEGORI_HEADERS);
    sheet.getRange(1, 1, 1, CK_KATEGORI_HEADERS.length)
         .setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    const now  = new Date();
    const rows = CK_DEFAULT_CATEGORIES.map(c => [c.id, c.nama, c.icon, c.warna, "ACTIVE", c.deskripsi, now]);
    sheet.getRange(2, 1, rows.length, CK_KATEGORI_HEADERS.length).setValues(rows);
  }
  return sheet;
}

function initSheetCKDokumen() {
  const ss = getDB();
  let sheet = ss.getSheetByName(SHEET_CK_DOKUMEN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CK_DOKUMEN);
    sheet.appendRow(CK_DOKUMEN_HEADERS);
    sheet.getRange(1, 1, 1, CK_DOKUMEN_HEADERS.length)
         .setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(16, 350);
  }
  return sheet;
}

// ── Drive Helper ──────────────────────────────────────────
function _getCKFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_CK);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_CK);
}

// ── Read ──────────────────────────────────────────────────
function getCKData(isAdmin) {
  try {
    // -- Kategori
    const catSheet = initSheetCKKategori();
    const catRows  = catSheet.getDataRange().getValues();
    let categories = [];
    for (let i = 1; i < catRows.length; i++) {
      const r = catRows[i];
      if (!r[0]) continue;
      categories.push({ id: String(r[0]), nama: r[1], icon: r[2] || "fa-folder", warna: r[3] || "indigo", status: r[4] || "ACTIVE", deskripsi: r[5] || "" });
    }
    const catMap = {};
    categories.forEach(c => { catMap[c.id] = c; });

    // -- Dokumen
    const docSheet = initSheetCKDokumen();
    const docRows  = docSheet.getDataRange().getValues();
    let documents  = [];
    for (let j = 1; j < docRows.length; j++) {
      const d = docRows[j];
      if (!d[0]) continue;
      const kategoriId = String(d[3] || "");
      const cat        = catMap[kategoriId];
      let versions     = [];
      try { versions = d[15] ? JSON.parse(d[15]) : []; } catch (e) { versions = []; }

      documents.push({
        id:            String(d[0]),
        nomorDokumen:  String(d[1] || "000"),
        nomor:         String(d[1] || "000"),
        nama:          d[2] || "",
        namaDokumen:   d[2] || "",
        kategoriId:    kategoriId,
        category:      cat ? cat.nama : "",
        kategoriNama:  cat ? cat.nama : "",
        deskripsi:     d[4] || "",
        status:        d[5] || "ACTIVE",
        version:       d[6] || "v1.0",
        versi:         d[6] || "v1.0",
        fileName:      d[7] || "",
        fileType:      d[8] || "",
        fileSize:      Number(d[9]) || 0,
        fileId:        d[10] || "",
        fileUrl:       d[11] || "",
        uploader:      d[12] || "",
        createdAt:     d[13] ? new Date(d[13]).toISOString() : "",
        tanggalUpload: d[13] ? new Date(d[13]).toISOString() : "",
        updatedAt:     d[14] ? new Date(d[14]).toISOString() : "",
        versions:      versions
      });
    }

    // Non-admin hanya melihat data ACTIVE
    if (!isAdmin) {
      categories = categories.filter(c => c.status === "ACTIVE");
      documents  = documents.filter(d => d.status  === "ACTIVE");
    }

    return { success: true, categories, documents };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Kategori: Create / Update ─────────────────────────────
function simpanCKKategori(kat, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat mengelola kategori." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    const sheet = initSheetCKKategori();
    const rows  = sheet.getDataRange().getValues();
    const waktu = new Date();
    const vals  = [kat.nama || "", kat.icon || "fa-folder", kat.warna || "indigo", kat.status || "ACTIVE", kat.deskripsi || "", waktu];
    let updated = false;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(kat.id)) {
        sheet.getRange(i + 1, 2, 1, vals.length).setValues([vals]);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow([kat.id].concat(vals));

    return { success: true, message: "Kategori berhasil disimpan." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Kategori: Delete ──────────────────────────────────────
function hapusCKKategori(id, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat menghapus kategori." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    const sheet = initSheetCKKategori();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Kategori berhasil dihapus." };
      }
    }
    return { success: false, error: "Kategori tidak ditemukan." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Dokumen: Upload file baru ke Drive ────────────────────
/**
 * Upload file ke Drive dan simpan metadata ke sheet.
 * payload = { id, nomorDokumen, nama, kategoriId, deskripsi, status,
 *             version, fileName, fileType, fileSize, fileData (base64),
 *             uploader, createdAt }
 */
function unggahCKDokumen(payload, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat mengunggah dokumen." };
  if (!payload || !payload.fileData) return { success: false, error: "Data file kosong." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    // Jaring pengaman: id WAJIB ada agar baris punya Primary Key yang valid.
    // Tanpa ini, getCKData() akan menyaring/menyembunyikan baris tsb dari semua user.
    if (!payload.id) {
      payload.id = "CK-DOC-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    }

    const folder   = _getCKFolder();
    const commaIdx = payload.fileData.indexOf(",");
    const b64      = commaIdx > -1 ? payload.fileData.slice(commaIdx + 1) : payload.fileData;
    const blob     = Utilities.newBlob(Utilities.base64Decode(b64), payload.fileType || "application/octet-stream", payload.fileName || "dokumen");
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sheet   = initSheetCKDokumen();
    const rows    = sheet.getDataRange().getValues();
    const now     = new Date();
    const version = payload.version || "v1.0";

    // Nomor dokumen dihitung otomatis di server (bukan dari client) supaya konsisten
    // dan tidak bentrok saat beberapa dokumen diunggah berurutan (mis. bulk upload).
    let maxNum = 0;
    for (let i = 1; i < rows.length; i++) {
      const n = parseInt(rows[i][1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
    const nomorDokumen = payload.nomorDokumen || String(maxNum + 1).padStart(3, "0");

    const rowVal  = [
      payload.id, nomorDokumen, payload.nama || "",
      payload.kategoriId || "", payload.deskripsi || "", payload.status || "ACTIVE",
      version, file.getName(), payload.fileType || "", payload.fileSize || blob.getBytes().length,
      file.getId(), file.getUrl(), payload.uploader || "Admin HR",
      now, now, JSON.stringify([])
    ];

    // Cek ID existing
    let rowIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) { rowIdx = i + 1; break; }
    }
    if (rowIdx > -1) {
      sheet.getRange(rowIdx, 1, 1, rowVal.length).setValues([rowVal]);
    } else {
      sheet.appendRow(rowVal);
    }

    return {
      success: true,
      message: "Dokumen berhasil diunggah.",
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      document: {
        id: payload.id,
        nomorDokumen: nomorDokumen,
        nomor: nomorDokumen,
        nama: payload.nama || "",
        kategoriId: payload.kategoriId || "",
        fileId: file.getId(),
        fileUrl: file.getUrl(),
        fileName: file.getName(),
        version: version
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Dokumen: Edit metadata (tanpa ganti file) ─────────────
function simpanCKDokumen(payload, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat menyimpan dokumen." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    const sheet = initSheetCKDokumen();
    const rows  = sheet.getDataRange().getValues();
    const now   = new Date();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[
          payload.nomorDokumen || rows[i][1],
          payload.nama         || rows[i][2],
          payload.kategoriId   || rows[i][3],
          payload.deskripsi    || rows[i][4],
          payload.status       || rows[i][5]
        ]]);
        sheet.getRange(i + 1, 15).setValue(now);
        return { success: true, message: "Dokumen berhasil diperbarui." };
      }
    }
    return { success: false, error: "Dokumen tidak ditemukan." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Dokumen: Rilis versi baru ─────────────────────────────
function perbaruiVersiCKDokumen(payload, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat merilis versi baru." };
  if (!payload || !payload.fileData) return { success: false, error: "Data berkas kosong." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    const sheet = initSheetCKDokumen();
    const rows  = sheet.getDataRange().getValues();
    let rowIdx  = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.documentId)) { rowIdx = i + 1; break; }
    }
    if (rowIdx === -1) throw new Error("Dokumen tidak ditemukan.");

    const folder   = _getCKFolder();
    const commaIdx = payload.fileData.indexOf(",");
    const b64      = commaIdx > -1 ? payload.fileData.slice(commaIdx + 1) : payload.fileData;
    const blob     = Utilities.newBlob(Utilities.base64Decode(b64), payload.fileType || "application/octet-stream", payload.fileName || "dokumen");
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const rowData    = rows[rowIdx - 1];
    let oldVersions  = [];
    try { oldVersions = rowData[15] ? JSON.parse(rowData[15]) : []; } catch (e) { oldVersions = []; }

    const newVer = payload.newVersion || payload.version || ("v" + (oldVersions.length + 1) + ".0");
    const now    = new Date();
    oldVersions.push({ version: newVer, fileName: payload.fileName || "", fileType: payload.fileType || "", fileSize: payload.fileSize || 0, uploadedAt: now.toISOString(), uploadedBy: payload.uploader || "Admin HR", notes: payload.notes || "" });

    sheet.getRange(rowIdx, 7, 1, 9).setValues([[
      newVer, payload.fileName || "", payload.fileType || "", payload.fileSize || 0,
      file.getId(), file.getUrl(), payload.uploader || "Admin HR", rowData[13], now
    ]]);
    sheet.getRange(rowIdx, 16).setValue(JSON.stringify(oldVersions));

    return { success: true, message: "Versi baru " + newVer + " berhasil dirilis." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Dokumen: Delete ───────────────────────────────────────
function hapusCKDokumen(id, isAdmin) {
  if (!isAdmin) return { success: false, error: "Akses ditolak: hanya Admin yang dapat menghapus dokumen." };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { return { success: false, error: "Server sedang sibuk, coba lagi." }; }

  try {
    const sheet = initSheetCKDokumen();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        if (rows[i][10]) {
          try { DriveApp.getFileById(rows[i][10]).setTrashed(true); }
          catch (e) { /* file mungkin sudah dihapus manual */ }
        }
        sheet.deleteRow(i + 1);
        return { success: true, message: "Dokumen berhasil dihapus." };
      }
    }
    return { success: false, error: "Dokumen tidak ditemukan." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── File Viewer: baca isi file sebagai base64 ─────────────
function _isCKOfficeFile(mimeType) {
  return [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/vnd.ms-powerpoint"
  ].includes(mimeType);
}

/**
 * Konversi file Office ke PDF via Advanced Drive Service (opsional).
 * Jika Drive API belum diaktifkan, fallback ke file asli.
 */
function _convertCKToPdf(fileId, mimeType) {
  if (typeof Drive === "undefined") {
    throw new Error("Advanced Drive Service tidak aktif -- tidak dapat mengonversi dokumen ke PDF.");
  }

  let targetMime = "application/vnd.google-apps.document";
  if (mimeType.includes("spreadsheet") || mimeType === "application/vnd.ms-excel") {
    targetMime = "application/vnd.google-apps.spreadsheet";
  } else if (mimeType.includes("presentation") || mimeType === "application/vnd.ms-powerpoint") {
    targetMime = "application/vnd.google-apps.presentation";
  }

  let tempFileId = null;
  try {
    const originalBlob = DriveApp.getFileById(fileId).getBlob();
    const tempFile      = Drive.Files.create({ name: "TEMP_CK_" + fileId, mimeType: targetMime }, originalBlob, { fields: "id" });
    tempFileId = tempFile.id;
    const pdfBlob = Drive.Files.export(tempFileId, "application/pdf");

    // Validasi ketat: hasil ekspor HARUS benar-benar berformat PDF (cek magic
    // bytes "%PDF"), jangan sampai file asli (non-PDF) lolos dan dianggap
    // berhasil dikonversi -- ini yang menyebabkan preview rusak/parsial.
    const bytes = pdfBlob.getBytes();
    const isRealPdf = bytes.length > 4 &&
      bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70; // %PDF
    if (!isRealPdf) {
      throw new Error("Hasil ekspor Drive bukan file PDF yang valid.");
    }
    return pdfBlob;
  } finally {
    if (tempFileId) {
      try { Drive.Files.remove(tempFileId); } catch (e) { /* abaikan kegagalan cleanup */ }
    }
  }
}

/**
 * Mengembalikan isi file dari Drive dalam format base64 untuk ditampilkan di browser.
 * Jika asPdf=true dan file adalah Office doc, dicoba konversi ke PDF dulu.
 */
function getCKFileData(fileId, asPdf) {
  try {
    if (!fileId) throw new Error("File ID tidak valid.");
    const file     = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    const blob     = (asPdf && mimeType !== MimeType.PDF && _isCKOfficeFile(mimeType))
                     ? _convertCKToPdf(fileId, mimeType)
                     : file.getBlob();
    const outMime  = blob.getContentType() || mimeType;
    return { success: true, data: "data:" + outMime + ";base64," + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
