/**
 * =========================================================
 * FILE: Summary.gs
 * FUNGSI: Penyimpanan & pembacaan data laporan bulanan
 *         (sheet "HRLP_Riwayat") dan MonthlyHistory (CLOSED).
 *
 * Alur:
 *   - simpanDataLaporan(appData) → upsert berdasarkan periode
 *   - getRiwayatLengkap()        → gabung CLOSED (MonthlyHistory)
 *                                  + OPEN (HRLP_Riwayat)
 *   - hapusRiwayat(periode, ts)  → hapus dari HRLP_Riwayat
 *                                  (CLOSED tidak bisa dihapus)
 * =========================================================
 */

// ── Helpers ───────────────────────────────────────────────
const NAMA_BULAN_ = [
  "Januari","Februari","Maret","April","Mei","Juni",
  "Juli","Agustus","September","Oktober","November","Desember"
];

function getLabelPeriode_(bulanVal, tahunVal) {
  const bNum = Number(bulanVal);
  if (!isNaN(bNum) && bNum >= 1 && bNum <= 12) {
    return NAMA_BULAN_[bNum - 1] + " " + tahunVal;
  }
  return String(bulanVal).trim() + " " + tahunVal;
}

// ── Init Sheet ────────────────────────────────────────────
function initSheetRiwayat() {
  const ss = getDB();
  let sheet = ss.getSheetByName("HRLP_Riwayat");
  if (!sheet) {
    sheet = ss.insertSheet("HRLP_Riwayat");
    sheet.appendRow([
      "Periode", "Timestamp", "Rev_Actual", "Rev_Budget",
      "FO_Payroll",  "FO_DW",  "FO_Casual",  "FO_Meals",  "FO_Arrivals",
      "HK_Payroll",  "HK_DW",  "HK_Casual",  "HK_Meals",  "HK_Occupied",
      "FBS_Payroll", "FBS_DW", "FBS_Casual",  "FBS_Meals", "FBS_Cover",
      "FBP_Payroll", "FBP_DW", "FBP_Casual",  "FBP_Meals", "FBP_FoodCover",
      "ENG_Payroll", "ENG_DW", "ENG_Casual",  "ENG_Meals",
      "HRD_Payroll", "HRD_DW", "HRD_Casual",  "HRD_Meals", "HRD_Training", "HRD_TotalEmp",
      "Data_JSON"
    ]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Cek apakah periode sudah CLOSED ───────────────────────
function isPeriodeClosed_(periodeLabel) {
  const ss      = getDB();
  const history = ss.getSheetByName("MonthlyHistory");
  if (!history) return false;
  const rows = history.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const label  = getLabelPeriode_(rows[i][1], rows[i][2]);
    const status = String(rows[i][7] || "");
    if (label.trim() === String(periodeLabel).trim() && status === "CLOSED") return true;
  }
  return false;
}

// ── Save ─────────────────────────────────────────────────
function simpanDataLaporan(appData) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk, coba lagi."); }

  try {
    const periode = appData.periode;
    if (!periode) throw new Error("Periode tidak ditemukan dalam data!");

    if (isPeriodeClosed_(periode)) {
      throw new Error("Periode " + periode + " telah CLOSED dan dikunci. Data tidak dapat diperbarui.");
    }

    const sheet = initSheetRiwayat();
    const rows  = sheet.getDataRange().getValues();
    let rowIdx  = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(periode).trim()) {
        rowIdx = i + 1;
        break;
      }
    }

    const d = appData;
    const rowData = [
      periode, new Date(),
      d.umum?.totRevAct  || 0, d.umum?.totRevBud  || 0,
      d.fo?.payTot  || 0, d.fo?.qtyDW  || 0, d.fo?.qtyCas  || 0, 0, d.fo?.arrAct  || 0,
      d.hk?.payTot  || 0, d.hk?.qtyDW  || 0, d.hk?.qtyCas  || 0, 0, d.hk?.occRoom || 0,
      d.fbs?.payTot || 0, d.fbs?.qtyDW || 0, d.fbs?.qtyCas || 0, 0, d.fbs?.cover  || 0,
      d.fbp?.payTot || 0, d.fbp?.qtyDW || 0, d.fbp?.qtyCas || 0, 0, d.fbp?.coverFood || 0,
      d.eng?.payTot || 0, d.eng?.qtyDW || 0, d.eng?.qtyCas || 0, 0,
      d.hrd?.payTot || 0, d.hrd?.qtyDW || 0, d.hrd?.qtyCas || 0,
      d.hrd?.mealsAct || 0, d.training?.totBud || 0, 0,
      JSON.stringify(d)
    ];

    if (rowIdx !== -1) {
      sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return "Data laporan berhasil disimpan.";
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Read ─────────────────────────────────────────────────
function getRiwayatLengkap() {
  try {
    const result      = [];
    const seenPeriods = {};
    const ss          = getDB();

    // 1. Ambil data CLOSED dari MonthlyHistory
    const history = ss.getSheetByName("MonthlyHistory");
    if (history) {
      const displayVals = history.getDataRange().getDisplayValues();
      const rawVals     = history.getDataRange().getValues();

      for (let i = 1; i < rawVals.length; i++) {
        const dRow = displayVals[i];
        const rRow = rawVals[i];
        const label  = getLabelPeriode_(rRow[1], rRow[2]);
        const status = dRow[7] || "CLOSED";

        // Cari kolom yang berisi JSON string valid (dari kanan ke kiri)
        let jsonStr = "";
        for (let c = rRow.length - 1; c >= 0; c--) {
          const v = String(rRow[c] || "").trim();
          if (v.startsWith("{") && v.endsWith("}")) { jsonStr = v; break; }
        }

        // Fallback: bangun JSON minimal dari data tabular
        if (!jsonStr) {
          jsonStr = JSON.stringify({
            status: status, periode: label,
            catatan: rRow[27] || "",
            umum: { totRevAct: parseFloat(rRow[17]) || 0, totRevBud: parseFloat(rRow[26]) || 0 },
            hrd: { payTot: parseFloat(rRow[18]) || 0, mealsAct: parseFloat(rRow[19]) || 0 }
          });
        }

        result.push({
          tanggalLengkap: dRow[3] + " " + dRow[4],
          periode: label,
          totRevAct: parseFloat(rRow[17]) || 0,
          rawData: jsonStr,
          status: status
        });
        seenPeriods[label] = true;
      }
    }

    // 2. Ambil data OPEN dari HRLP_Riwayat
    const sheet = initSheetRiwayat();
    const dVals = sheet.getDataRange().getDisplayValues();
    const rVals = sheet.getDataRange().getValues();

    for (let i = 1; i < rVals.length; i++) {
      const label = dVals[i][0] || "N/A";
      if (seenPeriods[label]) continue; // sudah ada di CLOSED, lewati

      const jsonStr = rVals[i][rVals[i].length - 1] || "";
      let parsed    = {};
      try { parsed = JSON.parse(jsonStr); } catch (e) { /* abaikan */ }

      result.push({
        tanggalLengkap: dVals[i][1] || "N/A",
        periode: label,
        totRevAct: parseFloat(rVals[i][2]) || 0,
        rawData: jsonStr,
        status: parsed.status || "OPEN"
      });
    }

    return result;
  } catch (e) {
    throw new Error("getRiwayatLengkap gagal: " + e.message);
  }
}

// ── Delete ────────────────────────────────────────────────
function hapusRiwayat(periodeTarget, timestampTarget) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (e) { throw new Error("Server sedang sibuk, coba lagi."); }

  try {
    if (isPeriodeClosed_(periodeTarget)) {
      throw new Error("Periode " + periodeTarget + " telah CLOSED. Data historis tidak dapat dihapus.");
    }

    const sheet = initSheetRiwayat();
    const rows  = sheet.getDataRange().getDisplayValues();

    for (let i = 1; i < rows.length; i++) {
      const matchPeriode = String(rows[i][0]).trim() === String(periodeTarget).trim();
      const matchTs      = timestampTarget
                           ? String(rows[i][1]).trim() === String(timestampTarget).trim()
                           : true;
      if (matchPeriode && matchTs) {
        sheet.deleteRow(i + 1);
        return "Riwayat laporan berhasil dihapus.";
      }
    }
    throw new Error("Data riwayat tidak ditemukan untuk periode: " + periodeTarget);
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}