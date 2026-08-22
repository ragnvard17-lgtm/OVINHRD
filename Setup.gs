/**
 * =========================================================
 * FILE: Setup.gs
 * FUNGSI: Inisialisasi SEMUA sheet database dalam satu kali
 *         eksekusi + menu khusus "HR LP" di Google Spreadsheet.
 *
 * CARA PAKAI:
 *   1. Buka Google Spreadsheet tempat script ini ter-attach.
 *   2. Reload / refresh halaman Spreadsheet.
 *   3. Menu "HR LP" → "Setup / Perbaiki Semua Sheet".
 *      Semua sheet (Data_Karyawan, Data_Training, dst.) akan
 *      dibuat otomatis beserta header jika belum ada.
 *      AMAN dijalankan ulang — tidak menghapus data yang sudah ada.
 *
 * Daftar sheet yang dibuat oleh setupSemuaSheet():
 *   Data_Karyawan, Data_Training, Data_Mandatory_Assignment,
 *   Data_Daily_Report, HR_Calendar, Data_Karyawan_Sakit,
 *   Data_Karyawan_Terlambat, HRLP_Riwayat,
 *   HRLP_CK_Kategori (+ seed 17 kategori bawaan), HRLP_CK_Dokumen
 *
 *   Note: HRLP_Perizinan dibuat otomatis saat data perizinan
 *   pertama kali disimpan via simpanPerizinan() di Perizinan.gs.
 * =========================================================
 */

/** Dipanggil otomatis saat Spreadsheet dibuka — tambahkan menu "HR LP". */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("HR LP")
    .addItem("Setup / Perbaiki Semua Sheet", "setupSemuaSheet")
    .addSeparator()
    .addItem("Buka Web App", "bukaWebApp_")
    .addToUi();
}

/** Placeholder — Web App URL perlu diganti manual setelah deploy. */
function bukaWebApp_() {
  SpreadsheetApp.getUi().alert(
    "Buka Web App",
    "Salin URL Web App dari menu Deploy → Manage Deployments dan buka di browser baru.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Inisialisasi semua sheet database.
 * Dipanggil dari menu "HR LP" → "Setup / Perbaiki Semua Sheet".
 */
function setupSemuaSheet() {
  const tasks = [
    { nama: "Data_Karyawan",             fn: initSheetKaryawan           },
    { nama: "Data_Training",             fn: initSheetTraining           },
    { nama: "Data_Mandatory_Assignment", fn: initSheetMandatoryAssignment },
    { nama: "Data_Daily_Report",         fn: initSheetDailyReport        },
    { nama: "HR_Calendar",               fn: initSheetHRCalendar         },
    { nama: "Data_Karyawan_Sakit",       fn: initSheetKaryawanSakit      },
    { nama: "Data_Karyawan_Terlambat",   fn: initSheetKaryawanTerlambat  },
    { nama: "HRLP_Riwayat",             fn: initSheetRiwayat            },
    { nama: "HRLP_CK_Kategori",         fn: initSheetCKKategori         },
    { nama: "HRLP_CK_Dokumen",          fn: initSheetCKDokumen          }
  ];

  const hasil = tasks.map(function (t) {
    try {
      if (typeof t.fn !== "function") return t.nama + ": DILEWATI (fungsi tidak ditemukan)";
      t.fn();
      return t.nama + ": OK ✓";
    } catch (e) {
      return t.nama + ": GAGAL — " + e.message;
    }
  });

  // HRLP_Perizinan dibuat otomatis oleh Perizinan.gs
  hasil.push("HRLP_Perizinan: (dibuat otomatis saat data pertama disimpan)");

  const pesan = hasil.join("\n");
  Logger.log(pesan);

  try {
    SpreadsheetApp.getUi().alert("Setup Selesai", pesan, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Jika dijalankan dari Editor bukan dari menu, UI tidak tersedia — abaikan.
  }

  return pesan;
}