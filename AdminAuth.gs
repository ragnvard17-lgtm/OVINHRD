/**
 * =========================================================
 * FILE: AdminAuth.gs
 * FUNGSI: Validasi password Admin untuk fitur "Mode Admin".
 *         Berjalan 100% di sisi server — password ASLI tidak
 *         pernah dikirim balik ke browser.
 *
 * CARA GANTI PASSWORD:
 *   Opsi 1 (langsung di kode):
 *     Ubah nilai ADMIN_PASSWORD_DEFAULT di bawah → simpan → deploy ulang.
 *
 *   Opsi 2 (tanpa edit kode / tanpa redeploy):
 *     Apps Script Editor → Project Settings (⚙) → Script Properties
 *     → tambah Key = ADMIN_PASSWORD, Value = password baru.
 *     Nilai Script Property ini akan MENGGANTIKAN ADMIN_PASSWORD_DEFAULT.
 * =========================================================
 */

const ADMIN_PASSWORD_DEFAULT = "ovin123"; // Ganti sesuai kebutuhan

/**
 * Dipanggil frontend via google.script.run.cekPasswordAdmin(password).
 * Mengembalikan true jika password cocok, false jika tidak.
 * Tidak pernah mengembalikan password asli.
 */
function cekPasswordAdmin(inputPassword) {
  try {
    if (!inputPassword) return false;
    const props = PropertiesService.getScriptProperties();
    const password = props.getProperty("ADMIN_PASSWORD") || ADMIN_PASSWORD_DEFAULT;
    return String(inputPassword) === String(password);
  } catch (e) {
    // Gagal aman: jika ada error tak terduga, tolak akses
    Logger.log("cekPasswordAdmin error: " + e.message);
    return false;
  }
}