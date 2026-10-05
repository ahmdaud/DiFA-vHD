/**
 * ============================================================
 *  MODUL DATA MASTER — kelola Pejabat Penandatangan & Klasifikasi Arsip
 * ============================================================
 */

/**
 * Mencari satu baris berdasarkan nilai kolom, menggunakan TextFinder
 * (jauh lebih cepat daripada memuat & loop seluruh sheet dengan
 * getDataRange().getValues(), terutama saat sheet sudah berisi
 * banyak baris).
 */
function findRowByColumnValue_(sheetName, columnName, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = headers.indexOf(columnName) + 1;
  if (col < 1) return null;

  const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
  const found = range.createTextFinder(String(value)).matchEntireCell(true).findNext();
  if (!found) return null;

  return { sheet: sheet, rowIndex: found.getRow(), headers: headers };
}

/** Menghapus cache data master (dipanggil setelah tambah/ubah/hapus Pejabat atau Klasifikasi). */
function invalidateMasterCache_(sheetName) {
  try {
    CacheService.getScriptCache().remove('sheet_' + sheetName);
  } catch (e) {
    // cache tidak tersedia — abaikan, tidak fatal
  }
}

/* ============================================================
 *  MASTER PEJABAT
 * ============================================================ */
function tambahPejabat(data) {
  if (!data.jabatan) return { success: false, message: 'Nama jabatan wajib diisi.' };

  const existing = findRowByColumnValue_(SHEET_NAMES.MASTER_PEJABAT, 'Jabatan', data.jabatan);
  if (existing) return { success: false, message: 'Jabatan tersebut sudah ada. Gunakan tombol Edit untuk mengubahnya.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_PEJABAT);
  sheet.appendRow([data.jabatan, data.kode || '', data.namaPejabat || '', data.email || '', data.noWhatsapp || '']);
  invalidateMasterCache_(SHEET_NAMES.MASTER_PEJABAT);
  return { success: true };
}

function updatePejabat(jabatanAsli, data) {
  const found = findRowByColumnValue_(SHEET_NAMES.MASTER_PEJABAT, 'Jabatan', jabatanAsli);
  if (!found) return { success: false, message: 'Data jabatan tidak ditemukan.' };

  found.sheet.getRange(found.rowIndex, 1, 1, 5).setValues([[
    data.jabatan || jabatanAsli, data.kode || '', data.namaPejabat || '', data.email || '', data.noWhatsapp || ''
  ]]);
  invalidateMasterCache_(SHEET_NAMES.MASTER_PEJABAT);
  return { success: true };
}

function hapusPejabat(jabatan) {
  const found = findRowByColumnValue_(SHEET_NAMES.MASTER_PEJABAT, 'Jabatan', jabatan);
  if (!found) return { success: false, message: 'Data jabatan tidak ditemukan.' };
  found.sheet.deleteRow(found.rowIndex);
  invalidateMasterCache_(SHEET_NAMES.MASTER_PEJABAT);
  return { success: true };
}

/* ============================================================
 *  MASTER KLASIFIKASI
 * ============================================================ */
function tambahKlasifikasi(data) {
  if (!data.kode) return { success: false, message: 'Kode klasifikasi wajib diisi.' };

  const existing = findRowByColumnValue_(SHEET_NAMES.MASTER_KLASIFIKASI, 'Kode Klasifikasi', data.kode);
  if (existing) return { success: false, message: 'Kode klasifikasi tersebut sudah ada. Gunakan tombol Edit untuk mengubahnya.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_KLASIFIKASI);
  sheet.appendRow([data.kode, data.nama || '', data.keterangan || '']);
  invalidateMasterCache_(SHEET_NAMES.MASTER_KLASIFIKASI);
  return { success: true };
}

function updateKlasifikasi(kodeAsli, data) {
  const found = findRowByColumnValue_(SHEET_NAMES.MASTER_KLASIFIKASI, 'Kode Klasifikasi', kodeAsli);
  if (!found) return { success: false, message: 'Data klasifikasi tidak ditemukan.' };

  found.sheet.getRange(found.rowIndex, 1, 1, 3).setValues([[
    data.kode || kodeAsli, data.nama || '', data.keterangan || ''
  ]]);
  invalidateMasterCache_(SHEET_NAMES.MASTER_KLASIFIKASI);
  return { success: true };
}

function hapusKlasifikasi(kode) {
  const found = findRowByColumnValue_(SHEET_NAMES.MASTER_KLASIFIKASI, 'Kode Klasifikasi', kode);
  if (!found) return { success: false, message: 'Data klasifikasi tidak ditemukan.' };
  found.sheet.deleteRow(found.rowIndex);
  invalidateMasterCache_(SHEET_NAMES.MASTER_KLASIFIKASI);
  return { success: true };
}