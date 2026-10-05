/**
 * ============================================================
 *  MODUL SURAT KELUAR & PENOMORAN OTOMATIS
 *  Sesuai PRD Bab 6:
 *  [No.Urut]/[Kode Penandatangan]/W8-A10/[Kode Klasifikasi]/[Bulan Romawi]/[Tahun]
 *  Contoh: 15/SEK.01/W8-A10/TI.01/IX/2026
 * ============================================================
 */

const KODE_INSTANSI = 'W8-A10';
const BULAN_ROMAWI = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII'];

function toBulanRomawi(bulanAngka) {
  return BULAN_ROMAWI[bulanAngka - 1] || '-';
}

/**
 * Nomor urut reset otomatis tiap tahun (PRD Bab 8) — dihitung dari
 * baris SuratKeluar yang tahunnya sama dengan parameter.
 */
function getNextNomorUrutKeluar(tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_KELUAR);
  if (sheet.getLastRow() < 2) return 1;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const tglCol = headers.indexOf('Tanggal');
  const noCol = headers.indexOf('No. Urut');

  let max = 0;
  values.forEach(row => {
    const tglVal = row[tglCol];
    if (!tglVal) return;
    const y = (tglVal instanceof Date) ? tglVal.getFullYear() : new Date(tglVal).getFullYear();
    if (y === Number(tahun)) {
      const n = Number(row[noCol]) || 0;
      if (n > max) max = n;
    }
  });
  return max + 1;
}

/** Dipanggil dari form untuk menampilkan preview nomor urut berikutnya (tanpa mengunci). */
function previewNomorUrutKeluar(tahun) {
  return getNextNomorUrutKeluar(tahun);
}

function findPejabatByKode(kode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = readSheetAsObjectsCached(ss, SHEET_NAMES.MASTER_PEJABAT, 30);
  return list.find(p => p['Kode'] === kode) || null;
}

/**
 * Dipanggil dari form Surat Keluar.
 * formData = {
 *   penandatanganKode, kodeKlasifikasi, tanggalSurat (YYYY-MM-DD),
 *   perihal, tujuan, status ('Diterbitkan' atau 'Konsep'),
 *   fileBase64 (opsional), fileFileName (opsional)
 * }
 */
function tambahSuratKeluar(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_KELUAR);

  const tanggalSurat = new Date(formData.tanggalSurat);
  const tahun = tanggalSurat.getFullYear();
  const bulanRomawi = toBulanRomawi(tanggalSurat.getMonth() + 1);
  const noUrut = getNextNomorUrutKeluar(tahun);

  // Penentuan penanda jenis surat (ST., UND., SE., SK.)
  let jenisTag = '';
  if (formData.jenisSurat) {
    const val = formData.jenisSurat;
    if (val.includes('Surat Tugas')) {
      jenisTag = 'ST.';
    } else if (val.includes('Undangan')) {
      jenisTag = 'UND.';
    } else if (val.includes('Surat Edaran')) {
      jenisTag = 'SE.';
    } else if (val.includes('Surat Keputusan')) {
      jenisTag = 'SK.';
    }
  }

  // Format nomor lengkap: Nomor/KodePenandatangan/KodeInstansi/JenisTag.KodeKlasifikasi/Bulan/Tahun
  const nomorLengkap = `${noUrut}/${formData.penandatanganKode}/${KODE_INSTANSI}/${jenisTag}${formData.kodeKlasifikasi}/${bulanRomawi}/${tahun}`;

  let fileUrl = '';
  if (formData.fileBase64 && formData.fileFileName) {
    fileUrl = simpanLampiranPDF(formData.fileBase64, formData.fileFileName, nomorLengkap.replace(/\//g, '_'));
  }

  const pejabat = findPejabatByKode(formData.penandatanganKode);
  const namaPejabat = pejabat ? pejabat['Jabatan'] : formData.penandatanganKode;
  const status = formData.status === 'Konsep' ? 'Konsep' : 'Diterbitkan';

  sheet.appendRow([
    noUrut, nomorLengkap, tanggalSurat, formData.penandatanganKode, namaPejabat,
    formData.kodeKlasifikasi, formData.perihal || '', formData.tujuan || '',
    fileUrl, Session.getActiveUser().getEmail() || 'Admin', status
  ]);

  return {
    success: true,
    nomorLengkap, noUrut, bulanRomawi, tahun, status
  };
}

/**
 * Mencari baris Surat Keluar berdasarkan Nomor Surat Lengkap (kunci unik).
 * Menggunakan TextFinder pada satu kolom saja, bukan membaca & meloop
 * seluruh sheet — jauh lebih cepat saat SuratKeluar sudah berisi banyak
 * baris (dipanggil setiap kali surat keluar diedit).
 */
function findSuratKeluarByNomor(nomorLengkap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_KELUAR);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const nomorCol = headers.indexOf('Nomor Surat Lengkap') + 1;
  if (nomorCol < 1) return null;

  const range = sheet.getRange(2, nomorCol, sheet.getLastRow() - 1, 1);
  const found = range.createTextFinder(String(nomorLengkap)).matchEntireCell(true).findNext();
  if (!found) return null;

  const rowIndex = found.getRow();
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const obj = {};
  headers.forEach((h, idx) => obj[h] = rowValues[idx]);
  return { rowIndex: rowIndex, data: obj, headers };
}

/**
 * Memperbarui Surat Keluar yang sudah ada. Nomor Surat, Tanggal, Penandatangan,
 * dan Kode Klasifikasi TIDAK DAPAT diubah karena sudah menjadi bagian dari
 * nomor surat yang terbit — hanya Perihal, Tujuan, Status, dan File yang bisa diedit.
 * formData = { perihal, tujuan, status, fileBase64 (opsional), fileFileName (opsional) }
 */
function updateSuratKeluar(nomorLengkap, formData) {
  const suratRow = findSuratKeluarByNomor(nomorLengkap);
  if (!suratRow) return { success: false, message: 'Surat dengan nomor tersebut tidak ditemukan.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_KELUAR);
  const headers = suratRow.headers;

  const setCell = (colName, value) => {
    const col = headers.indexOf(colName) + 1;
    if (col > 0) sheet.getRange(suratRow.rowIndex, col).setValue(value);
  };

  let fileUrl = suratRow.data['File PDF (URL)'] || '';
  if (formData.fileBase64 && formData.fileFileName) {
    fileUrl = simpanLampiranPDF(formData.fileBase64, formData.fileFileName, nomorLengkap.replace(/\//g, '_'));
  }

  const status = formData.status === 'Konsep' ? 'Konsep' : 'Diterbitkan';

  setCell('Perihal', formData.perihal || '');
  setCell('Tujuan', formData.tujuan || '');
  setCell('File PDF (URL)', fileUrl);
  setCell('Status', status);

  return { success: true, nomorLengkap, status };
}