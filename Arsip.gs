/**
 * ============================================================
 *  MODUL ARSIP & BACKUP TAHUNAN (PRD Bab 8)
 *  - Backup manual: unduh data + lampiran PDF tahun tertentu (.zip)
 *  - Reset tahunan: arsipkan data tahun lama ke sheet histori
 *    read-only, nomor urut otomatis reset (lihat getNextAgendaNumber
 *    di InboundParser.gs & getNextNomorUrutKeluar di SuratKeluar.gs
 *    yang sudah di-scope per tahun).
 * ============================================================
 */

function extractDriveFileId_(url) {
  if (!url) return null;
  const m = url.toString().match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function tahunDariNilai_(val, tahun) {
  if (!val) return false;
  const d = new Date(val);
  return !isNaN(d) && d.getFullYear() === Number(tahun);
}

/* ============================================================
 *  RINGKASAN ARSIP (dipanggil dari menu Arsip & Backup)
 * ============================================================ */
function getArsipInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tahunSekarang = new Date().getFullYear();

  const masukAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const keluarAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_KELUAR);

  const jmlMasukTahunIni = masukAll.filter(r => tahunDariNilai_(r['Diterima Tanggal'], tahunSekarang)).length;
  const jmlKeluarTahunIni = keluarAll.filter(r => tahunDariNilai_(r['Tanggal'], tahunSekarang)).length;

  const folder = getOrCreateFolder('Backup Tahunan');
  const files = folder.getFiles();
  const daftarBackup = [];
  while (files.hasNext()) {
    const f = files.next();
    daftarBackup.push({ nama: f.getName(), url: f.getUrl(), tanggal: f.getDateCreated().toISOString() });
  }
  daftarBackup.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

  const histori = [SHEET_NAMES.SURAT_MASUK, SHEET_NAMES.SURAT_KELUAR, SHEET_NAMES.DISPOSISI].map(nama => {
    const sheetHistori = ss.getSheetByName('Histori_' + nama);
    return { sheet: nama, jumlahBaris: sheetHistori ? Math.max(0, sheetHistori.getLastRow() - 1) : 0 };
  });

  return {
    tahunSekarang: tahunSekarang,
    jmlMasukTahunIni: jmlMasukTahunIni,
    jmlKeluarTahunIni: jmlKeluarTahunIni,
    daftarBackup: daftarBackup.slice(0, 10),
    histori: histori
  };
}

/* ============================================================
 *  BACKUP MANUAL (.zip berisi data + lampiran PDF)
 * ============================================================ */
function buildBackupSpreadsheet_(tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masukAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const keluarAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_KELUAR);
  const disposisiAll = readSheetAsObjects(ss, SHEET_NAMES.DISPOSISI);

  const masuk = masukAll.filter(r => tahunDariNilai_(r['Diterima Tanggal'], tahun));
  const keluar = keluarAll.filter(r => tahunDariNilai_(r['Tanggal'], tahun));
  const noAgendaTahunIni = {};
  masuk.forEach(r => { noAgendaTahunIni[r['No. Agenda']] = true; });
  const disposisi = disposisiAll.filter(r => noAgendaTahunIni[r['No. Agenda Surat']]);

  const judul = 'Backup Data Persuratan ' + tahun;
  const ssTemp = SpreadsheetApp.create(judul);

  const sheetMasuk = ssTemp.getSheets()[0];
  sheetMasuk.setName('Surat Masuk');
  if (masuk.length) {
    const headerMasuk = Object.keys(masuk[0]);
    sheetMasuk.appendRow(headerMasuk);
    sheetMasuk.getRange(1, 1, 1, headerMasuk.length).setFontWeight('bold').setBackground('#0f3d5c').setFontColor('#ffffff');
    masuk.forEach(r => sheetMasuk.appendRow(headerMasuk.map(h => r[h])));
  }

  const sheetKeluar = ssTemp.insertSheet('Surat Keluar');
  if (keluar.length) {
    const headerKeluar = Object.keys(keluar[0]);
    sheetKeluar.appendRow(headerKeluar);
    sheetKeluar.getRange(1, 1, 1, headerKeluar.length).setFontWeight('bold').setBackground('#0f3d5c').setFontColor('#ffffff');
    keluar.forEach(r => sheetKeluar.appendRow(headerKeluar.map(h => r[h])));
  }

  const sheetDisposisi = ssTemp.insertSheet('Disposisi');
  if (disposisi.length) {
    const headerDisposisi = Object.keys(disposisi[0]);
    sheetDisposisi.appendRow(headerDisposisi);
    sheetDisposisi.getRange(1, 1, 1, headerDisposisi.length).setFontWeight('bold').setBackground('#0f3d5c').setFontColor('#ffffff');
    disposisi.forEach(r => sheetDisposisi.appendRow(headerDisposisi.map(h => r[h])));
  }

  SpreadsheetApp.flush();
  return { id: ssTemp.getId(), masuk: masuk, keluar: keluar };
}

/** Menghasilkan file .zip berisi data (xlsx) + seluruh lampiran PDF tahun terkait. */
function jalankanBackupManual(tahun) {
  const built = buildBackupSpreadsheet_(tahun);
  const blobs = [];

  const dataFile = DriveApp.getFileById(built.id);
  blobs.push(dataFile.getAs(MimeType.MICROSOFT_EXCEL).setName('Data-Persuratan-' + tahun + '.xlsx'));
  dataFile.setTrashed(true);

  let jmlLampiran = 0;
  const urlLampiran = [];
  built.masuk.forEach(r => { if (r['Lampiran (URL)']) urlLampiran.push(r['Lampiran (URL)']); });
  built.keluar.forEach(r => { if (r['File PDF (URL)']) urlLampiran.push(r['File PDF (URL)']); });

  urlLampiran.forEach(url => {
    try {
      const fileId = extractDriveFileId_(url);
      if (fileId) {
        blobs.push(DriveApp.getFileById(fileId).getBlob());
        jmlLampiran++;
      }
    } catch (e) {
      Logger.log('Gagal ambil lampiran untuk backup: ' + url + ' - ' + e);
    }
  });

  const zipBlob = Utilities.zip(blobs, 'Backup-Persuratan-' + tahun + '.zip');
  const folder = getOrCreateFolder('Backup Tahunan');
  const zipFile = folder.createFile(zipBlob);
  zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { success: true, url: zipFile.getUrl(), jmlLampiran: jmlLampiran };
}

/* ============================================================
 *  RESET TAHUNAN: arsipkan ke sheet histori read-only
 * ============================================================ */

/** Memindahkan baris tahun tertentu dari sheet aktif ke sheet Histori_<nama>. Mengembalikan jumlah baris dipindah. */
function arsipkanTahun_(sheetNamaAktif, kolomTanggal, tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAktif = ss.getSheetByName(sheetNamaAktif);
  if (!sheetAktif || sheetAktif.getLastRow() < 2) return 0;

  const values = sheetAktif.getDataRange().getValues();
  const headers = values[0];
  const tglCol = headers.indexOf(kolomTanggal);
  const baris = values.slice(1);

  const untukArsip = [];
  const untukDisimpan = [];
  baris.forEach(row => {
    const tglVal = row[tglCol];
    const y = tglVal ? new Date(tglVal).getFullYear() : null;
    if (y === Number(tahun)) untukArsip.push(row);
    else untukDisimpan.push(row);
  });

  if (untukArsip.length === 0) return 0;

  const namaHistori = 'Histori_' + sheetNamaAktif;
  let sheetHistori = ss.getSheetByName(namaHistori);
  if (!sheetHistori) {
    sheetHistori = ss.insertSheet(namaHistori);
    sheetHistori.appendRow(headers);
    sheetHistori.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#5a6b7d').setFontColor('#ffffff');
    sheetHistori.setFrozenRows(1);
  }
  sheetHistori.getRange(sheetHistori.getLastRow() + 1, 1, untukArsip.length, headers.length).setValues(untukArsip);

  // Kosongkan sheet aktif, isi ulang hanya dengan baris yang TIDAK diarsipkan
  sheetAktif.getRange(2, 1, sheetAktif.getLastRow() - 1, headers.length).clearContent();
  if (untukDisimpan.length > 0) {
    sheetAktif.getRange(2, 1, untukDisimpan.length, headers.length).setValues(untukDisimpan);
  }

  // Proteksi sheet histori sebagai read-only (peringatan saat ada yang coba edit)
  try {
    const existing = sheetHistori.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (existing.length === 0) {
      sheetHistori.protect()
        .setDescription('Arsip Historis ' + tahun + ' - Read Only (dibuat otomatis oleh sistem)')
        .setWarningOnly(true);
    }
  } catch (e) {
    Logger.log('Gagal memproteksi sheet histori: ' + e);
  }

  return untukArsip.length;
}

/** Menjalankan reset tahunan untuk tahun tertentu (default: tahun lalu). Bisa dipanggil manual dari dashboard atau otomatis via trigger. */
function jalankanResetTahunan(tahun) {
  const tahunDitutup = tahun ? Number(tahun) : (new Date().getFullYear() - 1);

  const jmlMasuk = arsipkanTahun_(SHEET_NAMES.SURAT_MASUK, 'Diterima Tanggal', tahunDitutup);
  const jmlKeluar = arsipkanTahun_(SHEET_NAMES.SURAT_KELUAR, 'Tanggal', tahunDitutup);
  const jmlDisposisi = arsipkanTahun_(SHEET_NAMES.DISPOSISI, 'Tanggal Disposisi', tahunDitutup);

  return {
    success: true,
    tahunDiarsipkan: tahunDitutup,
    jmlMasuk: jmlMasuk,
    jmlKeluar: jmlKeluar,
    jmlDisposisi: jmlDisposisi
  };
}

/* ============================================================
 *  TRIGGER OTOMATIS SETIAP 1 JANUARI
 * ============================================================ */

/** Jalankan SEKALI dari editor untuk mengaktifkan pengecekan reset tahunan otomatis. */
function setupResetTahunanTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cekDanJalankanResetTahunanOtomatis') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cekDanJalankanResetTahunanOtomatis')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  SpreadsheetApp.getUi().alert('Trigger reset tahunan otomatis aktif — sistem akan mengecek setiap hari jam 01:00, dan menjalankan arsip otomatis setiap tanggal 1 Januari.');
}

/** Dipanggil harian oleh trigger. Hanya benar-benar jalan sekali setiap 1 Januari. */
function cekDanJalankanResetTahunanOtomatis() {
  const now = new Date();
  if (now.getMonth() !== 0 || now.getDate() !== 1) return; // hanya tanggal 1 Januari

  const props = PropertiesService.getScriptProperties();
  const tahunIni = now.getFullYear();
  const tahunTerakhirReset = Number(props.getProperty('LAST_RESET_YEAR') || 0);
  if (tahunTerakhirReset === tahunIni) return; // sudah pernah jalan tahun ini

  jalankanResetTahunan(tahunIni - 1);
  props.setProperty('LAST_RESET_YEAR', String(tahunIni));
}