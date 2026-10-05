/**
 * ============================================================
 *  APLIKASI PERSURATAN (E-OFFICE) - BACKEND
 *  Pengadilan - W8-A10
 * ============================================================
 *  Struktur Sheet yang dipakai sebagai "database":
 *   - SuratMasuk
 *   - SuratKeluar
 *   - Disposisi
 *   - MasterPejabat
 *   - MasterKlasifikasi
 *
 *  Jalankan fungsi initializeSheets() sekali dari editor Apps
 *  Script (menu Run) untuk membuat semua sheet & header otomatis.
 * ============================================================
 */

const APP_TITLE = 'E-Office Persuratan | W8-A10';

const SHEET_NAMES = {
  SURAT_MASUK: 'SuratMasuk',
  SURAT_KELUAR: 'SuratKeluar',
  DISPOSISI: 'Disposisi',
  MASTER_PEJABAT: 'MasterPejabat',
  MASTER_KLASIFIKASI: 'MasterKlasifikasi'
};

/* ============================================================
 *  ENTRY POINT WEB APP
 * ============================================================ */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';

  // Fallback bypass link dari email notifikasi disposisi (tanpa login)
  if (page === 'disposisi-web') {
    return renderDisposisiWebForm(e.parameter.kode || '');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.initialPage = page;
  return template
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Helper untuk include file HTML/CSS/JS terpisah ke dalam Index.html */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
 *  SETUP AWAL - Jalankan sekali dari editor (Run > initializeSheets)
 * ============================================================ */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsConfig = {
    [SHEET_NAMES.SURAT_MASUK]: [
      'No. Agenda', 'Diterima Tanggal', 'Nomor Naskah Dinas', 'Tanggal Naskah Dinas',
      'Dari', 'Sifat', 'Jenis', 'Index', 'Hal', 'Status Naskah',
      'Kode Klasifikasi', 'Pimpinan Tujuan', 'Email Pimpinan', 'Kode Tracking',
      'Lampiran (URL)', 'Status', 'Petunjuk Disposisi', 'Disposisi Kepada',
      'Catatan Disposisi', 'Tanggal Disposisi', 'Metode Disposisi', 'Status Diteruskan',
      'Pemegang Saat Ini', 'Email Pemegang Saat Ini'
    ],
    [SHEET_NAMES.SURAT_KELUAR]: [
      'No. Urut', 'Nomor Surat Lengkap', 'Tanggal', 'Kode Penandatangan',
      'Pejabat Penandatangan', 'Kode Klasifikasi', 'Perihal', 'Tujuan',
      'File PDF (URL)', 'Dibuat Oleh', 'Status'
    ],
    [SHEET_NAMES.DISPOSISI]: [
      'Kode Tracking', 'No. Agenda Surat', 'Pimpinan', 'Petunjuk Disposisi',
      'Diteruskan Kepada', 'Catatan', 'Tanggal Disposisi', 'Status Tindak Lanjut',
      'Tanggal Selesai', 'Email Diteruskan'
    ],
    [SHEET_NAMES.MASTER_PEJABAT]: [
      'Jabatan', 'Kode', 'Nama Pejabat', 'Email', 'No. WhatsApp'
    ],
    [SHEET_NAMES.MASTER_KLASIFIKASI]: [
      'Kode Klasifikasi', 'Nama Klasifikasi / Subjek', 'Keterangan / Uraian'
    ]
  };

  Object.keys(sheetsConfig).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const headers = sheetsConfig[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f3d5c')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  });

  seedMasterData(ss);

  // Hapus sheet default "Sheet1" jika masih ada dan kosong
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert('Setup selesai! Semua sheet & data master berhasil dibuat.');
}

function seedMasterData(ss) {
  const pejabatSheet = ss.getSheetByName(SHEET_NAMES.MASTER_PEJABAT);
  if (pejabatSheet.getLastRow() < 2) {
    const pejabat = [
      ['Ketua Pengadilan', 'KPA', '', '', ''],
      ['Wakil Ketua Pengadilan', 'KPA', '', '', ''],
      ['Hakim', 'HKM', '', '', ''],
      ['Panitera', 'PAN.PA', '', '', ''],
      ['Sekretaris', 'SEK.PA', '', '', ''],
      ['Panitera Muda Permohonan', 'PAN.01', '', '', ''],
      ['Panitera Muda Hukum', 'PAN.02', '', '', ''],
      ['Panitera Muda Gugatan', 'PAN.03', '', '', ''],
      ['Kasubbag Perencanaan, Pelaporan & TI', 'SEK.01', '', '', ''],
      ['Kasubbag Kepegawaian & ORTALA', 'SEK.02', '', '', ''],
      ['Kasubbag Umum & Keuangan', 'SEK.03', '', '', '']
    ];
    pejabatSheet.getRange(2, 1, pejabat.length, 5).setValues(pejabat);
  }

  const klasSheet = ss.getSheetByName(SHEET_NAMES.MASTER_KLASIFIKASI);
  if (klasSheet.getLastRow() < 2) {
    const klas = [
      ['HK', 'Hukum', 'Klasifikasi umum bidang hukum'],
      ['HK.01', 'Peraturan', 'Sub klasifikasi Hukum — Peraturan'],
      ['HK.02', 'Perkara', 'Sub klasifikasi Hukum — Perkara'],
      ['HM', 'Hubungan Masyarakat', ''],
      ['KP', 'Kepegawaian', 'Klasifikasi umum bidang kepegawaian'],
      ['KP.01', 'Formasi', 'Sub klasifikasi Kepegawaian — Formasi'],
      ['KP.02', 'Mutasi', 'Sub klasifikasi Kepegawaian — Mutasi'],
      ['PL', 'Perlengkapan / Barang Milik Negara (BMN)', ''],
      ['RT', 'Rumah Tangga', ''],
      ['TI', 'Teknologi Informasi', 'Klasifikasi umum bidang TI'],
      ['TI.01', 'Jaringan', 'Sub klasifikasi TI — Jaringan'],
      ['TI.02', 'Aplikasi', 'Sub klasifikasi TI — Aplikasi'],
      ['KU', 'Keuangan', 'Klasifikasi umum bidang keuangan'],
      ['KU.01', 'Anggaran', 'Sub klasifikasi Keuangan — Anggaran'],
      ['KU.02', 'Gaji', 'Sub klasifikasi Keuangan — Gaji']
    ];
    klasSheet.getRange(2, 1, klas.length, 3).setValues(klas);
  }
}

/* ============================================================
 *  DATA UNTUK DASHBOARD
 * ============================================================ */
function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const thisMonth = now.getMonth();
  const thisYear = now.getFullYear();

  const masukData = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const keluarData = readSheetAsObjects(ss, SHEET_NAMES.SURAT_KELUAR);

  const isThisMonth = (dateVal) => {
    if (!dateVal) return false;
    const d = new Date(dateVal);
    return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
  };

  const masukBulanIni = masukData.filter(r => isThisMonth(r['Diterima Tanggal']));
  const keluarBulanIni = keluarData.filter(r => isThisMonth(r['Tanggal']));

  const pending = masukData.filter(r =>
    !r['Status'] || (r['Status'] !== 'Selesai' && r['Status'] !== 'Closed')
  );
  const didisposisikan = masukData.filter(r => r['Status'] === 'Didisposisikan');
  const selesai = masukData.filter(r => r['Status'] === 'Selesai' || r['Status'] === 'Closed');

  // Distribusi klasifikasi (gabungan masuk + keluar) untuk chart
  const klasCount = {};
  masukData.concat(keluarData).forEach(r => {
    const kode = (r['Kode Klasifikasi'] || 'Lainnya').toString().split('.')[0];
    klasCount[kode] = (klasCount[kode] || 0) + 1;
  });

  // Tren 6 bulan terakhir
  const trend = buildMonthlyTrend(masukData, keluarData);

  return {
    totalMasukBulanIni: masukBulanIni.length,
    totalKeluarBulanIni: keluarBulanIni.length,
    totalPending: pending.length,
    totalDidisposisikan: didisposisikan.length,
    totalSelesai: selesai.length,
    klasifikasi: klasCount,
    trend: trend,
    recentSuratMasuk: masukData.slice(-8).reverse(),
    recentSuratKeluar: keluarData.slice(-6).reverse()
  };
}

function buildMonthlyTrend(masukData, keluarData) {
  const bulanLabel = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const now = new Date();

  // Kelompokkan sekali saja (satu kali pass per dataset) menjadi map
  // "YYYY-M" -> jumlah, bukan memfilter ulang seluruh data untuk tiap bulan.
  function groupByYearMonth(rows, tanggalField) {
    const map = {};
    rows.forEach(r => {
      const val = r[tanggalField];
      if (!val) return;
      const d = new Date(val);
      if (isNaN(d)) return;
      const key = d.getFullYear() + '-' + d.getMonth();
      map[key] = (map[key] || 0) + 1;
    });
    return map;
  }

  const masukMap = groupByYearMonth(masukData, 'Diterima Tanggal');
  const keluarMap = groupByYearMonth(keluarData, 'Tanggal');

  const labels = [];
  const masukCounts = [];
  const keluarCounts = [];

  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = d.getFullYear() + '-' + d.getMonth();
    labels.push(bulanLabel[d.getMonth()]);
    masukCounts.push(masukMap[key] || 0);
    keluarCounts.push(keluarMap[key] || 0);
  }

  return { labels, masukCounts, keluarCounts };
}

/**
 * Sama seperti readSheetAsObjects, tapi hasilnya disimpan sebentar di
 * CacheService agar pemanggilan berulang dalam rentang waktu singkat
 * (mis. beberapa kali klik menu berturut-turut) tidak perlu membaca
 * ulang seluruh sheet. Cocok untuk data yang jarang berubah (Master
 * Pejabat, Master Klasifikasi). Jangan pakai untuk sheet yang sering
 * ditulis tanpa invalidasi eksplisit, karena bisa menampilkan data basi.
 */
function readSheetAsObjectsCached(ss, sheetName, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const key = 'sheet_' + sheetName;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const data = readSheetAsObjects(ss, sheetName);
  try {
    cache.put(key, JSON.stringify(data), ttlSeconds || 20);
  } catch (e) {
    // Data terlalu besar untuk cache (>100KB) — lewati saja, tidak fatal.
  }
  return data;
}

/** Membaca sebuah sheet dan mengembalikan array of objects berbasis header baris 1 */
function readSheetAsObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const val = row[i];
        obj[h] = (val instanceof Date) ? val.toISOString() : val;
      });
      return obj;
    });
}

/* ============================================================
 *  DATA UNTUK MODUL LAIN (dipanggil dari sidebar menu)
 * ============================================================ */

/** Cek apakah sebuah nilai tanggal jatuh pada tahun tertentu. */
function cocokTahun_(val, tahun) {
  if (!val) return false;
  const d = new Date(val);
  return !isNaN(d) && d.getFullYear() === Number(tahun);
}

/** Mengumpulkan daftar tahun unik dari sebuah kolom tanggal, terbaru dulu. Tahun berjalan selalu disertakan. */
function getTahunUnik_(rows, kolomTanggal) {
  const set = {};
  rows.forEach(r => {
    const val = r[kolomTanggal];
    if (!val) return;
    const d = new Date(val);
    if (!isNaN(d)) set[d.getFullYear()] = true;
  });
  set[new Date().getFullYear()] = true;
  return Object.keys(set).map(Number).sort((a, b) => b - a);
}

/**
 * Daftar Surat Masuk, difilter per tahun (default: tahun berjalan) agar
 * tidak mengirim & merender seluruh riwayat surat sekaligus ke browser.
 * Mengembalikan { rows, tahunDipilih, tahunTersedia } supaya dropdown
 * filter tahun di client bisa dibangun dari satu kali panggilan saja.
 */
function getSuratMasukList(tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const tahunTersedia = getTahunUnik_(all, 'Diterima Tanggal');
  const tahunDipilih = tahun ? Number(tahun) : new Date().getFullYear();
  const rows = all.filter(r => cocokTahun_(r['Diterima Tanggal'], tahunDipilih)).reverse();
  return { rows: rows, tahunDipilih: tahunDipilih, tahunTersedia: tahunTersedia };
}

/**
 * Daftar Surat Keluar, difilter per tahun (default: tahun berjalan) —
 * lihat catatan pada getSuratMasukList().
 */
function getSuratKeluarList(tahun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = readSheetAsObjects(ss, SHEET_NAMES.SURAT_KELUAR);
  const tahunTersedia = getTahunUnik_(all, 'Tanggal');
  const tahunDipilih = tahun ? Number(tahun) : new Date().getFullYear();
  const rows = all.filter(r => cocokTahun_(r['Tanggal'], tahunDipilih)).reverse();
  return { rows: rows, tahunDipilih: tahunDipilih, tahunTersedia: tahunTersedia };
}

function getMasterPejabat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return readSheetAsObjectsCached(ss, SHEET_NAMES.MASTER_PEJABAT, 30);
}

function getMasterKlasifikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return readSheetAsObjectsCached(ss, SHEET_NAMES.MASTER_KLASIFIKASI, 30);
}

/**
 * Mengambil Master Pejabat & Master Klasifikasi sekaligus dalam SATU
 * pemanggilan google.script.run — dipakai untuk pre-load di awal buka
 * aplikasi supaya dropdown di modal manapun sudah siap tanpa harus
 * menunggu round-trip baru saat modal dibuka.
 */
function getMasterDataGabungan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    pejabat: readSheetAsObjectsCached(ss, SHEET_NAMES.MASTER_PEJABAT, 30),
    klasifikasi: readSheetAsObjectsCached(ss, SHEET_NAMES.MASTER_KLASIFIKASI, 30)
  };
}

function getDisposisiList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = readSheetAsObjects(ss, SHEET_NAMES.DISPOSISI);
  const suratList = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const suratMap = {};
  suratList.forEach(s => { suratMap[s['Kode Tracking']] = s; });

  // Surat lama (diinput sebelum alur "disposisi awal") belum punya baris di sheet Disposisi.
  // Dibuatkan baris awal virtual agar tetap tampil, bisa dilacak, & bisa dicetak.
  const sudahAda = {};
  rows.forEach(r => { sudahAda[r['Kode Tracking']] = true; });
  const barisAwalVirtual = suratList
    .filter(s => s['Kode Tracking'] && !sudahAda[s['Kode Tracking']])
    .map(s => ({
      'Kode Tracking': s['Kode Tracking'],
      'No. Agenda Surat': s['No. Agenda'],
      'Pimpinan': LABEL_INPUT_ADMIN,
      'Petunjuk Disposisi': '-',
      'Diteruskan Kepada': s['Pimpinan Tujuan'] || '-',
      'Catatan': '',
      'Tanggal Disposisi': s['Diterima Tanggal'],
      'Status Tindak Lanjut': s['Status'] === 'Selesai' ? 'Selesai' : STATUS_MENUNGGU,
      'Tanggal Selesai': '',
      'Email Diteruskan': s['Email Pimpinan'] || ''
    }));
  const semua = barisAwalVirtual.concat(rows);

  const lastIndexByKode = {};
  semua.forEach((r, idx) => { lastIndexByKode[r['Kode Tracking']] = idx; });

  return semua.map((r, idx) => {
    const surat = suratMap[r['Kode Tracking']] || {};
    return Object.assign({}, r, {
      'Pemegang Saat Ini': surat['Pemegang Saat Ini'] || surat['Pimpinan Tujuan'] || '-',
      'Status Surat': surat['Status'] || '-',
      'Adalah Hop Terakhir': lastIndexByKode[r['Kode Tracking']] === idx
    });
  }).reverse();
}