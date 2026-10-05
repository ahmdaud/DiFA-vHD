/**
 * ============================================================
 *  MODUL LAPORAN BULANAN (PRD Bab 7)
 *  Generasi Laporan Rekapitulasi Surat Eksekutif ke Google Slides
 *  & Export ke PDF / PPTX.
 * ============================================================
 */

/** Cek apakah kode klasifikasi surat cocok dengan filter (mendukung kode induk, mis. 'TI' cocok utk 'TI.01'). */
function cocokKlasifikasi_(kodeSurat, filterKode) {
  if (!filterKode || filterKode === 'Semua') return true;
  if (!kodeSurat) return false;
  const k = kodeSurat.toString();
  return k === filterKode || k.indexOf(filterKode + '.') === 0;
}

/**
 * Mengambil rekap Surat Masuk & Surat Keluar sesuai filter.
 * filters = { bulan: 0-12 (0 = semua bulan), tahun: number, kodeKlasifikasi: string }
 */
function getLaporanBulanan(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bulan = Number(filters.bulan) || 0;
  const tahun = Number(filters.tahun);
  const kodeKlasifikasi = filters.kodeKlasifikasi || 'Semua';

  const masukAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_MASUK);
  const keluarAll = readSheetAsObjects(ss, SHEET_NAMES.SURAT_KELUAR);

  function cocokTanggal(val) {
    if (!val) return false;
    const d = new Date(val);
    if (isNaN(d)) return false;
    if (d.getFullYear() !== tahun) return false;
    if (bulan && (d.getMonth() + 1) !== bulan) return false;
    return true;
  }

  const masuk = masukAll.filter(r => cocokTanggal(r['Diterima Tanggal']) && cocokKlasifikasi_(r['Kode Klasifikasi'], kodeKlasifikasi));
  const keluar = keluarAll.filter(r => cocokTanggal(r['Tanggal']) && cocokKlasifikasi_(r['Kode Klasifikasi'], kodeKlasifikasi));

  const klasCount = {};
  masuk.concat(keluar).forEach(r => {
    const kode = r['Kode Klasifikasi'] || 'Lainnya';
    klasCount[kode] = (klasCount[kode] || 0) + 1;
  });

  return {
    masuk: masuk,
    keluar: keluar,
    totalMasuk: masuk.length,
    totalKeluar: keluar.length,
    klasifikasi: klasCount
  };
}

const NAMA_BULAN_LAPORAN = ['Semua Bulan', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

/**
 * Membangun Presentasi Google Slides Laporan Rekapitulasi secara Otomatis.
 */
function buildLaporanSlides_(filters) {
  const data = getLaporanBulanan(filters);
  const periode = NAMA_BULAN_LAPORAN[Number(filters.bulan) || 0].toUpperCase() + ' ' + filters.tahun;
  const namaFile = 'Laporan_Rekapitulasi_Surat_' + periode.replace(/\s+/g, '_');

  // 1. Buat File Google Slides Baru
  const presentation = SlidesApp.create(namaFile);
  const slides = presentation.getSlides();

  // --- SLIDE 1: COVER SLIDE ---
  const slideCover = slides[0];
  slideCover.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 405)
            .getFill().setSolidFill('#0f3d5c'); // Latar belakang Navy Blue

  const titleBox = slideCover.insertTextBox('LAPORAN REKAPITULASI SURAT', 50, 130, 620, 80);
  titleBox.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor('#ffffff');

  const subTitleBox = slideCover.insertTextBox('PERIODE: ' + periode, 50, 200, 620, 50);
  subTitleBox.getText().getTextStyle().setFontSize(16).setForegroundColor('#cccccc');


  // --- SLIDE 2: EXECUTIVE SUMMARY & KPI CARDS ---
  const slideSummary = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  
  // Header Slide
  const headBox2 = slideSummary.insertTextBox('RINGKASAN EKSEKUTIF', 40, 30, 640, 40);
  headBox2.getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor('#0f3d5c');

  // KPI Card 1: Surat Masuk
  const card1 = slideSummary.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 90, 200, 110);
  card1.getFill().setSolidFill('#f0f4f8');
  card1.getBorder().setWeight(1).getLineFill().setSolidFill('#cccccc');
  const t1 = slideSummary.insertTextBox('SURAT MASUK\n\n' + data.totalMasuk, 50, 100, 180, 90);
  t1.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor('#0f3d5c');

  // KPI Card 2: Surat Keluar
  const card2 = slideSummary.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 260, 90, 200, 110);
  card2.getFill().setSolidFill('#f0f4f8');
  card2.getBorder().setWeight(1).getLineFill().setSolidFill('#cccccc');
  const t2 = slideSummary.insertTextBox('SURAT KELUAR\n\n' + data.totalKeluar, 270, 100, 180, 90);
  t2.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor('#0f3d5c');

  // KPI Card 3: Total Keseluruhan
  const card3 = slideSummary.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 480, 90, 200, 110);
  card3.getFill().setSolidFill('#0f3d5c');
  const t3 = slideSummary.insertTextBox('TOTAL NASKAH\n\n' + (data.totalMasuk + data.totalKeluar), 490, 100, 180, 90);
  t3.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor('#ffffff');

  // Tabel Distribusi Klasifikasi Singkat
  const klasKeys = Object.keys(data.klasifikasi);
  if (klasKeys.length > 0) {
    const rows = Math.min(klasKeys.length + 1, 6);
    const table = slideSummary.insertTable(rows, 2, 40, 220, 640, 140);
    table.getCell(0, 0).getText().setText('Kode Klasifikasi').getTextStyle().setBold(true).setFontSize(11);
    table.getCell(0, 1).getText().setText('Jumlah Surat').getTextStyle().setBold(true).setFontSize(11);
    
    for (let i = 0; i < rows - 1; i++) {
      const k = klasKeys[i];
      table.getCell(i + 1, 0).getText().setText(k).getTextStyle().setFontSize(10);
      table.getCell(i + 1, 1).getText().setText(data.klasifikasi[k].toString()).getTextStyle().setFontSize(10);
    }
  }


  // --- SLIDE 3: DETAIL SURAT MASUK (5 TERBARU) ---
  const slideMasuk = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const headBox3 = slideMasuk.insertTextBox('RINGKASAN SURAT MASUK TERBARU', 40, 30, 640, 40);
  headBox3.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor('#0f3d5c');

  const topMasuk = data.masuk.slice(0, 5);
  const rowsMasuk = topMasuk.length + 1;
  const tableMasuk = slideMasuk.insertTable(rowsMasuk, 5, 40, 80, 640, 270);
  
  const hMasuk = ['No. Agenda', 'Tgl Diterima', 'Dari', 'Hal', 'Status'];
  hMasuk.forEach((h, idx) => {
    tableMasuk.getCell(0, idx).getText().setText(h).getTextStyle().setBold(true).setFontSize(10).setForegroundColor('#0f3d5c');
  });

  topMasuk.forEach((r, rowIdx) => {
    tableMasuk.getCell(rowIdx + 1, 0).getText().setText(r['No. Agenda'] || '-').getTextStyle().setFontSize(9);
    tableMasuk.getCell(rowIdx + 1, 1).getText().setText(r['Diterima Tanggal'] ? Utilities.formatDate(new Date(r['Diterima Tanggal']), 'Asia/Jakarta', 'dd/MM/yyyy') : '-').getTextStyle().setFontSize(9);
    tableMasuk.getCell(rowIdx + 1, 2).getText().setText(r['Dari'] || '-').getTextStyle().setFontSize(9);
    tableMasuk.getCell(rowIdx + 1, 3).getText().setText(r['Hal'] || '-').getTextStyle().setFontSize(9);
    tableMasuk.getCell(rowIdx + 1, 4).getText().setText(r['Status'] || '-').getTextStyle().setFontSize(9);
  });


  // --- SLIDE 4: DETAIL SURAT KELUAR (5 TERBARU) ---
  const slideKeluar = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const headBox4 = slideKeluar.insertTextBox('RINGKASAN SURAT KELUAR TERBARU', 40, 30, 640, 40);
  headBox4.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor('#0f3d5c');

  const topKeluar = data.keluar.slice(0, 5);
  const rowsKeluar = topKeluar.length + 1;
  const tableKeluar = slideKeluar.insertTable(rowsKeluar, 5, 40, 80, 640, 270);

  const hKeluar = ['No. Surat', 'Tanggal', 'Tujuan', 'Perihal', 'Status'];
  hKeluar.forEach((h, idx) => {
    tableKeluar.getCell(0, idx).getText().setText(h).getTextStyle().setBold(true).setFontSize(10).setForegroundColor('#0f3d5c');
  });

  topKeluar.forEach((r, rowIdx) => {
    tableKeluar.getCell(rowIdx + 1, 0).getText().setText(r['Nomor Surat Lengkap'] || '-').getTextStyle().setFontSize(9);
    tableKeluar.getCell(rowIdx + 1, 1).getText().setText(r['Tanggal'] ? Utilities.formatDate(new Date(r['Tanggal']), 'Asia/Jakarta', 'dd/MM/yyyy') : '-').getTextStyle().setFontSize(9);
    tableKeluar.getCell(rowIdx + 1, 2).getText().setText(r['Tujuan'] || '-').getTextStyle().setFontSize(9);
    tableKeluar.getCell(rowIdx + 1, 3).getText().setText(r['Perihal'] || '-').getTextStyle().setFontSize(9); // Sudah diperbaiki
    tableKeluar.getCell(rowIdx + 1, 4).getText().setText(r['Status'] || '-').getTextStyle().setFontSize(9);
  });

  presentation.saveAndClose();
  return { id: presentation.getId(), name: namaFile };
}

/**
 * Meng-export Laporan Slide Deck ke format PDF
 */
function exportLaporanPDF(filters) {
  const built = buildLaporanSlides_(filters);
  const file = DriveApp.getFileById(built.id);
  const pdfBlob = file.getAs(MimeType.PDF).setName(built.name + '.pdf');
  
  const folder = getOrCreateFolder('Laporan Bulanan Export');
  const outFile = folder.createFile(pdfBlob);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Hapus file slide sementara di Drive
  file.setTrashed(true);

  return { success: true, url: outFile.getUrl() };
}

/**
 * Membagikan Link Google Slides (Bisa diedit/dipresentasikan langsung)
 */
function exportLaporanSlides(filters) {
  const built = buildLaporanSlides_(filters);
  const file = DriveApp.getFileById(built.id);
  
  const folder = getOrCreateFolder('Laporan Bulanan Export');
  file.moveTo(folder);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { success: true, url: file.getUrl() };
}