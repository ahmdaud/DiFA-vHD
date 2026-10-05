/**
 * ============================================================
 *  MODUL INBOUND EMAIL PARSING - DISPOSISI CEPAT
 *  Sesuai PRD Bab 3 (Alur Kerja), Bab 4 (Lembar Disposisi),
 *  Bab 5 (Tracking), Bab 9 (Keamanan & Mitigasi Risiko)
 * ============================================================
 */

const TRACKING_PREFIX = 'DISP';
const PROP_KEY_PROCESSED = 'PROCESSED_MSG_IDS';

// Baris pertama rantai disposisi dicatat otomatis saat admin menginput surat masuk
// ("Admin mengarahkan surat ke X"), sehingga rantai & lembar cetak lengkap sejak awal.
const LABEL_INPUT_ADMIN = 'Admin (Input Surat)';
const STATUS_MENUNGGU = 'Menunggu Disposisi';

/* ============================================================
 *  1. ADMIN INPUT SURAT MASUK -> KIRIM NOTIFIKASI KE PIMPINAN
 * ============================================================ */

/**
 * Dipanggil dari form input Surat Masuk (menu "Surat Masuk").
 * formData = {
 *   nomorNaskahDinas, tanggalNaskahDinas, dari, sifat, jenis, indexBidang,
 *   hal, statusNaskah, kodeKlasifikasi, pimpinanTujuan, emailPimpinan,
 *   lampiranBase64 (opsional), lampiranFileName (opsional)
 * }
 */
function tambahSuratMasuk(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_MASUK);

  // Tujuan awal bisa siapa saja di Data Master: Ketua / Wakil Ketua / Panitera / Sekretaris / Hakim / dst.
  const tujuanInput = (formData.pimpinanTujuan || '').toString().trim();
  if (!tujuanInput) throw new Error('Tujuan surat (Ditujukan Kepada) wajib dipilih.');

  const pejabatTujuan = findPejabatByJabatanText(tujuanInput);
  const jabatanTujuan = pejabatTujuan ? pejabatTujuan['Jabatan'] : tujuanInput;
  const emailTujuan = (formData.emailPimpinan || (pejabatTujuan && pejabatTujuan['Email']) || '').toString().trim();

  const noAgenda = getNextAgendaNumber(sheet);
  const diterimaTanggal = new Date();
  const kodeTracking = generateTrackingCode(noAgenda);

  let lampiranUrl = '';
  if (formData.lampiranBase64 && formData.lampiranFileName) {
    lampiranUrl = simpanLampiranPDF(formData.lampiranBase64, formData.lampiranFileName, kodeTracking);
  }

  // Urutan HARUS sama persis dengan header sheet SuratMasuk di Code.gs
  sheet.appendRow([
    noAgenda, diterimaTanggal, formData.nomorNaskahDinas || '', formData.tanggalNaskahDinas || '',
    formData.dari || '', formData.sifat || '', formData.jenis || '', formData.indexBidang || '',
    formData.hal || '', formData.statusNaskah || '', formData.kodeKlasifikasi || '',
    jabatanTujuan, emailTujuan, kodeTracking,
    lampiranUrl, 'Terkirim', '', '', '', '', '', '',
    jabatanTujuan, emailTujuan
  ]);

  // Catat langkah pertama rantai disposisi: admin mengarahkan surat ke tujuan awal.
  ss.getSheetByName(SHEET_NAMES.DISPOSISI).appendRow([
    kodeTracking, noAgenda, LABEL_INPUT_ADMIN, '-', jabatanTujuan, '',
    diterimaTanggal, STATUS_MENUNGGU, '', emailTujuan
  ]);

  const notif = kirimNotifikasiDisposisi({
    noAgenda, diterimaTanggal, nomorNaskahDinas: formData.nomorNaskahDinas,
    tanggalNaskahDinas: formData.tanggalNaskahDinas, dari: formData.dari, sifat: formData.sifat,
    jenis: formData.jenis, indexBidang: formData.indexBidang, hal: formData.hal,
    pimpinanTujuan: jabatanTujuan, emailPimpinan: emailTujuan, pejabatTujuan: pejabatTujuan,
    kodeTracking, lampiranUrl
  });

  return {
    success: true, kodeTracking, noAgenda, tujuan: jabatanTujuan,
    emailTerkirim: notif.emailTerkirim, waTerkirim: notif.waTerkirim
  };
}

/** Nomor agenda di-scope per tahun berjalan (reset otomatis tiap 1 Jan sesuai PRD Bab 8). */
function getNextAgendaNumber(sheet) {
  if (sheet.getLastRow() < 2) return 1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const noCol = headers.indexOf('No. Agenda');
  const tglCol = headers.indexOf('Diterima Tanggal');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const tahunIni = new Date().getFullYear();

  let max = 0;
  values.forEach(row => {
    const tglVal = row[tglCol];
    if (!tglVal) return;
    const y = (tglVal instanceof Date) ? tglVal.getFullYear() : new Date(tglVal).getFullYear();
    if (y === tahunIni) {
      const n = Number(row[noCol]) || 0;
      if (n > max) max = n;
    }
  });
  return max + 1;
}

/** Kode unik pelacakan, contoh: DISP-2026-09-001 (lihat PRD Bab 3 & 9) */
function generateTrackingCode(noAgenda) {
  const now = new Date();
  const y = now.getFullYear();
  const m = ('0' + (now.getMonth() + 1)).slice(-2);
  const seq = ('00' + noAgenda).slice(-3);
  return `${TRACKING_PREFIX}-${y}-${m}-${seq}`;
}

function simpanLampiranPDF(base64Data, fileName, kodeTracking) {
  const folder = getOrCreateFolder('Lampiran Surat Masuk');
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, 'application/pdf', `${kodeTracking}_${fileName}`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

/**
 * Mengirim email notifikasi disposisi ke pimpinan.
 * Subjek WAJIB memuat kode tracking dalam kurung siku — Gmail
 * mempertahankan subjek pada balasan ("Re: [DISP-...]") sehingga
 * parser bisa mengenali surat mana yang dibalas.
 */
function kirimNotifikasiDisposisi(s) {
  const subject = `[${s.kodeTracking}] Surat Masuk: ${s.hal}`;
  const bypassUrl = `${ScriptApp.getService().getUrl()}?page=disposisi-web&kode=${encodeURIComponent(s.kodeTracking)}`;

  const pejabatList = readSheetAsObjects(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.MASTER_PEJABAT);
  const daftarTujuan = pejabatList.map(p => p['Jabatan']).filter(Boolean).join(', ');
  const htmlBody = `
    <div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:auto;border:1px solid #e6ebf2;border-radius:10px;overflow:hidden;">
      <div style="background:#0f3d5c;color:#fff;padding:16px 22px;">
        <div style="font-size:15px;font-weight:700;">📨 Surat Masuk Baru — Menunggu Disposisi</div>
        <div style="font-size:12px;opacity:0.8;margin-top:2px;">Kode Tracking: ${s.kodeTracking}</div>
      </div>
      <div style="padding:20px 22px;color:#1a2b3c;font-size:13.5px;line-height:1.6;">
        <table style="width:100%;border-collapse:collapse;">
          <tr><td style="color:#6b7c93;width:150px;">Ditujukan Kepada</td><td><b>${s.pimpinanTujuan || '-'}</b></td></tr>
          <tr><td style="color:#6b7c93;">No. Agenda</td><td><b>${s.noAgenda}</b></td></tr>
          <tr><td style="color:#6b7c93;">Diterima Tanggal</td><td>${Utilities.formatDate(s.diterimaTanggal, Session.getScriptTimeZone(), 'dd MMMM yyyy')}</td></tr>
          <tr><td style="color:#6b7c93;">Nomor Naskah Dinas</td><td>${s.nomorNaskahDinas || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Tanggal Naskah Dinas</td><td>${s.tanggalNaskahDinas || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Dari</td><td>${s.dari || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Sifat</td><td>${s.sifat || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Jenis</td><td>${s.jenis || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Index</td><td>${s.indexBidang || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Hal</td><td><b>${s.hal || '-'}</b></td></tr>
          ${s.lampiranUrl ? `<tr><td style="color:#6b7c93;">Lampiran</td><td><a href="${s.lampiranUrl}" target="_blank">Lihat File PDF</a></td></tr>` : ''}
        </table>

        <p style="margin:20px 0 8px;">Untuk mendisposisikan surat ini, <b>cukup balas (Reply) email ini</b> — tidak perlu login. Isi 3 baris berikut <u>di atas kutipan email ini</u>, lalu klik <b>Send</b>:</p>

        <div style="background:#f4f7fb;border:1px dashed #c7d3e0;border-radius:8px;padding:14px 16px;font-family:Consolas,monospace;font-size:12.5px;white-space:pre-wrap;">Petunjuk Disposisi : (Setuju/Tolak/Selesaikan/Untuk diketahui)
Diteruskan Kepada  : (satu atau lebih jabatan, pisahkan dengan titik koma ; )
Catatan Tambahan   : (opsional)</div>

        <p style="margin:14px 0 6px;color:#6b7c93;font-size:12px;">Pilihan tujuan disposisi (boleh pilih lebih dari satu sekaligus, contoh: "Sekretaris; Kasubbag Umum & Keuangan"):</p>
        <div style="background:#eef2f7;border-radius:6px;padding:10px 14px;font-size:12px;color:#33475b;">${daftarTujuan || '(Data master pejabat belum diisi)'}</div>

        <p style="margin:18px 0 6px;color:#6b7c93;font-size:12px;">Tidak nyaman membalas email? Gunakan tautan alternatif berikut (tersedia kotak centang untuk pilih beberapa tujuan sekaligus):</p>
        <a href="${bypassUrl}" style="display:inline-block;background:#14b8a6;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-size:13px;font-weight:600;">Buka Formulir Disposisi Web →</a>
      </div>
      <div style="background:#f4f7fb;padding:12px 22px;font-size:11px;color:#8fa4b8;">
        Email ini dikirim otomatis oleh Sistem E-Office Persuratan W8-A10. Mohon tidak meneruskan email ini ke pihak lain.
      </div>
    </div>`;

  const plainBody = `Surat Masuk Baru [${s.kodeTracking}]\nDitujukan kepada: ${s.pimpinanTujuan || '-'}\nHal: ${s.hal}\nDari: ${s.dari}\n\n`
    + `Balas email ini dengan format:\nPetunjuk Disposisi : ...\nDiteruskan Kepada  : (pilih dari: ${daftarTujuan})\nCatatan Tambahan   : ...\n\n`
    + `Atau gunakan link berikut: ${bypassUrl}`;

  const hasil = { emailTerkirim: false, waTerkirim: false };

  // Email hanya dikirim jika alamatnya ada (penerima bisa saja hanya memakai WhatsApp).
  if (s.emailPimpinan) {
    try {
      GmailApp.sendEmail(s.emailPimpinan, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'E-Office Persuratan W8-A10'
      });
      hasil.emailTerkirim = true;
    } catch (e) {
      Logger.log('Gagal kirim email notifikasi surat masuk ke ' + s.pimpinanTujuan + ': ' + e);
    }
  }

  // WhatsApp (isi sama dengan email, versi teks polos) ke "No. WhatsApp" tujuan di Data Master.
  // Dibungkus try/catch agar kegagalan kirim tidak menggagalkan penyimpanan surat.
  try {
    const pejabatTujuan = s.pejabatTujuan || findPejabatByJabatanText(s.pimpinanTujuan);
    if (pejabatTujuan && pejabatTujuan['No. WhatsApp']) {
      const hasilWa = kirimWhatsApp_(pejabatTujuan['No. WhatsApp'], plainBody);
      hasil.waTerkirim = !!hasilWa.success;
      if (!hasilWa.success) Logger.log('Gagal kirim WhatsApp notifikasi surat masuk ke ' + s.pimpinanTujuan + ': ' + hasilWa.message);
    }
  } catch (e) {
    Logger.log('Gagal kirim WhatsApp notifikasi surat masuk: ' + e);
  }

  return hasil;
}

/* ============================================================
 *  2. TRIGGER TERJADWAL - MEMBACA BALASAN EMAIL PIMPINAN
 * ============================================================ */

/** Jalankan SEKALI dari editor Apps Script (Run) untuk mengaktifkan pengecekan otomatis tiap 5 menit. */
function setupInboundTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkInboundReplies') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkInboundReplies')
    .timeBased()
    .everyMinutes(5)
    .create();
  SpreadsheetApp.getUi().alert('Trigger inbound parsing aktif — sistem akan mengecek balasan email setiap 5 menit.');
}

/** Dipanggil otomatis oleh trigger terjadwal. Mencari balasan ber-kode tracking yang belum diproses. */
function checkInboundReplies() {
  const threads = GmailApp.search('subject:"[DISP-" is:unread in:inbox newer_than:14d');
  const processedIds = getProcessedMessageIds();

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const lastMsg = messages[messages.length - 1];
    const msgId = lastMsg.getId();

    if (processedIds.has(msgId) || !lastMsg.isUnread()) return;

    const subjectGabungan = thread.getFirstMessageSubject() + ' ' + lastMsg.getSubject();
    const kodeMatch = subjectGabungan.match(/DISP-\d{4}-\d{2}-\d{3}/);
    if (!kodeMatch) return;

    const kodeTracking = kodeMatch[0];
    const suratRow = findSuratByKode(kodeTracking);

    if (!suratRow) { lastMsg.markRead(); markProcessed(msgId); return; }
    if (suratRow.data['Status'] === 'Selesai') {
      // Surat sudah ditutup admin — tidak menerima disposisi lanjutan lagi.
      lastMsg.markRead(); markProcessed(msgId);
      return;
    }

    // --- Whitelisting Sender sesuai PRD Bab 9 ---
    // Dicocokkan terhadap DAFTAR pemegang surat saat ini (bisa lebih dari
    // satu orang jika disposisi difanout ke beberapa tujuan sekaligus),
    // supaya rantai berjenjang tervalidasi di setiap hop, bukan hanya hop pertama.
    const fromEmail = extractEmailAddress(lastMsg.getFrom());
    const daftarEmailSah = pisahkanEmail_(suratRow.data['Email Pemegang Saat Ini'] || suratRow.data['Email Pimpinan'] || '')
      .map(e => e.toLowerCase());
    if (daftarEmailSah.length && daftarEmailSah.indexOf(fromEmail.toLowerCase()) === -1) {
      Logger.log(`Disposisi DITOLAK - pengirim tidak terdaftar: ${fromEmail} untuk ${kodeTracking}`);
      lastMsg.markRead();
      markProcessed(msgId);
      return;
    }

    const replyText = extractReplyText(lastMsg.getPlainBody());
    const parsed = parseDisposisiText(replyText);

    // Kenali PEMEGANG mana yang membalas (berdasarkan email pengirim), supaya pada
    // disposisi paralel hanya pemegang tersebut yang "selesai bertindak".
    const namaPemegang = pisahkanTarget_(suratRow.data['Pemegang Saat Ini']);
    const emailPemegang = pisahkanSejajar_(suratRow.data['Email Pemegang Saat Ini']).map(e => e.toLowerCase());
    const idxPengirim = emailPemegang.indexOf(fromEmail.toLowerCase());
    const dariPengirim = idxPengirim > -1 ? namaPemegang[idxPengirim] : null;

    prosesDisposisi(suratRow, {
      petunjuk: parsed.petunjuk || '(tidak terbaca otomatis - cek email asli)',
      diteruskanKepada: parsed.diteruskanKepada || '-',
      catatan: parsed.catatan || replyText.substring(0, 300),
      metode: 'Inbound Email Parsing'
    }, dariPengirim);

    lastMsg.markRead();
    markProcessed(msgId);
  });
}

function findSuratByKode(kode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SURAT_MASUK);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const kodeCol = headers.indexOf('Kode Tracking');

  for (let i = 1; i < values.length; i++) {
    if (values[i][kodeCol] === kode) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = values[i][idx]);
      return { rowIndex: i + 1, data: obj, headers };
    }
  }
  return null;
}

/** Memotong teks balasan agar hanya mengambil isi baru, tanpa kutipan email asli. */
function extractReplyText(plainBody) {
  const quoteMarkers = [
    /\nOn .+wrote:/i,
    /\nPada .+menulis:/i,
    /\n>.*Surat Masuk Baru/i,
    /\n-{2,}\s*Forwarded message/i,
    /\n_{5,}/
  ];
  let text = plainBody;
  for (const marker of quoteMarkers) {
    const idx = text.search(marker);
    if (idx > -1) text = text.substring(0, idx);
  }
  return text.trim();
}

/** Parsing 3 baris berlabel (Petunjuk/Diteruskan/Catatan) dari teks balasan bebas. */
function parseDisposisiText(text) {
  const getLine = (label) => {
    const re = new RegExp(label + '\\s*[:\\-]\\s*(.+)', 'i');
    const m = text.match(re);
    return m ? m[1].trim() : '';
  };
  return {
    petunjuk: getLine('Petunjuk\\s*Disposisi'),
    diteruskanKepada: getLine('Diteruskan\\s*Kepada'),
    catatan: getLine('Catatan\\s*Tambahan')
  };
}

function extractEmailAddress(fromHeader) {
  const m = fromHeader.match(/<(.+)>/);
  return m ? m[1] : fromHeader.trim();
}

function getProcessedMessageIds() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY_PROCESSED);
  const arr = raw ? JSON.parse(raw) : [];
  return new Set(arr.slice(-500));
}

function markProcessed(msgId) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_KEY_PROCESSED);
  const arr = raw ? JSON.parse(raw) : [];
  arr.push(msgId);
  props.setProperty(PROP_KEY_PROCESSED, JSON.stringify(arr.slice(-500)));
}

/* ============================================================
 *  3. PROSES DISPOSISI (dipakai parser email & fallback web)
 *     + DISPOSISI BERJENJANG: rantai tak terbatas sampai Selesai
 * ============================================================ */

/**
 * Pemisah untuk daftar yang disimpan sistem (Pemegang Saat Ini, dst): HANYA titik koma / baris baru.
 * Koma tidak dipakai karena ada nama jabatan yang memuat koma
 * (mis. "Kasubbag Perencanaan, Pelaporan & TI").
 */
function pisahkanTarget_(teks) {
  if (!teks) return [];
  return teks.toString().split(/[;\n]/).map(s => s.trim()).filter(Boolean);
}

/** Sama seperti pisahkanTarget_, tetapi elemen kosong DIPERTAHANKAN agar indeks nama & email tetap sejajar. */
function pisahkanSejajar_(teks) {
  if (!teks) return [];
  return teks.toString().split(';').map(s => s.trim());
}

function pisahkanEmail_(teks) {
  if (!teks) return [];
  return teks.toString().split(/[;,\s]+/).map(s => s.trim()).filter(Boolean);
}

/**
 * Memecah teks tujuan yang DIKETIK BEBAS pimpinan (balasan email / formulir web).
 * Pemisah utama ';' atau baris baru. Koma hanya dianggap pemisah bila setiap potongannya
 * dikenali sebagai jabatan di Data Master — sehingga "Panitera, Sekretaris" terpecah
 * dengan benar, sedangkan "Kasubbag Perencanaan, Pelaporan & TI" tetap utuh.
 */
function pisahkanTujuanBebas_(teks) {
  if (!teks) return [];
  const hasil = [];
  teks.toString().split(/[;\n]/).map(s => s.trim()).filter(Boolean).forEach(bagian => {
    if (bagian.indexOf(',') === -1 || findPejabatByJabatanText(bagian, true)) { hasil.push(bagian); return; }
    const potongan = bagian.split(',').map(s => s.trim()).filter(Boolean);
    if (potongan.length > 1 && potongan.every(p => findPejabatByJabatanText(p))) {
      potongan.forEach(p => hasil.push(p));
    } else {
      hasil.push(bagian);
    }
  });
  return hasil;
}

/** Setelah pemegang bertindak, baris "Menunggu Disposisi" miliknya ditandai sudah diproses. */
function tandaiMenungguDiproses_(kodeTracking, dariNama) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DISPOSISI);
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getDataRange().getValues();
  const h = values[0];
  const cKode = h.indexOf('Kode Tracking'), cKepada = h.indexOf('Diteruskan Kepada'), cStatus = h.indexOf('Status Tindak Lanjut');
  for (let i = 1; i < values.length; i++) {
    if (values[i][cKode] === kodeTracking && values[i][cStatus] === STATUS_MENUNGGU &&
        String(values[i][cKepada]).trim() === String(dariNama).trim()) {
      sheet.getRange(i + 1, cStatus + 1).setValue('Sudah Didisposisikan');
    }
  }
}

/**
 * Mencatat satu langkah (hop) dalam rantai disposisi — BISA ke beberapa
 * tujuan sekaligus (fan-out) dalam satu aksi. Setiap tujuan yang valid
 * langsung dikirimi email lanjutan secara paralel (tidak perlu satu-satu),
 * dan masing-masing tetap bisa membalas independen untuk meneruskan lagi.
 * Dipanggil oleh: balasan email (parser), formulir web bypass, dan
 * tombol "Lanjutkan Disposisi" manual oleh admin di dashboard.
 *
 * suratRow : hasil findSuratByKode() (data surat SEBELUM hop ini dicatat)
 * dariNama : nama pemegang surat SAAT INI (yang sedang mendisposisikan)
 * disposisi: { petunjuk, diteruskanKepada (boleh multi, pisah ; atau ,), catatan }
 * metode   : 'Inbound Email Parsing' | 'Web Bypass (Fallback Link)' | 'Dicatat Manual oleh Admin'
 */
function catatDisposisiHop(suratRow, dariNama, disposisi, metode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMasuk = ss.getSheetByName(SHEET_NAMES.SURAT_MASUK);
  const headers = suratRow.headers;
  const now = new Date();

  const setCell = (colName, value) => {
    const col = headers.indexOf(colName) + 1;
    if (col > 0) sheetMasuk.getRange(suratRow.rowIndex, col).setValue(value);
  };

  let targetList = pisahkanTujuanBebas_(disposisi.diteruskanKepada);
  // "-" / kosong = pihak terakhir, tidak meneruskan ke siapa pun.
  const tidakMeneruskan = targetList.length === 0 || targetList.every(t => /^[-\u2013\u2014\s]*$/.test(t));
  if (tidakMeneruskan) targetList = [];

  setCell('Status', 'Didisposisikan');
  setCell('Petunjuk Disposisi', disposisi.petunjuk);
  setCell('Disposisi Kepada', targetList.length ? targetList.join('; ') : '-');
  setCell('Catatan Disposisi', disposisi.catatan);
  setCell('Tanggal Disposisi', now);
  setCell('Metode Disposisi', metode);

  // --- Daftar "pemegang saat ini" bisa berisi >1 orang sekaligus. Yang baru
  //     saja bertindak (dariNama) dikeluarkan; semua tujuan baru ditambahkan,
  //     sehingga tiap orang bisa membalas independen di waktunya sendiri. ---
  const pemegangLama = pisahkanTarget_(suratRow.data['Pemegang Saat Ini']);
  const emailPemegangLama = pisahkanSejajar_(suratRow.data['Email Pemegang Saat Ini']);
  const emailMap = {};
  pemegangLama.forEach((nm, i) => { emailMap[nm] = emailPemegangLama[i] || ''; });

  let pemegangBaru = tidakMeneruskan ? pemegangLama.slice() : pemegangLama.filter(nm => nm !== dariNama);

  const sheetDisposisi = ss.getSheetByName(SHEET_NAMES.DISPOSISI);
  const ringkasanStatus = [];

  targetList.forEach(target => {
    const tujuanPejabat = findPejabatByJabatanText(target);
    let statusLanjut, emailDiteruskan = '';

    const punyaEmail = tujuanPejabat && tujuanPejabat['Email'];
    const punyaWhatsApp = tujuanPejabat && tujuanPejabat['No. WhatsApp'];

    if (punyaEmail || punyaWhatsApp) {
      kirimDisposisiLanjutan(suratRow.data, dariNama, disposisi, tujuanPejabat, targetList);
      statusLanjut = 'Diteruskan ke ' + tujuanPejabat['Jabatan'];
      emailDiteruskan = tujuanPejabat['Email'] || '';
      if (pemegangBaru.indexOf(tujuanPejabat['Jabatan']) === -1) pemegangBaru.push(tujuanPejabat['Jabatan']);
      emailMap[tujuanPejabat['Jabatan']] = tujuanPejabat['Email'] || '';
    } else if (tujuanPejabat) {
      statusLanjut = 'Perlu Diteruskan Manual (email/WhatsApp ' + tujuanPejabat['Jabatan'] + ' belum diisi)';
      if (pemegangBaru.indexOf(tujuanPejabat['Jabatan']) === -1) pemegangBaru.push(tujuanPejabat['Jabatan']);
    } else {
      statusLanjut = 'Perlu Diteruskan Manual (jabatan "' + target + '" tidak dikenali)';
      if (pemegangBaru.indexOf(target) === -1) pemegangBaru.push(target);
    }

    ringkasanStatus.push(target + ': ' + statusLanjut);

    sheetDisposisi.appendRow([
      suratRow.data['Kode Tracking'], suratRow.data['No. Agenda'], dariNama,
      disposisi.petunjuk, target, disposisi.catatan, now, statusLanjut, '', emailDiteruskan
    ]);
  });

  if (tidakMeneruskan) {
    const st = 'Tidak diteruskan (pihak terakhir) — menunggu admin menutup surat';
    ringkasanStatus.push(st);
    sheetDisposisi.appendRow([
      suratRow.data['Kode Tracking'], suratRow.data['No. Agenda'], dariNama,
      disposisi.petunjuk, '-', disposisi.catatan, now, st, '', ''
    ]);
  }
  tandaiMenungguDiproses_(suratRow.data['Kode Tracking'], dariNama);

  const emailPemegangBaru = pemegangBaru.map(nm => emailMap[nm] || '');
  setCell('Pemegang Saat Ini', pemegangBaru.join('; '));
  setCell('Email Pemegang Saat Ini', emailPemegangBaru.join('; '));
  setCell('Status Diteruskan', ringkasanStatus.join(' | '));

  return { statusLanjut: ringkasanStatus.join(' | '), targetList };
}

/** Dipanggil dari parser email & formulir web bypass — dariNama diambil dari pemegang surat saat ini. */
function prosesDisposisi(suratRow, disposisi, dariNamaOpsional) {
  const dariNama = dariNamaOpsional || suratRow.data['Pemegang Saat Ini'] || suratRow.data['Pimpinan Tujuan'];
  const hasil = catatDisposisiHop(suratRow, dariNama, disposisi, disposisi.metode);
  kirimKonfirmasiDisposisi(suratRow.data, disposisi, dariNama);
  return hasil;
}

/** Dipanggil dari tombol "Lanjutkan Disposisi" manual di menu Disposisi (dashboard). */
function lanjutkanDisposisiManual(kodeTracking, disposisiInput) {
  const suratRow = findSuratByKode(kodeTracking);
  if (!suratRow) return { success: false, message: 'Surat dengan kode tersebut tidak ditemukan.' };
  if (suratRow.data['Status'] === 'Selesai') {
    return { success: false, message: 'Surat ini sudah berstatus Selesai — tidak bisa diteruskan lagi.' };
  }
  if (!disposisiInput.petunjuk) {
    return { success: false, message: 'Petunjuk disposisi wajib diisi.' };
  }

  const dariNama = suratRow.data['Pemegang Saat Ini'] || suratRow.data['Pimpinan Tujuan'];
  const disposisi = {
    petunjuk: disposisiInput.petunjuk,
    diteruskanKepada: disposisiInput.diteruskanKepada || '-',
    catatan: disposisiInput.catatan || ''
  };
  const hasil = catatDisposisiHop(suratRow, dariNama, disposisi, 'Dicatat Manual oleh Admin');

  return { success: true, statusLanjut: hasil.statusLanjut };
}

/** Mencari data pejabat di MasterPejabat berdasarkan teks nama jabatan (exact match diutamakan). */
function findPejabatByJabatanText(text, hanyaExact) {
  if (!text) return null;
  const clean = text.toString().trim().toLowerCase();
  if (!clean || /^[-\u2013\u2014]+$/.test(clean)) return null;

  const list = readSheetAsObjectsCached(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.MASTER_PEJABAT, 30);
  let found = list.find(p => (p['Jabatan'] || '').toString().trim().toLowerCase() === clean);
  if (found || hanyaExact || clean.length < 4) return found || null;

  found = list.find(p => {
    const jabatan = (p['Jabatan'] || '').toString().trim().toLowerCase();
    return jabatan && (clean.includes(jabatan) || jabatan.includes(clean));
  });
  return found || null;
}

/**
 * Kirim email disposisi lanjutan ke pejabat/staf berikutnya dalam rantai.
 * Email ini REPLYABLE — penerima bisa membalasnya untuk meneruskan lagi
 * ke pejabat/staf berikutnya, sehingga rantai bisa berlanjut tanpa batas
 * sampai seseorang menandainya cukup / admin menutup surat.
 */
function kirimDisposisiLanjutan(suratData, dariNama, disposisi, tujuanPejabat, semuaTarget) {
  const subject = `[${suratData['Kode Tracking']}] Disposisi Diteruskan: ${suratData['Hal']}`;
  const bypassUrl = `${ScriptApp.getService().getUrl()}?page=disposisi-web&kode=${encodeURIComponent(suratData['Kode Tracking'])}`;

  const pejabatList = readSheetAsObjects(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.MASTER_PEJABAT);
  const daftarTujuan = pejabatList.map(p => p['Jabatan']).filter(Boolean).join(', ');

  const lainnya = (semuaTarget || []).filter(t => t !== tujuanPejabat['Jabatan']);
  const infoBersama = lainnya.length
    ? `<p style="margin:8px 0 0;color:#6b7c93;font-size:12px;">Disposisi ini juga diteruskan bersamaan kepada: <b>${lainnya.join(', ')}</b>.</p>`
    : '';

  const htmlBody = `
    <div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:auto;border:1px solid #e6ebf2;border-radius:10px;overflow:hidden;">
      <div style="background:#0d9488;color:#fff;padding:16px 22px;">
        <div style="font-size:15px;font-weight:700;">📋 Disposisi Diteruskan Kepada Anda</div>
        <div style="font-size:12px;opacity:0.85;margin-top:2px;">Kode Tracking: ${suratData['Kode Tracking']}</div>
      </div>
      <div style="padding:20px 22px;color:#1a2b3c;font-size:13.5px;line-height:1.6;">
        <table style="width:100%;border-collapse:collapse;">
          <tr><td style="color:#6b7c93;width:150px;">Dari</td><td>${suratData['Dari'] || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Hal</td><td><b>${suratData['Hal'] || '-'}</b></td></tr>
          <tr><td style="color:#6b7c93;">Diteruskan oleh</td><td>${dariNama || '-'}</td></tr>
          <tr><td style="color:#6b7c93;">Petunjuk Disposisi</td><td><b>${disposisi.petunjuk}</b></td></tr>
          <tr><td style="color:#6b7c93;">Catatan</td><td>${disposisi.catatan || '-'}</td></tr>
          ${suratData['Lampiran (URL)'] ? `<tr><td style="color:#6b7c93;">Lampiran</td><td><a href="${suratData['Lampiran (URL)']}" target="_blank">Lihat File PDF</a></td></tr>` : ''}
        </table>
        <p style="margin:18px 0 0;">Surat ini didisposisikan kepada <b>${tujuanPejabat['Jabatan']}</b> oleh ${dariNama}.</p>
        ${infoBersama}

        <p style="margin:16px 0 8px;">Untuk meneruskan lagi ke pihak lain, <b>balas (Reply) email ini</b> — isi 3 baris berikut <u>di atas kutipan email ini</u>:</p>
        <div style="background:#f4f7fb;border:1px dashed #c7d3e0;border-radius:8px;padding:14px 16px;font-family:Consolas,monospace;font-size:12.5px;white-space:pre-wrap;">Petunjuk Disposisi : (Setuju/Tolak/Selesaikan/dll)
Diteruskan Kepada  : (satu atau lebih jabatan, pisahkan dengan titik koma ; — atau "-" jika Anda pihak terakhir)
Catatan Tambahan   : (opsional)</div>
        <p style="margin:14px 0 6px;color:#6b7c93;font-size:12px;">Pilihan tujuan disposisi (boleh pilih lebih dari satu sekaligus):</p>
        <div style="background:#eef2f7;border-radius:6px;padding:10px 14px;font-size:12px;color:#33475b;">${daftarTujuan || '-'}</div>

        <p style="margin:18px 0 6px;color:#6b7c93;font-size:12px;">Atau gunakan formulir web (tersedia kotak centang untuk pilih beberapa tujuan sekaligus):</p>
        <a href="${bypassUrl}" style="display:inline-block;background:#0f3d5c;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-size:13px;font-weight:600;">Lihat Detail &amp; Teruskan Disposisi →</a>
      </div>
      <div style="background:#f4f7fb;padding:12px 22px;font-size:11px;color:#8fa4b8;">
        Email ini bagian dari alur disposisi berjenjang Sistem E-Office Persuratan W8-A10.
      </div>
    </div>`;

  const plainBody = `Surat "${suratData['Hal']}" telah didisposisikan kepada Anda oleh ${dariNama}.\n`
    + `Petunjuk: ${disposisi.petunjuk}\nCatatan: ${disposisi.catatan || '-'}\n`
    + (lainnya.length ? `Diteruskan bersamaan kepada: ${lainnya.join(', ')}\n` : '')
    + `\nUntuk meneruskan lagi, balas email ini dengan format:\nPetunjuk Disposisi : ...\nDiteruskan Kepada  : (satu/lebih, pisah ; — pilih dari: ${daftarTujuan})\nCatatan Tambahan   : ...\n\n`
    + `Atau gunakan link berikut: ${bypassUrl}`;

  if (tujuanPejabat['Email']) {
    try {
      GmailApp.sendEmail(tujuanPejabat['Email'], subject, plainBody, {
        htmlBody: htmlBody,
        name: 'E-Office Persuratan W8-A10'
      });
    } catch (e) {
      Logger.log('Gagal mengirim disposisi lanjutan: ' + e);
    }
  }

  // Kirim WhatsApp dengan isi yang sama dengan email (versi teks polos),
  // ke "No. WhatsApp" pejabat tujuan jika sudah diisi di Data Master.
  if (tujuanPejabat['No. WhatsApp']) {
    try {
      const hasilWa = kirimWhatsApp_(tujuanPejabat['No. WhatsApp'], plainBody);
      if (!hasilWa.success) Logger.log('Gagal kirim WhatsApp disposisi lanjutan ke ' + tujuanPejabat['Jabatan'] + ': ' + hasilWa.message);
    } catch (e) {
      Logger.log('Gagal kirim WhatsApp disposisi lanjutan: ' + e);
    }
  }
}

function kirimKonfirmasiDisposisi(suratData, disposisi, dariNama) {
  const subject = `[${suratData['Kode Tracking']}] Disposisi Diterima — ${suratData['Hal']}`;
  const body = `Disposisi Anda untuk surat "${suratData['Hal']}" telah berhasil dicatat sistem.\n\n`
    + `Petunjuk   : ${disposisi.petunjuk}\n`
    + `Diteruskan : ${disposisi.diteruskanKepada}\n`
    + `Catatan    : ${disposisi.catatan}\n`
    + `Metode     : ${disposisi.metode}\n\nTerima kasih.`;

  // Cari email milik "dariNama" secara spesifik dari daftar pemegang (bisa berisi >1 nama).
  const namaList = pisahkanTarget_(suratData['Pemegang Saat Ini']);
  const emailList = pisahkanSejajar_(suratData['Email Pemegang Saat Ini']);
  const idx = namaList.indexOf(dariNama);
  const emailTujuan = (idx > -1 ? emailList[idx] : '') || suratData['Email Pimpinan'];
  if (!emailTujuan) return;

  try {
    GmailApp.sendEmail(emailTujuan, subject, body, { name: 'E-Office Persuratan W8-A10' });
  } catch (e) {
    Logger.log('Gagal kirim konfirmasi: ' + e);
  }
}

/* ============================================================
 *  4. FALLBACK BYPASS WEB (satu klik, tanpa login)
 * ============================================================ */
function renderDisposisiWebForm(kode) {
  const suratRow = kode ? findSuratByKode(kode) : null;
  const template = HtmlService.createTemplateFromFile('DisposisiForm');
  template.kode = kode;
  template.surat = suratRow ? suratRow.data : null;
  template.pejabatList = readSheetAsObjects(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAMES.MASTER_PEJABAT);
  return template.evaluate()
    .setTitle('Formulir Disposisi' + (kode ? ' — ' + kode : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Dipanggil dari DisposisiForm.html saat pimpinan submit lewat web (fallback). */
function submitDisposisiWeb(kode, petunjuk, diteruskanKepada, catatan) {
  const suratRow = findSuratByKode(kode);
  if (!suratRow) return { success: false, message: 'Kode tracking tidak ditemukan.' };
  if (suratRow.data['Status'] === 'Selesai') {
    return { success: false, message: 'Surat ini sudah ditutup (Selesai) oleh admin — tidak bisa diteruskan lagi.' };
  }
  if (!petunjuk) {
    return { success: false, message: 'Petunjuk disposisi wajib diisi.' };
  }

  prosesDisposisi(suratRow, {
    petunjuk, diteruskanKepada: diteruskanKepada || '-', catatan: catatan || '',
    metode: 'Web Bypass (Fallback Link)'
  });

  return { success: true, message: 'Disposisi berhasil disimpan.' };
}

/* ============================================================
 *  5. FUNGSI UJI COBA CEPAT (jalankan manual dari editor)
 * ============================================================ */
function TES_kirimContohNotifikasi() {
  // Ganti email di bawah dengan email Anda sendiri untuk uji coba
  const hasil = tambahSuratMasuk({
    nomorNaskahDinas: '001/UM/IX/2026',
    tanggalNaskahDinas: '2026-09-20',
    dari: 'Pengadilan Tinggi',
    sifat: 'Penting',
    jenis: 'Undangan',
    indexBidang: 'Kesekretariatan',
    hal: 'Undangan Rapat Koordinasi',
    statusNaskah: 'Asli',
    kodeKlasifikasi: 'RT',
    pimpinanTujuan: 'Ketua Pengadilan',
    emailPimpinan: 'ganti-dengan-email-anda@gmail.com'
  });
  Logger.log(hasil);
}

/* ============================================================
 *  6. MODUL TRACKING / PELACAKAN DISPOSISI (PRD Bab 5)
 * ============================================================ */

/**
 * Menandai satu baris disposisi sebagai selesai ditindaklanjuti,
 * sekaligus menutup surat masuk terkait (Status -> "Selesai" / Closed).
 */
function tandaiDisposisiSelesai(kodeTracking) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDisposisi = ss.getSheetByName(SHEET_NAMES.DISPOSISI);
  const values = sheetDisposisi.getDataRange().getValues();
  const headers = values[0];
  const kodeCol = headers.indexOf('Kode Tracking');
  const statusCol = headers.indexOf('Status Tindak Lanjut');
  const selesaiCol = headers.indexOf('Tanggal Selesai');
  const now = new Date();

  let updated = false;
  for (let i = 1; i < values.length; i++) {
    if (values[i][kodeCol] === kodeTracking) {
      sheetDisposisi.getRange(i + 1, statusCol + 1).setValue('Selesai');
      sheetDisposisi.getRange(i + 1, selesaiCol + 1).setValue(now);
      updated = true;
    }
  }

  const suratRow = findSuratByKode(kodeTracking);
  if (!updated && !suratRow) {
    return { success: false, message: 'Data surat/disposisi dengan kode tersebut tidak ditemukan.' };
  }

  if (suratRow) {
    const sheetMasuk = ss.getSheetByName(SHEET_NAMES.SURAT_MASUK);
    const col = suratRow.headers.indexOf('Status') + 1;
    if (col > 0) sheetMasuk.getRange(suratRow.rowIndex, col).setValue('Selesai');
  }

  return { success: true, message: 'Disposisi berhasil ditandai selesai.' };
}