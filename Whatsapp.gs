/**
 * ============================================================
 *  MODUL INTEGRASI WHATSAPP (FONNTE)
 *  Digunakan sebagai pelengkap notifikasi email pada alur
 *  disposisi — isi pesan WhatsApp dibuat sama dengan isi email
 *  (versi teks polos), dikirim ke "No. WhatsApp" pejabat yang
 *  terdaftar di Data Master.
 * ============================================================
 */

// Token API Fonnte. Bisa dipindah ke Script Properties nanti kalau
// suatu saat perlu ganti tanpa harus edit kode (Project Settings >
// Script Properties > key: FONNTE_TOKEN).
const FONNTE_TOKEN = 'JT3LZWuvn6mf2aQj41EK';
const FONNTE_ENDPOINT = 'https://api.fonnte.com/send';

/**
 * Merapikan nomor HP Indonesia ke format yang diminta Fonnte (62xxxxxxxxxx).
 * Menerima input dengan format apa pun: 08123..., +62812..., 62812..., 812...
 */
function normalisasiNomorWA_(nomor) {
  if (!nomor) return null;
  let n = String(nomor).trim().replace(/[^0-9]/g, '');
  if (!n) return null;
  if (n.charAt(0) === '0') n = '62' + n.slice(1);
  else if (n.indexOf('62') !== 0) n = '62' + n;
  return n;
}

/**
 * Mengirim satu pesan WhatsApp lewat Fonnte. Selalu mengembalikan
 * { success, message } — tidak pernah melempar exception ke pemanggil,
 * supaya kegagalan kirim WA tidak sampai menggagalkan proses utama
 * (simpan surat / catat disposisi) yang memanggilnya.
 */
function kirimWhatsApp_(nomor, pesan) {
  const nomorBersih = normalisasiNomorWA_(nomor);
  if (!nomorBersih) {
    return { success: false, message: 'Nomor WhatsApp tidak valid atau kosong: "' + nomor + '"' };
  }
  if (!pesan) {
    return { success: false, message: 'Isi pesan kosong.' };
  }

  try {
    const response = UrlFetchApp.fetch(FONNTE_ENDPOINT, {
      method: 'post',
      headers: { Authorization: FONNTE_TOKEN },
      payload: { target: nomorBersih, message: pesan },
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    let result;
    try {
      result = JSON.parse(response.getContentText());
    } catch (parseErr) {
      return { success: false, message: 'Respons Fonnte tidak dikenali (HTTP ' + statusCode + '): ' + response.getContentText() };
    }

    if (statusCode !== 200 || result.status === false || result.status === 'false') {
      return { success: false, message: (result && (result.reason || result.detail)) || ('Fonnte menolak permintaan (HTTP ' + statusCode + ').') };
    }

    return { success: true, message: 'Pesan terkirim ke ' + nomorBersih + '.' };
  } catch (e) {
    return { success: false, message: 'Gagal menghubungi Fonnte: ' + e };
  }
}

/**
 * Dipanggil dari menu "Uji Coba WhatsApp" di Data Master.
 * Kalau pesan tidak diisi, dipakai teks uji coba bawaan.
 */
function ujiCobaKirimWhatsApp(nomor, pesan) {
  const isiPesan = pesan && pesan.trim()
    ? pesan
    : 'Uji coba pengiriman WhatsApp dari aplikasi ' + APP_TITLE + '.\nJika Anda menerima pesan ini, integrasi Fonnte sudah berjalan dengan baik.';
  return kirimWhatsApp_(nomor, isiPesan);
}