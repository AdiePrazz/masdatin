/**
 * Bot Telegram Berita Acara PDPB
 * KPU Kabupaten Malang
 *
 * Flow:
 *  0. Pilih Fungsi   → tombol: Infografis PDPB | BA Pleno PDPB
 *     - Infografis   → kirim link sipikat.streamlit.app
 *     - BA Pleno     → lanjut ke flow pembuatan BA
 *  1. Nomor BA
 *  2. Triwulan       → tombol inline 1-4
 *  3. Tahun
 *  4. Tanggal rapat
 *  5. Jam rapat
 *  6. Pilih instansi → tombol inline multi-pilih
 *  7. Isi masukan tiap instansi
 *  8. Upload Excel   → validasi sheet "REKAPITULASI PDPB", ambil data otomatis
 *  9. Konfirmasi     → generate & kirim .docx
 */

require('dotenv').config();
const TelegramBot  = require('node-telegram-bot-api');
const fs           = require('fs');
const path         = require('path');
const https        = require('https');
const http         = require('http');
const { getSesi, resetSesi, DAFTAR_INSTANSI } = require('./session');
const { bacaExcel }  = require('./excelParser');
const { generateBA } = require('./generateBA');

const TOKEN = process.env.TELEGRAM_TOKEN;
if (!TOKEN) { console.error('❌ TELEGRAM_TOKEN belum diisi di .env'); process.exit(1); }

const bot = new TelegramBot(TOKEN, { polling: true });
const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

// ─── Helper: validasi tanggal ─────────────────────────────────────────────────
function parseTanggal(str) {
  const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (!m) return null;
  const [, d, mo, y] = m.map(Number);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return `${y}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}

// ─── Helper: format angka ─────────────────────────────────────────────────────
const fmt = n => parseInt(n).toLocaleString('id-ID');

// ─── Helper: download file dari Telegram ─────────────────────────────────────
function downloadBuffer(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith('https') ? https : http;
    lib.get(url, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ─── Keyboard pemilihan fungsi utama ─────────────────────────────────────────
const kbFungsi = {
  inline_keyboard: [[
    { text: '📊 Infografis PDPB Malang', callback_data: 'FUNGSI:INFOGRAFIS' },
  ],[
    { text: '📄 BA Pleno PDPB',          callback_data: 'FUNGSI:BA_PLENO'   },
  ]],
};

// ─── Keyboard inline triwulan ──────────────────────────────────────────────────
const kbTriwulan = {
  inline_keyboard: [[
    { text: '1️⃣ Kesatu (Jan-Mar)',   callback_data: 'TW:1' },
    { text: '2️⃣ Kedua (Apr-Jun)',    callback_data: 'TW:2' },
  ],[
    { text: '3️⃣ Ketiga (Jul-Sep)',   callback_data: 'TW:3' },
    { text: '4️⃣ Keempat (Okt-Des)', callback_data: 'TW:4' },
  ]],
};

// ─── Keyboard inline instansi ─────────────────────────────────────────────────
function kbInstansi(dipilih = []) {
  const rows = DAFTAR_INSTANSI.map((nama, i) => {
    const no = i + 1;
    const checked = dipilih.includes(no) ? '✅ ' : '';
    return [{ text: `${checked}${no}. ${nama}`, callback_data: `INST:${no}` }];
  });
  rows.push([{ text: '✔️ Selesai Pilih Instansi', callback_data: 'INST:DONE' }]);
  return { inline_keyboard: rows };
}

// ─── Ringkasan data sebelum konfirmasi ────────────────────────────────────────
function ringkasan(data) {
  const tw = ['','Kesatu','Kedua','Ketiga','Keempat'];
  const total = parseInt(data.jumlahLakiLaki) + parseInt(data.jumlahPerempuan);
  let t =
    `📋 *Ringkasan Berita Acara*\n\n` +
    `📌 Nomor BA        : \`${data.nomorBA}\`\n` +
    `📅 Triwulan        : ${tw[data.triwulan]}\n` +
    `📅 Tahun           : ${data.tahun}\n` +
    `📅 Tanggal Rapat   : ${data.tanggalRapat.split('-').reverse().join('/')}\n` +
    `⏰ Jam Rapat       : ${data.jamRapat} WIB\n\n` +
    `📊 *Rekapitulasi Pemilih* _(dari Excel)_:\n` +
    `   Kecamatan  : ${fmt(data.jumlahKecamatan)}\n` +
    `   Desa/Kel   : ${fmt(data.jumlahDesaKel)}\n` +
    `   Laki-laki  : ${fmt(data.jumlahLakiLaki)}\n` +
    `   Perempuan  : ${fmt(data.jumlahPerempuan)}\n` +
    `   *Total     : ${fmt(total)}*\n\n` +
    `🏛️ *Masukan Instansi*:\n`;
  data.masukanInstansi.forEach((inst, i) => {
    t += `   ${i+1}. ${inst.nama}\n`;
    t += `      _"${inst.isi.substring(0,80)}${inst.isi.length>80?'...':''}"_\n`;
  });
  t += `\nApakah data sudah benar?`;
  return t;
}

const kbKonfirmasi = {
  inline_keyboard: [[
    { text: '✅ Ya, Buat Dokumen', callback_data: 'KONFIRMASI:YA'   },
    { text: '❌ Tidak, Ulangi',    callback_data: 'KONFIRMASI:TIDAK' },
  ]],
};

// ─── Kirim pesan helper ───────────────────────────────────────────────────────
async function kirim(chatId, text, opts={}) {
  try {
    return await bot.sendMessage(chatId, text, { parse_mode: 'Markdown', ...opts });
  } catch(e) {
    return await bot.sendMessage(chatId, text.replace(/[*_`]/g,''), opts);
  }
}

async function edit(chatId, msgId, text, opts={}) {
  try {
    return await bot.editMessageText(text, { chat_id: chatId, message_id: msgId, parse_mode: 'Markdown', ...opts });
  } catch(e) { /* pesan tidak berubah, abaikan */ }
}

// ═════════════════════════════════════════════════════════════════════════════
// HANDLER PESAN TEKS
// ═════════════════════════════════════════════════════════════════════════════
bot.on('message', async (msg) => {
  if (!msg.text && !msg.document) return;

  const chatId = msg.chat.id;
  const userId = String(chatId);
  const teks   = (msg.text || '').trim();
  const lower  = teks.toLowerCase();

  // ── Perintah global ──────────────────────────────────────────────────────
  if (['/start', '/mulai'].includes(lower) || lower === 'mulai') {
    resetSesi(userId);
    const sess = getSesi(userId);
    sess.step = 'PILIH_FUNGSI';
    await kirim(chatId,
      `🗳️ *Hello! What can i do for you?*\n\n` +
      `Silahkan pilih fungsi yang Anda inginkan:\n` +
      `_Ketik /batal kapan saja untuk membatalkan._`,
      { reply_markup: kbFungsi }
    );
    return;
  }
  //   await kirim(chatId,
  //     `✏️ *Langkah 1 dari 8*\n\nMasukkan *Nomor Berita Acara*:\n\nContoh: \`63/PK.01-BA/3507/2025\``
  //   );
  //   const sess = getSesi(userId);
  //   sess.step = 'NOMOR_BA';
  //   return;
  // }

  if (['/batal', 'batal'].includes(lower)) {
    resetSesi(userId);
    await kirim(chatId, '❌ Proses dibatalkan.\n\nKetik /mulai untuk memulai kembali.');
    return;
  }

  const sess = getSesi(userId);

  // ── Langkah 1: Nomor BA ───────────────────────────────────────────────────
  if (sess.step === 'NOMOR_BA') {
    if (!teks) { await kirim(chatId, '⚠️ Nomor BA tidak boleh kosong.'); return; }
    sess.data.nomorBA = teks;
    sess.step = 'TRIWULAN';
    await kirim(chatId,
      `✅ Nomor BA disimpan.\n\n✏️ *Langkah 2 dari 8*\n\nPilih *Triwulan*:`,
      { reply_markup: kbTriwulan }
    );
    return;
  }

  // ── Langkah 3: Tahun ──────────────────────────────────────────────────────
  if (sess.step === 'TAHUN') {
    const th = parseInt(teks);
    if (isNaN(th) || th < 2025 || th > 2029) {
      await kirim(chatId, '⚠️ Tahun tidak valid. Masukkan tahun antara 2025–2029.'); return;
    }
    sess.data.tahun = th;
    sess.step = 'TANGGAL';
    await kirim(chatId,
      `✅ Tahun: *${th}*\n\n✏️ *Langkah 4 dari 8*\n\nMasukkan *Tanggal Rapat Pleno*:\n\nFormat: \`DD/MM/YYYY\`\nContoh: \`08/12/2025\``
    );
    return;
  }

  // ── Langkah 4: Tanggal ───────────────────────────────────────────────────
  if (sess.step === 'TANGGAL') {
    const tgl = parseTanggal(teks);
    if (!tgl) { await kirim(chatId, '⚠️ Format salah. Gunakan DD/MM/YYYY\nContoh: 08/12/2025'); return; }
    sess.data.tanggalRapat = tgl;
    sess.step = 'JAM';
    await kirim(chatId,
      `✅ Tanggal: *${teks}*\n\n✏️ *Langkah 5 dari 8*\n\nMasukkan *Jam Rapat*:\n\nFormat: \`HH.MM\`\nContoh: \`10.00\``
    );
    return;
  }

  // ── Langkah 5: Jam ───────────────────────────────────────────────────────
  if (sess.step === 'JAM') {
    if (!teks.match(/^\d{1,2}[.:]\d{2}$/)) {
      await kirim(chatId, '⚠️ Format jam salah.\nContoh: 10.00 atau 09.30'); return;
    }
    sess.data.jamRapat = teks.replace(':','.');
    sess.step = 'PILIH_INSTANSI';
    sess.instansiDipilih = [];
    await kirim(chatId,
      `✅ Jam: *${teks} WIB*\n\n✏️ *Langkah 6 dari 8*\n\nPilih *instansi yang hadir* (bisa lebih dari satu).\nTekan nama instansi untuk memilih ✅, lalu tekan *Selesai*:`,
      { reply_markup: kbInstansi([]) }
    );
    return;
  }

  // ── Langkah 7: Isi masukan per instansi ──────────────────────────────────
  if (sess.step === 'ISI_MASUKAN') {
    if (!teks) { await kirim(chatId, '⚠️ Masukan tidak boleh kosong.'); return; }
    sess.data.masukanInstansi[sess.masukanIdx].isi = teks;
    sess.masukanIdx++;

    if (sess.masukanIdx < sess.instansiDipilih.length) {
      const noBerikut = sess.instansiDipilih[sess.masukanIdx];
      const namaBerikut = DAFTAR_INSTANSI[noBerikut - 1];
      await kirim(chatId,
        `✅ Masukan tersimpan.\n\n` +
        `✏️ Masukkan isi masukan dari:\n*${namaBerikut}*\n\n` +
        `_(${sess.masukanIdx+1}/${sess.instansiDipilih.length})_`
      );
    } else {
      // Semua masukan selesai → minta upload Excel
      sess.step = 'UPLOAD_EXCEL';
      await kirim(chatId,
        `✅ Semua masukan instansi tersimpan.\n\n` +
        `✏️ *Langkah 8 dari 8*\n\n` +
        `📊 Silakan *upload file Excel* rekap pemilih (Model-A Rekap):\n\n` +
        `• Wajib ada sheet bernama *"REKAPITULASI PDPB"*`
      );
    }
    return;
  }

  // Pesan teks lain saat step UPLOAD_EXCEL
  if (sess.step === 'UPLOAD_EXCEL' && !msg.document) {
    await kirim(chatId, '📎 Silakan *kirim file Excel* (.xlsx) di atas. Jangan ketik teks biasa.');
    return;
  }

  // Default
  if (sess.step === 'MULAI' || !sess.step) {
    await kirim(chatId, `Ketik /mulai untuk memulai pembuatan Berita Acara.`);
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// HANDLER UPLOAD DOKUMEN (Excel)
// ═════════════════════════════════════════════════════════════════════════════
bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const userId = String(chatId);
  const sess   = getSesi(userId);

  if (sess.step !== 'UPLOAD_EXCEL') {
    await kirim(chatId, '⚠️ Upload file tidak diperlukan sekarang.\nKetik /mulai untuk memulai.');
    return;
  }

  const doc  = msg.document;
  const name = (doc.file_name || '').toLowerCase();

  // Validasi ekstensi
  if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
    await kirim(chatId, '❌ File harus berformat *Excel (.xlsx)*.\nSilakan upload ulang file yang benar.');
    return;
  }

  await kirim(chatId, '⏳ Membaca file Excel...');

  try {
    // Download file
    const fileInfo  = await bot.getFile(doc.file_id);
    const fileUrl   = `https://api.telegram.org/file/bot${TOKEN}/${fileInfo.file_path}`;
    const buffer    = await downloadBuffer(fileUrl);

    // Parse Excel
    const hasil = bacaExcel(buffer);

    // Simpan ke session
    sess.data.jumlahKecamatan  = hasil.jumlahKecamatan;
    sess.data.jumlahDesaKel    = hasil.jumlahDesaKel;
    sess.data.jumlahLakiLaki   = hasil.jumlahLakiLaki;
    sess.data.jumlahPerempuan  = hasil.jumlahPerempuan;

    const total = hasil.jumlahLakiLaki + hasil.jumlahPerempuan;

    let infoExcel =
      `✅ *File Excel berhasil dibaca!*\n\n` +
      `📊 *Data yang diambil dari Excel:*\n` +
      `   Kecamatan : ${fmt(hasil.jumlahKecamatan)}\n` +
      `   Desa/Kel  : ${fmt(hasil.jumlahDesaKel)}\n` +
      `   Laki-laki : ${fmt(hasil.jumlahLakiLaki)}\n` +
      `   Perempuan : ${fmt(hasil.jumlahPerempuan)}\n` +
      `   *Total    : ${fmt(total)}*`;

    if (hasil.infoTriwulan) {
      infoExcel += `\n\n_Info dari Excel: ${hasil.infoTriwulan}_`;
    }

    await kirim(chatId, infoExcel);

    // Tampilkan konfirmasi
    sess.step = 'KONFIRMASI';
    await kirim(chatId, ringkasan(sess.data), { reply_markup: kbKonfirmasi });

  } catch(err) {
    await kirim(chatId,
      `❌ *Gagal membaca Excel:*\n\n${err.message}\n\nSilakan upload ulang file yang benar.`
    );
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// HANDLER CALLBACK QUERY (tombol inline)
// ═════════════════════════════════════════════════════════════════════════════
bot.on('callback_query', async (query) => {
  const chatId = query.message.chat.id;
  const msgId  = query.message.message_id;
  const userId = String(chatId);
  const data   = query.data;
  const sess   = getSesi(userId);

  await bot.answerCallbackQuery(query.id); // hapus loading

  // ── Pilih Fungsi Utama ────────────────────────────────────────────────────
  if (data.startsWith('FUNGSI:') && sess.step === 'PILIH_FUNGSI') {
    const pilihan = data.split(':')[1];

    if (pilihan === 'INFOGRAFIS') {
      await edit(chatId, msgId, `📊 *Infografis PDPB Malang*`);
      await kirim(chatId,
        `📊 *Infografis PDPB Malang*\n\n` +
        `Silakan akses dashboard infografis pemilih melalui tautan berikut:\n\n` +
        `🔗 https://sipikat.streamlit.app\n\n` +
        `_Ketik /mulai untuk kembali ke menu utama._`
      );
      resetSesi(userId);
      return;
    }

    if (pilihan === 'BA_PLENO') {
      await edit(chatId, msgId, `📄 *Pembuatan BA Pleno PDPB*`);
      sess.step = 'NOMOR_BA';
      await kirim(chatId,
        `✏️ *Langkah 1 dari 8*\n\nMasukkan *Nomor Berita Acara*:\n\nContoh: \`63/PK.01-BA/3507/2025\``
      );
      return;
    }
  }

  // ── Pilih Triwulan ────────────────────────────────────────────────────────
  if (data.startsWith('TW:') && sess.step === 'TRIWULAN') {
    const tw = parseInt(data.split(':')[1]);
    sess.data.triwulan = tw;
    sess.step = 'TAHUN';

    const label = ['','Kesatu','Kedua','Ketiga','Keempat'];
    await edit(chatId, msgId, `✅ Triwulan: *${label[tw]}*`);
    await kirim(chatId,
      `✏️ *Langkah 3 dari 8*\n\nMasukkan *Tahun*:\n\nContoh: \`2025\``
    );
    return;
  }

  // ── Pilih / toggle Instansi ───────────────────────────────────────────────
  if (data.startsWith('INST:') && sess.step === 'PILIH_INSTANSI') {
    const val = data.split(':')[1];

    if (val === 'DONE') {
      // Selesai pilih instansi
      if (sess.instansiDipilih.length === 0) {
        await bot.answerCallbackQuery(query.id, { text: '⚠️ Pilih minimal 1 instansi!', show_alert: true });
        return;
      }

      // Inisialisasi data masukan
      sess.data.masukanInstansi = sess.instansiDipilih.map(no => ({
        nama: DAFTAR_INSTANSI[no - 1],
        isi: '',
      }));
      sess.masukanIdx = 0;
      sess.step = 'ISI_MASUKAN';

      let listPilih = `✅ *Instansi yang dipilih:*\n`;
      sess.instansiDipilih.forEach((no, i) => {
        listPilih += `${i+1}. ${DAFTAR_INSTANSI[no-1]}\n`;
      });
      await edit(chatId, msgId, listPilih);

      await kirim(chatId,
        `✏️ *Langkah 7 dari 8 — Masukan Instansi*\n\n` +
        `Masukkan isi masukan/pernyataan dari:\n*${DAFTAR_INSTANSI[sess.instansiDipilih[0]-1]}*\n\n` +
        `_(1/${sess.instansiDipilih.length})_`
      );
      return;
    }

    // Toggle pilihan
    const no = parseInt(val);
    const idx = sess.instansiDipilih.indexOf(no);
    if (idx === -1) {
      sess.instansiDipilih.push(no);
    } else {
      sess.instansiDipilih.splice(idx, 1);
    }
    sess.instansiDipilih.sort((a,b) => a-b);

    // Update keyboard dengan tanda centang
    try {
      await bot.editMessageReplyMarkup(
        kbInstansi(sess.instansiDipilih),
        { chat_id: chatId, message_id: msgId }
      );
    } catch(e) { /* abaikan jika keyboard sama */ }
    return;
  }

  // ── Konfirmasi ─────────────────────────────────────────────────────────────
  if (data.startsWith('KONFIRMASI:') && sess.step === 'KONFIRMASI') {
    const jawab = data.split(':')[1];

    if (jawab === 'TIDAK') {
      resetSesi(userId);
      await edit(chatId, msgId, '🔄 Proses diulang dari awal.');
      await kirim(chatId, 'Ketik /mulai untuk memulai kembali.');
      return;
    }

    if (jawab === 'YA') {
      await edit(chatId, msgId, '⏳ Sedang membuat dokumen Word...');
      try {
        const fname   = `BA_PDPB_TW${sess.data.triwulan}_${sess.data.tahun}_${Date.now()}.docx`;
        const outPath = path.join(OUTPUT_DIR, fname);
        await generateBA(sess.data, outPath);

        await bot.sendDocument(chatId, outPath, {
          caption: `📄 *Berita Acara PDPB Triwulan ${sess.data.triwulan} Tahun ${sess.data.tahun}*\n\nDokumen berhasil dibuat ✅`,
          parse_mode: 'Markdown',
        });

        setTimeout(() => { try { fs.unlinkSync(outPath); } catch(e){} }, 120_000);
        resetSesi(userId);
        await kirim(chatId, '✅ Selesai! Ketik /mulai untuk membuat Berita Acara baru.');
      } catch(err) {
        await kirim(chatId, `❌ Gagal membuat dokumen:\n${err.message}\n\nKetik /mulai untuk mencoba lagi.`);
        console.error(err);
        resetSesi(userId);
      }
    }
    return;
  }
});

// ─── Error polling ────────────────────────────────────────────────────────────
bot.on('polling_error', err => console.error('[polling error]', err.message));

console.log('🤖 MAS DATIN BOT aktif...');
