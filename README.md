# Bot Telegram Berita Acara PDPB — KPU Kabupaten Malang

## Instalasi

```bash
# 1. Masuk ke folder
cd kpu-bot-v2

# 2. Install dependency
npm install

# 3. Buat file .env
copy .env.example .env        # Windows
cp .env.example .env          # Linux/Mac

# 4. Isi token Telegram di file .env
TELEGRAM_TOKEN=token_dari_botfather

# 5. Jalankan bot
node telegram-bot.js
```

## Cara Dapat Token Telegram
1. Buka Telegram → cari @BotFather
2. Ketik /newbot → ikuti instruksi
3. Salin token ke file .env

## Alur Bot
| Langkah | Input |
|---------|-------|
| 1 | Nomor BA (ketik manual) |
| 2 | Triwulan (tombol 1-4) |
| 3 | Tahun (ketik manual) |
| 4 | Tanggal rapat DD/MM/YYYY |
| 5 | Jam rapat HH.MM |
| 6 | Pilih instansi (tombol centang) |
| 7 | Isi masukan tiap instansi |
| 8 | Upload file Excel (sheet: REKAPITULASI PDPB) |
| → | Bot kirim file .docx otomatis |

## Agar Bot Berjalan Terus (VPS)
```bash
npm install -g pm2
pm2 start telegram-bot.js --name kpu-ba-bot
pm2 save && pm2 startup
```
