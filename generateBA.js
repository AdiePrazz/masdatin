/**
 * Generator Berita Acara PDPB v2
 * Font: Arial 12pt
 * Kop: Logo KPU di tengah + teks di bawahnya
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign,
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── Konstanta ────────────────────────────────────────────────────────────────
const FONT = 'Arial';
const SIZE = 24; // 12pt = 24 half-points

const TRIWULAN = {
  1: { romawi: 'KESATU',  teks: 'Kesatu'  },
  2: { romawi: 'KEDUA',   teks: 'Kedua'   },
  3: { romawi: 'KETIGA',  teks: 'Ketiga'  },
  4: { romawi: 'KEEMPAT', teks: 'Keempat' },
};

const NAMA_HARI  = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const NAMA_BULAN = ['Januari','Februari','Maret','April','Mei','Juni',
                    'Juli','Agustus','September','Oktober','November','Desember'];

// ─── Angka → Teks ─────────────────────────────────────────────────────────────
function angkaKeTeks(n) {
  n = parseInt(n);
  if (n === 0) return 'Nol';
  const s = ['','Satu','Dua','Tiga','Empat','Lima','Enam','Tujuh','Delapan','Sembilan','Sepuluh',
             'Sebelas','Dua Belas','Tiga Belas','Empat Belas','Lima Belas','Enam Belas',
             'Tujuh Belas','Delapan Belas','Sembilan Belas'];
  const p = ['','','Dua Puluh','Tiga Puluh','Empat Puluh','Lima Puluh',
             'Enam Puluh','Tujuh Puluh','Delapan Puluh','Sembilan Puluh'];
  if (n < 20)    return s[n];
  if (n < 100)   return p[Math.floor(n/10)] + (n%10 ? ' '+s[n%10] : '');
  if (n < 200)   return 'Seratus'   + (n%100 ? ' '+angkaKeTeks(n%100) : '');
  if (n < 1000)  return s[Math.floor(n/100)]+' Ratus' + (n%100 ? ' '+angkaKeTeks(n%100) : '');
  if (n < 2000)  return 'Seribu'    + (n%1000  ? ' '+angkaKeTeks(n%1000) : '');
  if (n < 1e6)   return angkaKeTeks(Math.floor(n/1000))+' Ribu' + (n%1000 ? ' '+angkaKeTeks(n%1000) : '');
  return n.toString();
}

function fmt(n) { return parseInt(n).toLocaleString('id-ID'); }

// ─── Border helpers ───────────────────────────────────────────────────────────
const bdr  = (c='000000') => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const allB = ()  => ({ top: bdr(), bottom: bdr(), left: bdr(), right: bdr() });
const noB  = ()  => ({
  top:    { style: BorderStyle.NONE },
  bottom: { style: BorderStyle.NONE },
  left:   { style: BorderStyle.NONE },
  right:  { style: BorderStyle.NONE },
});

// ─── Run / Paragraph helpers ──────────────────────────────────────────────────
const run = (text, opts={}) =>
  new TextRun({ text, font: FONT, size: SIZE, ...opts });

const para = (children, opts={}) =>
  new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });

const centerPara = (children, opts={}) =>
  para(children, { alignment: AlignmentType.CENTER, ...opts });

const justPara = (text, opts={}) =>
  para([run(text)], { alignment: AlignmentType.JUSTIFIED, ...opts });

// ─────────────────────────────────────────────────────────────────────────────
// FUNGSI UTAMA
// ─────────────────────────────────────────────────────────────────────────────
async function generateBA(data, outputPath) {
  const {
    nomorBA, triwulan, tahun, tanggalRapat, jamRapat,
    jumlahKecamatan, jumlahDesaKel, jumlahLakiLaki, jumlahPerempuan,
    masukanInstansi,
  } = data;

  const tw      = TRIWULAN[triwulan];
  const total   = parseInt(jumlahLakiLaki) + parseInt(jumlahPerempuan);
  const [tY,tM,tD] = tanggalRapat.split('-').map(Number);
  const tglTeks = `${angkaKeTeks(tD)} bulan ${NAMA_BULAN[tM-1]} tahun ${angkaKeTeks(parseInt(tahun))}`;
  const namaHari = NAMA_HARI[new Date(tY, tM-1, tD).getDay()];
  const tahunTeks = angkaKeTeks(parseInt(tahun));

  // Logo
  const logoPath = path.join(__dirname, 'kpu_logo.png');
  const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : null;

  // ══════════════════════════════════════════════════
  // KOP SURAT — Logo di tengah, teks di bawah logo
  // ══════════════════════════════════════════════════
  const kopChildren = [];

  if (logoBuffer) {
    kopChildren.push(
      centerPara([
        new ImageRun({
          data: logoBuffer,
          transformation: { width: 145, height: 113 },
          type: 'png',
        }),
      ], { spacing: { after: 60 } })
    );
  }

  kopChildren.push(
    centerPara([run('KOMISI PEMILIHAN UMUM', { bold: true, size: 26 })],
      { spacing: { before: 0, after: 0 } })
  );
  kopChildren.push(
    centerPara([run('KABUPATEN MALANG', { bold: true, size: 26 })],
      { spacing: { before: 0, after: 260 } })
  );

  // Garis bawah kop
  const garisBawahKop = new Paragraph({
    border: { bottom: { style: BorderStyle.NONE, size: 18, color: '000000', space: 1 } },
    children: [run('')],
    spacing: { after: 200 },
  });

  // ══════════════════════════════════════════════════
  // JUDUL
  // ══════════════════════════════════════════════════
  const spasing_isi = { line: 360, lineRule: 'auto' };

  const judulSection = [
    centerPara([run('BERITA ACARA', { underline: {} })],
      { spacing: { before: 160, after: 0 } }),
    centerPara([run(`Nomor : ${nomorBA}`)],
      { spacing: { before: 0, after: 120 } }),
    centerPara([run('TENTANG')],
      { spacing: { before: 120, after: 0 } }),
    centerPara([run('REKAPITULASI DAFTAR PEMILIH BERKELANJUTAN')],
      { spacing: { before: 0, after: 0 } }),
    centerPara([run(`TRIWULAN ${tw.romawi} TAHUN ${tahun}`)],
      { spacing: { before: 0, after: 240 } }),
  ];

  // ══════════════════════════════════════════════════
  // ISI PARAGRAF
  // ══════════════════════════════════════════════════
  const isi1 = justPara(
    `Pada hari ini ${namaHari}, tanggal ${tglTeks}, bertempat di Aula Tumapel Komisi Pemilihan Umum (KPU) Kabupaten Malang, pukul ${jamRapat} WIB, KPU Kabupaten Malang telah melaksanakan Rapat Pleno Terbuka Rekapitulasi Daftar Pemilih Berkelanjutan Triwulan ${tw.teks} Tahun ${tahunTeks} Tingkat Kabupaten Malang.`,
    { indent: { firstLine: 720 }, spacing: { after: 160, ...spasing_isi } }
  );

  const isi2 = justPara(
    'Dalam Rapat tersebut, KPU Kabupaten Malang menetapkan Rekapitulasi Pemilih Berkelanjutan Kabupaten Malang dengan rincian sebagai berikut:',
    { indent: { firstLine: 720 }, spacing: { after: 160, ...spasing_isi } }
  );

  // ══════════════════════════════════════════════════
  // POIN 1 — Tabel Rekapitulasi
  // ══════════════════════════════════════════════════
  const poin1 = para([run('1.\tRekapitulasi Daftar Pemilih Berkelanjutan')],
    { spacing: { after: 80 } });

  const mkCell = (text, bold=false, bg='FFFFFF') =>
    new TableCell({
      borders: allB(),
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [centerPara([run(text, { bold, size: 20 })])],
    });

  const tabelRekap = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [1865, 1865, 1866, 1866, 1864],
    rows: [
      // Header 1 - judul tabel
      new TableRow({ children: [
        new TableCell({
          borders: allB(), columnSpan: 5,
          shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [centerPara([run('REKAPITULASI DAFTAR PEMILIH BERKELANJUTAN', { bold: true, size: 20 })])],
        }),
      ]}),
      // Header 2 - nama kolom
      new TableRow({ children: [
        mkCell('JUMLAH KECAMATAN', true, 'D9D9D9'),
        mkCell('JUMLAH DESA/\nKELURAHAN', true, 'D9D9D9'),
        mkCell('JUMLAH LAKI-LAKI', true, 'D9D9D9'),
        mkCell('JUMLAH PEREMPUAN', true, 'D9D9D9'),
        mkCell('TOTAL', true, 'D9D9D9'),
      ]}),
      // Data
      new TableRow({ children: [
        mkCell(fmt(jumlahKecamatan)),
        mkCell(fmt(jumlahDesaKel)),
        mkCell(fmt(jumlahLakiLaki)),
        mkCell(fmt(jumlahPerempuan)),
        mkCell(fmt(total)),
      ]}),
    ],
  });

  // ══════════════════════════════════════════════════
  // POIN 2 — Masukan Instansi
  // ══════════════════════════════════════════════════
  const poin2Header = para([run('2.\tMenerima masukan dari:')],
    { spacing: { before: 160, after: 80 } });

  const masukanParas = [];
  masukanInstansi.forEach((inst, idx) => {
    masukanParas.push(
      para([run(`2.${idx+1}.\t${inst.nama}`, { bold: true })],
        { indent: { left: 720 }, spacing: { before: 120, after: 60, ...spasing_isi } })
    );
    inst.isi.split('\n').filter(b => b.trim()).forEach(brs => {
      masukanParas.push(
        para([run(`-\t${brs.trim()}`)],
          { indent: { left: 1080 }, spacing: { after: 60, ...spasing_isi } })
      );
    });
  });

  // ══════════════════════════════════════════════════
  // PENUTUP
  // ══════════════════════════════════════════════════
  const penutup = justPara(
    'Daftar Pemilih Berkelanjutan tersebut selanjutnya ditetapkan secara lebih rinci dalam dokumen Rekapitulasi Tingkat kabupaten sebagaimana terlampir yang merupakan bagian tidak terpisahkan dari Berita Acara ini.',
    { indent: { firstLine: 720 }, spacing: { before: 200, after: 200, ...spasing_isi } }
  );

  // ══════════════════════════════════════════════════
  // TABEL TEMPAT & TANGGAL
  // ══════════════════════════════════════════════════
  const mkInfoCell = (text, w) =>
    new TableCell({
      borders: noB(), width: { size: w, type: WidthType.DXA },
      margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: [para([run(text)])],
    });

  const tglTable = new Table({
    width: { size: 5000, type: WidthType.DXA },
    indent: { size: 4826, type: WidthType.DXA },
    borders: noB(),
    columnWidths: [1800, 200, 2500],
    rows: [
      new TableRow({ 
        children: [mkInfoCell('Dibuat di', 2000), mkInfoCell(':', 200), mkInfoCell('Kepanjen', 2800)],
        borders: noB()
      }),
      new TableRow({ 
        children: [mkInfoCell('Pada Tanggal', 2000), mkInfoCell(':', 200), mkInfoCell(`${String(tD).padStart(2,'0')} ${NAMA_BULAN[tM-1]} ${tahun}`, 2800)],
        borders: noB()
      }),
    ],
  });

  const demikian = justPara(
    'Demikian Berita Acara ini dibuat untuk dipergunakan sebagaimana mestinya.',
    { indent: { firstLine: 720 }, spacing: { before: 160, after: 300 } }
  );

  // ══════════════════════════════════════════════════
  // TANDA TANGAN
  // ══════════════════════════════════════════════════
  const ttdData = [
    { no:'1.', nama:'ABDUL FATAH',                   jbt:'KETUA',   sisi:'kiri'  },
    { no:'2.', nama:'NURHASIN',                       jbt:'ANGGOTA', sisi:'kanan' },
    { no:'3.', nama:'MARHAENDRA PRAMUDYA MAHARDIKA',  jbt:'ANGGOTA', sisi:'kiri'  },
    { no:'4.', nama:'ASKARI',                         jbt:'ANGGOTA', sisi:'kanan' },
    { no:'5.', nama:'BANGKIT MARHAENDRA',             jbt:'ANGGOTA', sisi:'kiri'  },
  ];

  const mkTtdCell = (text, w) =>
    new TableCell({
      borders: noB(), width: { size: w, type: WidthType.DXA },
      margins: { top: 100, bottom: 100, left: 60, right: 60 },
      children: [para([run(text)])],
    });

  // Baris kosong untuk memperlebar sel TTD (ruang tanda tangan)
  const emptyLine = () => new Paragraph({ children: [new TextRun({ text: '', size: SIZE })] });

  // Sel kotak TTD — border hitam tipis, tinggi diperlebar via paragraf kosong
  const mkTtdBox = (text, w) =>
    new TableCell({
      borders: {
        top:    { style: BorderStyle.NONE, size: 6, color: '000000' },
        bottom: { style: BorderStyle.NONE, size: 6, color: '000000' },
        left:   { style: BorderStyle.NONE, size: 6, color: '000000' },
        right:  { style: BorderStyle.NONE, size: 6, color: '000000' },
      },
      width: { size: w, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 80, right: 80 },
      verticalAlign: VerticalAlign.TOP,
      children: [
        // Nomor di atas
        para([run(text, { size: 20 })]),
        // 4 baris kosong = ~ruang tanda tangan ±3cm
        emptyLine(), emptyLine(), emptyLine(), emptyLine(),
      ],
    });

  const ttdTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [400, 3200, 1400, 2163, 2163],
    rows: ttdData.map(t =>
      new TableRow({ children: [
        mkTtdCell(t.no, 400),
        mkTtdCell(t.nama, 3200),
        mkTtdCell(t.jbt, 1400),
        mkTtdCell(t.sisi==='kiri'  ? `${t.no.replace('.','')}.…………………` : '', 2163),
        mkTtdCell(t.sisi==='kanan' ? `${t.no.replace('.','')}.…………………` : '', 2163),
      ]})
    ),
  });

  // ══════════════════════════════════════════════════
  // RAKIT DOKUMEN
  // ══════════════════════════════════════════════════
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: SIZE } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 },
        },
      },
      children: [
        ...kopChildren,
        garisBawahKop,
        ...judulSection,
        isi1,
        isi2,
        poin1,
        tabelRekap,
        poin2Header,
        ...masukanParas,
        penutup,
        tglTable,
        demikian,
        centerPara([run('KOMISI PEMILIHAN UMUM')],
          { spacing: { before: 160, after: 0 } }),
        centerPara([run('KABUPATEN MALANG')],
          { spacing: { before: 0, after: 160 } }),
        ttdTable,
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  return outputPath;
}

module.exports = { generateBA };
