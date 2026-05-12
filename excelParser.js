/**
 * Parser file Excel Model-A Rekap PDPB
 * Membaca sheet "REKAPITULASI PDPB" dan mengambil data total
 */

const XLSX = require('xlsx');

const NAMA_SHEET = 'REKAPITULASI PDPB';

/**
 * Baca file Excel dan kembalikan data rekapitulasi
 * @param {Buffer} buffer - Buffer file Excel
 * @returns {{ jumlahKecamatan, jumlahDesaKel, jumlahLakiLaki, jumlahPerempuan, total, triwulanTeks }}
 */
function bacaExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });

  // Validasi sheet
  const sheetNames = wb.SheetNames.map(s => s.trim().toUpperCase());
  const idxSheet = sheetNames.indexOf(NAMA_SHEET.toUpperCase());
  if (idxSheet === -1) {
    throw new Error(
      `❌ Sheet "${NAMA_SHEET}" tidak ditemukan.\n` +
      `Sheet yang ada: ${wb.SheetNames.join(', ')}\n\n` +
      `Pastikan file Excel menggunakan sheet bernama *"REKAPITULASI PDPB"* (sesuai template).`
    );
  }

  const ws = wb.Sheets[wb.SheetNames[idxSheet]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Cari baris TOTAL (kolom A = 'TOTAL')
  let totalRow = null;
  let dataRows = []; // baris data kecamatan (kolom A = angka)

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const colA = row[0];

    if (typeof colA === 'string' && colA.trim().toUpperCase() === 'TOTAL') {
      totalRow = row;
    }
    if (typeof colA === 'number' && colA >= 1) {
      dataRows.push(row);
    }
  }

  if (!totalRow) {
    throw new Error(
      `❌ Baris "TOTAL" tidak ditemukan di sheet "${NAMA_SHEET}".\n` +
      `Pastikan format file Excel sesuai template.`
    );
  }

  // Kolom: A=0, B=1, C=2, D=3(L), E=4(P), F=5(L+P)
  const jumlahDesaKel    = Number(totalRow[2]) || 0;
  const jumlahLakiLaki   = Number(totalRow[3]) || 0;
  const jumlahPerempuan  = Number(totalRow[4]) || 0;
  const total            = Number(totalRow[5]) || (jumlahLakiLaki + jumlahPerempuan);
  const jumlahKecamatan  = dataRows.length;

  // Baca info triwulan dari sel A5 (biasanya ada teks "TRIWULAN ... TAHUN ...")
  let infoTriwulan = '';
  for (let i = 0; i < rows.length; i++) {
    const val = rows[i][0];
    if (typeof val === 'string' && val.toUpperCase().includes('TRIWULAN')) {
      infoTriwulan = val.trim();
      break;
    }
  }

  return {
    jumlahKecamatan,
    jumlahDesaKel,
    jumlahLakiLaki,
    jumlahPerempuan,
    total,
    infoTriwulan,
  };
}

module.exports = { bacaExcel, NAMA_SHEET };
