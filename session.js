/**
 * Manajemen sesi percakapan per user
 */

const sessions = new Map();

const DAFTAR_INSTANSI = [
  'Badan Pengawas Pemilu Kabupaten Malang',
  'Dinas Kependudukan dan Pencatatan Sipil Kabupaten Malang',
  'Komando Distrik Militer 0818 Malang-Batu',
  'Pangkalan TNI Angkatan Udara Abdulrachman Saleh',
  'Pangkalan TNI Angkatan Laut Malang',
  'Kepolisian Resor Batu',
  'Kepolisian Resor Malang',
];

function buatSesi() {
  return {
    step: 'MULAI',
    data: {
      nomorBA: '',
      triwulan: null,
      tahun: null,
      tanggalRapat: '',
      jamRapat: '',
      jumlahKecamatan: 0,
      jumlahDesaKel: 0,
      jumlahLakiLaki: 0,
      jumlahPerempuan: 0,
      masukanInstansi: [],
    },
    instansiDipilih: [],
    masukanIdx: 0,
    excelSudahUpload: false,
    // Simpan messageId keyboard sebelumnya agar bisa dihapus
    lastKeyboardMsgId: null,
  };
}

const getSesi    = (id) => { if (!sessions.has(id)) sessions.set(id, buatSesi()); return sessions.get(id); };
const resetSesi  = (id) => { sessions.set(id, buatSesi()); return sessions.get(id); };
const hapusSesi  = (id) => sessions.delete(id);

module.exports = { getSesi, resetSesi, hapusSesi, DAFTAR_INSTANSI };
