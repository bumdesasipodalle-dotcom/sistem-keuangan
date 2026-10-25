// ============================================================
// BUMDES KEUANGAN — Google Apps Script Backend
// Fleksibel, Universal Data Reader (Membaca format apapun)
// ============================================================

function getSS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}
  throw new Error('Spreadsheet tidak terhubung. Pastikan script ini dibuat melalui menu Extensions > Apps Script pada Google Sheets BUMDes Anda.');
}

// ============================================================
// SERVE & INCLUDE
// ============================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('BUMDes — Sistem Keuangan Terpadu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  return doGet(e);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 1. PENCARIAN SHEET FLEKSIBEL (Case-Insensitive & Alias)
// ============================================================
function findSheet(ss, aliases) {
  if (!ss) return null;
  if (typeof aliases === 'string') aliases = [aliases];

  // A. Exact match
  for (let i = 0; i < aliases.length; i++) {
    const s = ss.getSheetByName(aliases[i]);
    if (s) return s;
  }

  // B. Case-insensitive & hapus spasi/underscore
  const clean = function(str) {
    return String(str || '').toLowerCase().replace(/[\s_\-]+/g, '').trim();
  };
  const targetNorms = aliases.map(clean);
  const all = ss.getSheets();

  for (let i = 0; i < all.length; i++) {
    const norm = clean(all[i].getName());
    if (targetNorms.indexOf(norm) >= 0) return all[i];
  }

  // C. Partial match
  for (let i = 0; i < all.length; i++) {
    const norm = clean(all[i].getName());
    for (let j = 0; j < targetNorms.length; j++) {
      const t = targetNorms[j];
      if (t.length >= 3 && (norm.indexOf(t) >= 0 || t.indexOf(norm) >= 0)) {
        return all[i];
      }
    }
  }

  return null;
}

// ============================================================
// 2. PARSER ANGKA UNIVERSAL (Mendukung Rp, 1.500.000, 1500000, dll)
// ============================================================
function parseNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;

  let s = String(val).trim();
  if (!s) return 0;

  let isNegative = false;
  if (s.indexOf('(') === 0 && s.indexOf(')') === s.length - 1) {
    isNegative = true;
    s = s.slice(1, -1).trim();
  } else if (s.indexOf('-') === 0) {
    isNegative = true;
    s = s.slice(1).trim();
  }

  // Hilangkan prefix mata uang
  s = s.replace(/^(Rp|IDR)\.?\s*/i, '').trim();

  // Format ribuan titik vs desimal koma khas Indonesia
  if (s.indexOf('.') >= 0 && s.indexOf(',') >= 0) {
    if (s.lastIndexOf('.') < s.lastIndexOf(',')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.indexOf('.') >= 0) {
    const parts = s.split('.');
    if (parts.length > 2) {
      s = s.replace(/\./g, '');
    } else if (parts.length === 2 && parts[1].length === 3) {
      s = parts[0] + parts[1];
    }
  } else if (s.indexOf(',') >= 0) {
    const parts = s.split(',');
    if (parts.length > 2) {
      s = s.replace(/,/g, '');
    } else if (parts.length === 2) {
      if (parts[1].length === 3) {
        s = parts[0] + parts[1];
      } else {
        s = parts[0] + '.' + parts[1];
      }
    }
  }

  s = s.replace(/[^0-9.-]/g, '');
  const n = parseFloat(s);
  if (isNaN(n)) return 0;
  return isNegative ? -n : n;
}

// ============================================================
// 3. PARSER TANGGAL UNIVERSAL (Date, Serial Excel, DD/MM/YYYY, Teks Indo)
// ============================================================
const INDO_MONTHS = {
  januari: 1, jan: 1,
  februari: 2, pebruari: 2, feb: 2,
  maret: 3, mar: 3,
  april: 4, apr: 4,
  mei: 5, may: 5,
  juni: 6, jun: 6,
  juli: 7, jul: 7,
  agustus: 8, ags: 8, agu: 8, aug: 8,
  september: 9, sep: 9, sept: 9,
  oktober: 10, okt: 10, oct: 10,
  november: 11, nopember: 11, nov: 11,
  desember: 12, des: 12, dec: 12
};

function parseDateInfo(val) {
  if (val === null || val === undefined || val === '') return null;

  // A. Date Object Asli
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = val.getMonth() + 1;
    const d = val.getDate();
    return {
      year: y, month: m, day: d,
      formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
    };
  }

  // B. Serial Number Google Sheets/Excel (20000 - 80000)
  if (typeof val === 'number' && val > 20000 && val < 80000) {
    const dt = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(dt.getTime())) {
      const y = dt.getUTCFullYear();
      const m = dt.getUTCMonth() + 1;
      const d = dt.getUTCDate();
      return {
        year: y, month: m, day: d,
        formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
      };
    }
  }

  const s = String(val).trim();
  if (!s) return null;

  // C. YYYY-MM-DD atau YYYY/MM/DD
  let match = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) {
    const y = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const d = parseInt(match[3], 10);
    return {
      year: y, month: m, day: d,
      formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
    };
  }

  // D. DD/MM/YYYY atau DD-MM-YYYY
  match = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (match) {
    const d = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const y = parseInt(match[3], 10);
    return {
      year: y, month: m, day: d,
      formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
    };
  }

  // E. Format teks tanggal Indonesia: '15 Januari 2024', '15-Okt-2025'
  match = s.match(/^(\d{1,2})\s*[-/ ]\s*([a-zA-Z]+)\s*[-/ ]\s*(\d{4})/);
  if (match) {
    const d = parseInt(match[1], 10);
    const m = INDO_MONTHS[match[2].toLowerCase()];
    const y = parseInt(match[3], 10);
    if (m) {
      return {
        year: y, month: m, day: d,
        formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
      };
    }
  }

  // F. Hanya Bulan & Tahun: 'Januari 2024'
  match = s.match(/^([a-zA-Z]+)\s*[-/ ]\s*(\d{4})/);
  if (match) {
    const m = INDO_MONTHS[match[1].toLowerCase()];
    const y = parseInt(match[2], 10);
    if (m) {
      return {
        year: y, month: m, day: 1,
        formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-01'
      };
    }
  }

  // G. MM/YYYY atau MM-YYYY
  match = s.match(/^(\d{1,2})[-/](\d{4})/);
  if (match) {
    const m = parseInt(match[1], 10);
    const y = parseInt(match[2], 10);
    if (m >= 1 && m <= 12) {
      return {
        year: y, month: m, day: 1,
        formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-01'
      };
    }
  }

  // H. Hanya Tahun 4 Digit: '2024'
  match = s.match(/^(\d{4})$/);
  if (match) {
    const y = parseInt(match[1], 10);
    if (y > 2000 && y < 2100) {
      return {
        year: y, month: 1, day: 1,
        formatted: y + '-01-01'
      };
    }
  }

  // I. Native Date parse (hanya jika ada pemisah angka)
  if (/\d+/.test(s) && (s.indexOf('-') >= 0 || s.indexOf('/') >= 0 || s.indexOf('.') >= 0)) {
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      const y = dt.getFullYear();
      const m = dt.getMonth() + 1;
      const d = dt.getDate();
      return {
        year: y, month: m, day: d,
        formatted: y + '-' + String(m < 10 ? '0' + m : m) + '-' + String(d < 10 ? '0' + d : d)
      };
    }
  }

  return null;
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

// ============================================================
// 4. PEMBACA DATA SHEET PINTAR (Membaca format header & baris apapun)
// ============================================================
function readSheetData(sheet, typePrefix) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];

  const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (!rawValues || rawValues.length === 0) return [];

  // A. Cari baris mana yang merupakan header
  const keywords = ['tanggal','tgl','date','waktu','jumlah','nominal','total','debet','kredit','pemasukan','pengeluaran','uraian','keterangan','deskripsi','unit','kategori','no','id','saldo','masuk','keluar'];
  
  let headerRowIdx = 0;
  let maxMatches = -1;

  for (let r = 0; r < Math.min(5, rawValues.length); r++) {
    const row = rawValues[r];
    let matches = 0;
    for (let c = 0; c < row.length; c++) {
      const cStr = String(row[c] || '').toLowerCase().replace(/[\s_\-]+/g, '');
      if (keywords.some(function(k) { return cStr.indexOf(k) >= 0; })) matches++;
    }
    if (matches > maxMatches) {
      maxMatches = matches;
      headerRowIdx = r;
    }
  }

  const rawHeaders = rawValues[headerRowIdx];
  const cleanHeaders = rawHeaders.map(function(h) { return String(h || '').trim(); });
  const normHeaders = cleanHeaders.map(function(h) { return h.toLowerCase().replace(/[\s_\-]+/g, ''); });

  // B. Proses baris data setelah header
  const results = [];

  for (let r = headerRowIdx + 1; r < rawValues.length; r++) {
    const row = rawValues[r];

    // Cek apakah ada sel yang berisi data (bukan baris kosong melompong)
    const hasContent = row.some(function(cell) {
      return cell !== null && cell !== undefined && String(cell).trim() !== '';
    });
    if (!hasContent) continue;

    // Abaikan baris footer / total akumulasi di spreadsheet
    const firstCellText = String(row[0] || '').trim().toUpperCase();
    if (firstCellText === 'TOTAL' || firstCellText === 'JUMLAH' || firstCellText.indexOf('TOTAL ') === 0) {
      continue;
    }

    const obj = {};

    // 1. Simpan field asli
    cleanHeaders.forEach(function(h, colIdx) {
      if (h) {
        obj[h] = row[colIdx] !== undefined && row[colIdx] !== null ? row[colIdx] : '';
      }
    });

    // 2. Normalisasi kolom ke properti baku aplikasi
    // ID
    let idVal = '';
    const idIdx = normHeaders.findIndex(function(h) { return h === 'id' || h === 'no' || h === 'nomor' || h === 'kode'; });
    if (idIdx >= 0 && row[idIdx]) {
      idVal = String(row[idIdx]).trim();
    }
    obj.ID = idVal || (typePrefix + '-' + (r + 1));
    obj._rowIndex = r + 1;

    // Tanggal
    let tglVal = '';
    const tglIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('tanggal') >= 0 || h.indexOf('tgl') >= 0 || h === 'date' || h === 'waktu';
    });
    if (tglIdx >= 0 && row[tglIdx]) {
      tglVal = row[tglIdx];
    } else {
      for (let c = 0; c < row.length; c++) {
        if (row[c] instanceof Date || parseDateInfo(row[c])) {
          tglVal = row[c];
          break;
        }
      }
    }
    const pDate = parseDateInfo(tglVal);
    obj.Tanggal = pDate ? pDate.formatted : String(tglVal || '');
    obj._dateParsed = pDate;

    // Jumlah / Nominal
    let jmlVal = 0;
    const jmlIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('jumlah') >= 0 || h.indexOf('nominal') >= 0 || h.indexOf('total') >= 0 || 
        h.indexOf('nilai') >= 0 || h.indexOf('besar') >= 0 || h.indexOf('pemasukan') >= 0 || 
        h.indexOf('pengeluaran') >= 0 || h.indexOf('debet') >= 0 || h.indexOf('kredit') >= 0 || 
        h.indexOf('masuk') >= 0 || h.indexOf('keluar') >= 0 || h.indexOf('subtotal') >= 0 || h.indexOf('harga') >= 0;
    });
    if (jmlIdx >= 0 && row[jmlIdx] !== undefined && row[jmlIdx] !== '') {
      jmlVal = parseNumber(row[jmlIdx]);
    } else {
      for (let c = 0; c < row.length; c++) {
        if (c !== idIdx && c !== tglIdx) {
          const num = parseNumber(row[c]);
          if (num > 0) {
            jmlVal = num;
            break;
          }
        }
      }
    }
    obj.Jumlah = jmlVal;

    // Keterangan / Uraian
    let ketVal = '';
    const ketIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('keterangan') >= 0 || h.indexOf('uraian') >= 0 || h.indexOf('deskripsi') >= 0 || 
        h.indexOf('catatan') >= 0 || h.indexOf('rincian') >= 0 || h === 'nama' || h === 'namabarang';
    });
    if (ketIdx >= 0 && row[ketIdx]) {
      ketVal = String(row[ketIdx]).trim();
    } else {
      for (let c = 0; c < row.length; c++) {
        if (c !== idIdx && c !== tglIdx && c !== jmlIdx && typeof row[c] === 'string' && row[c].trim()) {
          ketVal = row[c].trim();
          break;
        }
      }
    }
    obj.Keterangan = ketVal || '-';

    // Unit Usaha
    let unitVal = '';
    const unitIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('unitusaha') >= 0 || h === 'unit' || h === 'namaunit' || h.indexOf('usaha') >= 0 || h.indexOf('divisi') >= 0;
    });
    if (unitIdx >= 0 && row[unitIdx]) {
      unitVal = String(row[unitIdx]).trim();
    }
    obj.Unit_Usaha = unitVal || 'Umum';

    // Kategori
    let katVal = '';
    const katIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('kategori') >= 0 || h.indexOf('jenis') >= 0 || h.indexOf('pos') >= 0 || h.indexOf('akun') >= 0 || h.indexOf('golongan') >= 0;
    });
    if (katIdx >= 0 && row[katIdx]) {
      katVal = String(row[katIdx]).trim();
    }
    obj.Kategori = katVal || 'Lain-lain';

    // Metode Bayar
    let mbVal = '';
    const mbIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('metode') >= 0 || h.indexOf('bayar') >= 0 || h.indexOf('via') >= 0 || h.indexOf('rekening') >= 0;
    });
    if (mbIdx >= 0 && row[mbIdx]) {
      mbVal = String(row[mbIdx]).trim();
    }
    obj.Metode_Bayar = mbVal || 'Tunai';

    // Vendor / Penerima
    let vendorVal = '';
    const vendorIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('vendor') >= 0 || h.indexOf('penerima') >= 0 || h.indexOf('supplier') >= 0 || h.indexOf('toko') >= 0 || h.indexOf('pihakketiga') >= 0;
    });
    if (vendorIdx >= 0 && row[vendorIdx]) {
      vendorVal = String(row[vendorIdx]).trim();
    }
    obj.Vendor = vendorVal || '-';
    obj.Penerima = vendorVal || '-';

    // Bukti
    let buktiVal = '';
    const buktiIdx = normHeaders.findIndex(function(h) {
      return h.indexOf('bukti') >= 0 || h.indexOf('nota') >= 0 || h.indexOf('kuitansi') >= 0 || h.indexOf('faktur') >= 0;
    });
    if (buktiIdx >= 0 && row[buktiIdx]) {
      buktiVal = String(row[buktiIdx]).trim();
    }
    obj.Bukti = buktiVal || '-';

    // Buku Kas (Masuk, Keluar, Saldo)
    const masukIdx = normHeaders.findIndex(function(h) { return h === 'masuk' || h === 'debet' || h === 'debit' || h === 'penerimaan'; });
    const keluarIdx = normHeaders.findIndex(function(h) { return h === 'keluar' || h === 'kredit' || h === 'credit' || h === 'pengeluaran'; });
    const saldoIdx = normHeaders.findIndex(function(h) { return h === 'saldo' || h === 'sisasaldo'; });

    if (masukIdx >= 0) obj.Masuk = parseNumber(row[masukIdx]);
    if (keluarIdx >= 0) obj.Keluar = parseNumber(row[keluarIdx]);
    if (saldoIdx >= 0) obj.Saldo = parseNumber(row[saldoIdx]);

    // Anggaran
    const angIdx = normHeaders.findIndex(function(h) { return h === 'anggaran' || h === 'target' || h === 'plafon'; });
    const realIdx = normHeaders.findIndex(function(h) { return h === 'realisasi' || h === 'aktual'; });
    if (angIdx >= 0) obj.Anggaran = parseNumber(row[angIdx]);
    if (realIdx >= 0) obj.Realisasi = parseNumber(row[realIdx]);

    results.push(obj);
  }

  return results;
}

// ============================================================
// 5. FILTER FLEKSIBEL (Tahun, Bulan, Unit Usaha)
// ============================================================
function applyFilter(rows, filter) {
  if (!filter) return rows;
  let res = rows;

  // Filter Tahun
  if (filter.tahun && filter.tahun !== 'all' && filter.tahun !== 'semua' && filter.tahun !== '') {
    const tTarget = parseInt(filter.tahun);
    if (!isNaN(tTarget)) {
      res = res.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        if (dp && dp.year) {
          return dp.year === tTarget;
        }
        return String(r.Tanggal || '').indexOf(String(tTarget)) >= 0;
      });
    }
  }

  // Filter Bulan
  if (filter.bulan && filter.bulan !== 'all' && filter.bulan !== 'semua' && filter.bulan !== '') {
    const bTarget = parseInt(filter.bulan);
    if (!isNaN(bTarget)) {
      res = res.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        return dp && dp.month === bTarget;
      });
    }
  }

  // Filter Unit Usaha
  if (filter.unit && filter.unit !== 'all' && filter.unit !== 'semua' && filter.unit !== '') {
    const uTarget = String(filter.unit).toLowerCase().trim();
    res = res.filter(function(r) {
      const u = String(r.Unit_Usaha || '').toLowerCase().trim();
      return u === uTarget || u.indexOf(uTarget) >= 0 || uTarget.indexOf(u) >= 0;
    });
  }

  return res;
}

// ============================================================
// SETUP SPREADSHEET (Inisialisasi tabel default)
// ============================================================
function setupSpreadsheet() {
  const ss = getSS();
  const sheets = [
    { name: 'PENGGUNA',      headers: ['ID','Username','Password','Nama_Lengkap','Role','Status','Dibuat'] },
    { name: 'MODAL',         headers: ['ID','Tanggal','Sumber_Modal','Keterangan','Jumlah','Jenis','Status','Dibuat'] },
    { name: 'PENDAPATAN',    headers: ['ID','Tanggal','Unit_Usaha','Kategori','Keterangan','Jumlah','Metode_Bayar','Bukti','Dibuat'] },
    { name: 'BIAYA_VARIABEL',headers: ['ID','Tanggal','Unit_Usaha','Kategori','Keterangan','Jumlah','Vendor','Bukti','Dibuat'] },
    { name: 'BIAYA_LAINNYA', headers: ['ID','Tanggal','Kategori','Keterangan','Jumlah','Penerima','Bukti','Dibuat'] },
    { name: 'ANGGARAN',      headers: ['ID','Tahun','Bulan','Unit_Usaha','Kategori','Uraian','Anggaran','Realisasi','Dibuat'] },
    { name: 'ALOKASI_PAD',   headers: ['ID','Tahun','Bulan','Total_Laba','PAD_Desa_Pct','PAD_Desa_Nominal','Pengembangan_Pct','Pengembangan_Nominal','Pengurus_Pct','Pengurus_Nominal','Cadangan_Pct','Cadangan_Nominal','Sosial_Pct','Sosial_Nominal','Tanggal_Alokasi','Dibuat'] },
    { name: 'UNIT_USAHA',    headers: ['ID','Nama_Unit','Jenis_Usaha','Penanggung_Jawab','Modal_Awal','Status','Tanggal_Berdiri','Keterangan','Dibuat'] },
    { name: 'KAS',           headers: ['ID','Tanggal','Jenis','Referensi_ID','Keterangan','Masuk','Keluar','Saldo','Dibuat'] },
    { name: 'PENGATURAN',    headers: ['Kunci','Nilai','Keterangan'] }
  ];

  sheets.forEach(function(def) {
    let sheet = findSheet(ss, [def.name]);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sheet.getRange(1, 1, 1, def.headers.length).setBackground('#0f4c81').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  // Pengaturan default jika belum ada
  const ps = findSheet(ss, ['PENGATURAN']);
  if (ps && ps.getLastRow() <= 1) {
    const defaultPengaturan = [
      ['nama_bumdes',     'BUMDes Maju Bersama',      'Nama BUMDes'],
      ['nama_desa',       'Desa Sejahtera',            'Nama Desa'],
      ['kecamatan',       'Kecamatan Makmur',          'Nama Kecamatan'],
      ['kabupaten',       'Kabupaten Sejahtera',       'Nama Kabupaten'],
      ['tahun_aktif',     String(new Date().getFullYear()), 'Tahun Buku Aktif'],
      ['pad_desa_pct',    '40',                        'Persen alokasi PAD ke Desa (%)'],
      ['pengembangan_pct','30',                        'Persen Pengembangan Usaha (%)'],
      ['pengurus_pct',    '20',                        'Persen untuk Pengurus (%)'],
      ['cadangan_pct',    '5',                         'Persen Dana Cadangan (%)'],
      ['sosial_pct',      '5',                         'Persen Dana Sosial (%)'],
      ['direktur',        'H. Ahmad Syafi\'i',         'Nama Direktur BUMDes'],
      ['sekretaris',      'Dewi Sartika, S.E.',        'Nama Sekretaris BUMDes'],
      ['bendahara',       'Budi Santoso, A.Md.',       'Nama Bendahara BUMDes'],
      ['kepala_desa',     'Drs. H. Mulyono',           'Nama Kepala Desa']
    ];
    ps.getRange(2, 1, defaultPengaturan.length, 3).setValues(defaultPengaturan);
  }

  // Pengguna default jika belum ada
  const usrSheet = findSheet(ss, ['PENGGUNA']);
  if (usrSheet && usrSheet.getLastRow() <= 1) {
    const defaultUsers = [
      [generateId('USR'), 'admin', 'admin123', 'Administrator Utama', 'Admin', 'Aktif', new Date()],
      [generateId('bendahara'), 'bendahara', 'bendahara123', 'Bendahara BUMDes', 'Bendahara', 'Aktif', new Date()],
      [generateId('pengawas'), 'pengawas', 'pengawas123', 'Pengawas BPD / Desa', 'Pengawas', 'Aktif', new Date()]
    ];
    usrSheet.getRange(2, 1, defaultUsers.length, 7).setValues(defaultUsers);
  }

  return { success: true, message: 'Database BUMDes berhasil disinkronkan!' };
}

// ============================================================
// AUTHENTICATION & LOGIN
// ============================================================
function loginUser(username, password) {
  try {
    username = (username || '').trim().toLowerCase();
    password = (password || '').trim();

    if (!username || !password) return { success: false, message: 'Username dan kata sandi wajib diisi!' };

    const ss = getSS();
    let sheet = findSheet(ss, ['PENGGUNA', 'Pengguna', 'Users']);
    if (!sheet || sheet.getLastRow() <= 1) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['PENGGUNA']);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    const uIdx = headers.indexOf('username');
    const pIdx = headers.indexOf('password');
    const nIdx = headers.indexOf('nama_lengkap') >= 0 ? headers.indexOf('nama_lengkap') : headers.indexOf('nama');
    const rIdx = headers.indexOf('role');
    const sIdx = headers.indexOf('status');
    const idIdx = headers.indexOf('id');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[uIdx] || '').trim().toLowerCase() === username) {
        if (String(row[pIdx] || '').trim() !== password) {
          return { success: false, message: 'Kata sandi tidak sesuai!' };
        }
        if (String(row[sIdx] || 'Aktif').trim() !== 'Aktif') {
          return { success: false, message: 'Akun Anda berstatus non-aktif. Hubungi Admin.' };
        }
        return {
          success: true,
          message: 'Login berhasil!',
          user: {
            id: row[idIdx] || ('USR-' + i),
            username: String(row[uIdx]),
            nama: String(row[nIdx] || row[uIdx]),
            role: String(row[rIdx] || 'Staff')
          }
        };
      }
    }

    return { success: false, message: 'Pengguna dengan username "' + username + '" tidak ditemukan!' };
  } catch (e) {
    return { success: false, message: 'Gagal verifikasi login: ' + e.message };
  }
}

function getPengguna() {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['PENGGUNA', 'Pengguna', 'Users']);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return readSheetData(sheet, 'USR');
  } catch (e) {
    return [];
  }
}

function savePengguna(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['PENGGUNA', 'Pengguna', 'Users']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['PENGGUNA']);
    }
    const uname = (data.Username || '').trim().toLowerCase();
    const rows = sheet.getDataRange().getValues();

    if (data.ID) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.ID)) {
          const pass = data.Password ? data.Password : rows[i][2];
          sheet.getRange(i + 1, 1, 1, 7).setValues([[
            data.ID, uname, pass, data.Nama_Lengkap, data.Role, data.Status || 'Aktif', rows[i][6] || new Date()
          ]]);
          return { success: true, id: data.ID };
        }
      }
    }

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === uname) {
        return { success: false, message: 'Username sudah digunakan!' };
      }
    }

    const id = generateId('USR');
    sheet.appendRow([id, uname, data.Password || '123456', data.Nama_Lengkap, data.Role, data.Status || 'Aktif', new Date()]);
    return { success: true, id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function hapusPengguna(id) {
  return deleteRowById('PENGGUNA', id);
}

function ubahPassword(id, passLama, passBaru) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['PENGGUNA', 'Pengguna', 'Users']);
    if (!sheet) return { success: false, message: 'Sheet pengguna tidak ditemukan' };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        if (String(rows[i][2]).trim() !== String(passLama).trim()) {
          return { success: false, message: 'Kata sandi saat ini tidak cocok!' };
        }
        sheet.getRange(i + 1, 3).setValue(passBaru);
        return { success: true, message: 'Kata sandi berhasil diubah!' };
      }
    }
    return { success: false, message: 'Pengguna tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// PENGATURAN & AUTO-DETEKSI TAHUN AKTIF
// ============================================================
function getPengaturan() {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['PENGATURAN', 'Pengaturan', 'Setting', 'Config']);
    const result = {};

    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          result[String(data[i][0]).trim()] = String(data[i][1] !== undefined ? data[i][1] : '').trim();
        }
      }
    }

    // Auto-deteksi SEMUA TAHUN yang ada transaksinya di spreadsheet
    const detectedYears = new Set();
    const curYear = new Date().getFullYear();
    detectedYears.add(curYear);

    const sheetsToScan = [
      ['PENDAPATAN', 'Pendapatan', 'Pemasukan'],
      ['MODAL', 'Modal', 'Penyertaan Modal'],
      ['BIAYA_VARIABEL', 'Biaya Variabel', 'HPP'],
      ['BIAYA_LAINNYA', 'Biaya Lainnya', 'Biaya Operasional'],
      ['KAS', 'Kas', 'Buku Kas']
    ];

    sheetsToScan.forEach(function(aliases) {
      const sh = findSheet(ss, aliases);
      if (sh) {
        const rows = readSheetData(sh, 'SCAN');
        rows.forEach(function(r) {
          const dp = r._dateParsed || parseDateInfo(r.Tanggal);
          if (dp && dp.year > 2000 && dp.year < 2100) {
            detectedYears.add(dp.year);
          }
        });
      }
    });

    const sortedYears = Array.from(detectedYears).sort(function(a, b) { return b - a; });
    result.available_years = sortedYears;

    if (!result.tahun_aktif) {
      result.tahun_aktif = String(sortedYears[0] || curYear);
    }

    return result;
  } catch (e) {
    return {
      nama_bumdes: 'BUMDes',
      tahun_aktif: String(new Date().getFullYear()),
      available_years: [new Date().getFullYear()]
    };
  }
}

function savePengaturan(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['PENGATURAN', 'Pengaturan']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['PENGATURAN']);
    }
    const existing = sheet.getDataRange().getValues();

    Object.keys(data).forEach(function(key) {
      let found = false;
      for (let i = 1; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === key) {
          sheet.getRange(i + 1, 2).setValue(String(data[key]));
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, String(data[key]), '']);
      }
    });

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// MASTER UNIT USAHA
// ============================================================
function getUnitUsaha() {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['UNIT_USAHA', 'Unit Usaha', 'UNIT USAHA', 'Unit', 'Master Unit']);
    let units = [];
    if (sheet) {
      units = readSheetData(sheet, 'UU');
    }

    // Auto-ekstrak nama unit dari transaksi pendapatan & biaya jika sheet master kosong
    if (units.length === 0) {
      const known = new Set();
      const pend = getPendapatan({});
      const bv   = getBiayaVariabel({});

      [].concat(pend, bv).forEach(function(r) {
        const u = String(r.Unit_Usaha || '').trim();
        if (u && u !== 'Umum' && u !== '-') known.add(u);
      });

      units = Array.from(known).map(function(u, i) {
        return {
          ID: 'UU-' + (i + 1),
          Nama_Unit: u,
          Jenis_Usaha: 'Usaha Desa',
          Penanggung_Jawab: '-',
          Modal_Awal: 0,
          Status: 'Aktif'
        };
      });
    }

    return units;
  } catch (e) {
    return [];
  }
}

function saveUnitUsaha(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['UNIT_USAHA', 'Unit Usaha']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['UNIT_USAHA']);
    }
    const tgl = data.Tanggal_Berdiri ? new Date(data.Tanggal_Berdiri) : new Date();
    const modalAwal = parseNumber(data.Modal_Awal);

    if (data.ID) {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(data.ID).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            data.ID, data.Nama_Unit, data.Jenis_Usaha || 'Usaha Desa', data.Penanggung_Jawab || '-',
            modalAwal, data.Status || 'Aktif', tgl, data.Keterangan || '', rows[i][8] || new Date()
          ]]);
          return { success: true, id: data.ID };
        }
      }
    }
    const id = generateId('UU');
    sheet.appendRow([id, data.Nama_Unit, data.Jenis_Usaha || 'Usaha Desa', data.Penanggung_Jawab || '-', modalAwal, data.Status || 'Aktif', tgl, data.Keterangan || '', new Date()]);
    return { success: true, id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteUnitUsaha(id) {
  return deleteRowById('UNIT_USAHA', id);
}

// ============================================================
// PENYERTAAN MODAL
// ============================================================
function getModal(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['MODAL', 'Modal', 'Penyertaan Modal', 'PENYERTAAN_MODAL']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'MDL');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function saveModal(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['MODAL', 'Modal', 'Penyertaan Modal']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['MODAL']);
    }
    const tgl = data.Tanggal ? new Date(data.Tanggal) : new Date();
    const jml = parseNumber(data.Jumlah);
    let id = data.ID;

    if (id) {
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 8).setValues([[
            id, tgl, data.Sumber_Modal || 'Pemerintah Desa', data.Keterangan || '', jml, data.Jenis || 'Uang', data.Status || 'Aktif', rows[i][7] || new Date()
          ]]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        id = generateId('MDL');
        sheet.appendRow([id, tgl, data.Sumber_Modal || 'Pemerintah Desa', data.Keterangan || '', jml, data.Jenis || 'Uang', data.Status || 'Aktif', new Date()]);
      }
    } else {
      id = generateId('MDL');
      sheet.appendRow([id, tgl, data.Sumber_Modal || 'Pemerintah Desa', data.Keterangan || '', jml, data.Jenis || 'Uang', data.Status || 'Aktif', new Date()]);
    }

    syncKas(id, tgl, 'Modal', '[Modal] ' + (data.Sumber_Modal || '') + ' - ' + (data.Keterangan || ''), jml, 0);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteModal(id) {
  removeKasByRefId(id);
  return deleteRowById('MODAL', id);
}

// ============================================================
// PENDAPATAN
// ============================================================
function getPendapatan(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['PENDAPATAN', 'Pendapatan', 'Pemasukan', 'Penjualan', 'Pendapatan Usaha']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'PND');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function savePendapatan(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['PENDAPATAN', 'Pendapatan', 'Pemasukan']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['PENDAPATAN']);
    }
    const tgl = data.Tanggal ? new Date(data.Tanggal) : new Date();
    const jml = parseNumber(data.Jumlah);
    let id = data.ID;

    if (id) {
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '',
            jml, data.Metode_Bayar || 'Tunai', data.Bukti || '', rows[i][8] || new Date()
          ]]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        id = generateId('PND');
        sheet.appendRow([id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '', jml, data.Metode_Bayar || 'Tunai', data.Bukti || '', new Date()]);
      }
    } else {
      id = generateId('PND');
      sheet.appendRow([id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '', jml, data.Metode_Bayar || 'Tunai', data.Bukti || '', new Date()]);
    }

    const ketKas = '[' + (data.Unit_Usaha || 'Umum') + '] ' + (data.Kategori || '') + ' - ' + (data.Keterangan || '');
    syncKas(id, tgl, 'Pendapatan', ketKas, jml, 0);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deletePendapatan(id) {
  removeKasByRefId(id);
  return deleteRowById('PENDAPATAN', id);
}

// ============================================================
// BIAYA VARIABEL (HPP)
// ============================================================
function getBiayaVariabel(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['BIAYA_VARIABEL', 'Biaya Variabel', 'BIAYA VARIABEL', 'HPP', 'Harga Pokok', 'Biaya Produksi']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'BV');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function saveBiayaVariabel(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['BIAYA_VARIABEL', 'Biaya Variabel', 'HPP']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['BIAYA_VARIABEL']);
    }
    const tgl = data.Tanggal ? new Date(data.Tanggal) : new Date();
    const jml = parseNumber(data.Jumlah);
    let id = data.ID;

    if (id) {
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '',
            jml, data.Vendor || '', data.Bukti || '', rows[i][8] || new Date()
          ]]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        id = generateId('BV');
        sheet.appendRow([id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '', jml, data.Vendor || '', data.Bukti || '', new Date()]);
      }
    } else {
      id = generateId('BV');
      sheet.appendRow([id, tgl, data.Unit_Usaha || '', data.Kategori || '', data.Keterangan || '', jml, data.Vendor || '', data.Bukti || '', new Date()]);
    }

    const ketKas = '[HPP: ' + (data.Unit_Usaha || 'Umum') + '] ' + (data.Kategori || '') + ' - ' + (data.Keterangan || '');
    syncKas(id, tgl, 'Biaya Variabel', ketKas, 0, jml);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteBiayaVariabel(id) {
  removeKasByRefId(id);
  return deleteRowById('BIAYA_VARIABEL', id);
}

// ============================================================
// BIAYA LAINNYA (OPERASIONAL)
// ============================================================
function getBiayaLainnya(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['BIAYA_LAINNYA', 'Biaya Lainnya', 'BIAYA LAINNYA', 'Biaya Operasional', 'BIAYA OPERASIONAL', 'Operasional', 'Beban']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'BL');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function saveBiayaLainnya(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['BIAYA_LAINNYA', 'Biaya Lainnya', 'Biaya Operasional']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['BIAYA_LAINNYA']);
    }
    const tgl = data.Tanggal ? new Date(data.Tanggal) : new Date();
    const jml = parseNumber(data.Jumlah);
    let id = data.ID;

    if (id) {
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 8).setValues([[
            id, tgl, data.Kategori || '', data.Keterangan || '', jml,
            data.Penerima || '', data.Bukti || '', rows[i][7] || new Date()
          ]]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        id = generateId('BL');
        sheet.appendRow([id, tgl, data.Kategori || '', data.Keterangan || '', jml, data.Penerima || '', data.Bukti || '', new Date()]);
      }
    } else {
      id = generateId('BL');
      sheet.appendRow([id, tgl, data.Kategori || '', data.Keterangan || '', jml, data.Penerima || '', data.Bukti || '', new Date()]);
    }

    const ketKas = '[Operasional] ' + (data.Kategori || '') + ' - ' + (data.Keterangan || '');
    syncKas(id, tgl, 'Biaya Operasional', ketKas, 0, jml);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteBiayaLainnya(id) {
  removeKasByRefId(id);
  return deleteRowById('BIAYA_LAINNYA', id);
}

// ============================================================
// RAB ANGGARAN
// ============================================================
function getAnggaran(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['ANGGARAN', 'Anggaran', 'RAB', 'Rencana Anggaran']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'ANG');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function saveAnggaran(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['ANGGARAN', 'Anggaran']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['ANGGARAN']);
    }
    const angg = parseNumber(data.Anggaran);
    const real = parseNumber(data.Realisasi);
    let id = data.ID;

    if (id) {
      const rows = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim() || (data._rowIndex && i + 1 === data._rowIndex)) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            id, data.Tahun, data.Bulan, data.Unit_Usaha || '', data.Kategori || '',
            data.Uraian || '', angg, real, rows[i][8] || new Date()
          ]]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        id = generateId('ANG');
        sheet.appendRow([id, data.Tahun, data.Bulan, data.Unit_Usaha || '', data.Kategori || '', data.Uraian || '', angg, real, new Date()]);
      }
    } else {
      id = generateId('ANG');
      sheet.appendRow([id, data.Tahun, data.Bulan, data.Unit_Usaha || '', data.Kategori || '', data.Uraian || '', angg, real, new Date()]);
    }
    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteAnggaran(id) {
  return deleteRowById('ANGGARAN', id);
}

// ============================================================
// ALOKASI PAD (BAGI HASIL)
// ============================================================
function getAlokasiPAD(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['ALOKASI_PAD', 'Alokasi PAD', 'ALOKASI PAD', 'PAD']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'PAD');
    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function saveAlokasiPAD(data) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['ALOKASI_PAD', 'Alokasi PAD']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['ALOKASI_PAD']);
    }
    const id = data.ID || generateId('PAD');
    const tgl = data.Tanggal_Alokasi ? new Date(data.Tanggal_Alokasi) : new Date();

    const row = [
      id, data.Tahun, data.Bulan,
      parseNumber(data.Total_Laba),
      parseNumber(data.PAD_Desa_Pct), parseNumber(data.PAD_Desa_Nominal),
      parseNumber(data.Pengembangan_Pct), parseNumber(data.Pengembangan_Nominal),
      parseNumber(data.Pengurus_Pct), parseNumber(data.Pengurus_Nominal),
      parseNumber(data.Cadangan_Pct), parseNumber(data.Cadangan_Nominal),
      parseNumber(data.Sosial_Pct), parseNumber(data.Sosial_Nominal),
      tgl, new Date()
    ];

    sheet.appendRow(row);

    const totKeluar = parseNumber(data.PAD_Desa_Nominal) + parseNumber(data.Sosial_Nominal);
    if (totKeluar > 0) {
      syncKas(id, tgl, 'Alokasi PAD', 'Bagi Hasil PAD & Sosial ' + data.Tahun, 0, totKeluar);
    }

    return { success: true, id: id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteAlokasiPAD(id) {
  removeKasByRefId(id);
  return deleteRowById('ALOKASI_PAD', id);
}

// ============================================================
// BUKU KAS UMUM (KAS)
// ============================================================
function getKas(filter) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['KAS', 'Kas', 'Buku Kas', 'Buku Kas Umum', 'BKU', 'BUKU_KAS']);
    if (!sheet) return [];
    const rows = readSheetData(sheet, 'KAS');

    let running = 0;
    rows.forEach(function(r) {
      const msk = parseNumber(r.Masuk);
      const klr = parseNumber(r.Keluar);
      running += (msk - klr);
      if (r.Saldo === undefined || r.Saldo === null || isNaN(r.Saldo)) {
        r.Saldo = running;
      }
    });

    return applyFilter(rows, filter);
  } catch (e) {
    return [];
  }
}

function syncKas(refId, tanggal, jenis, keterangan, masuk, keluar) {
  try {
    const ss = getSS();
    let sheet = findSheet(ss, ['KAS', 'Kas', 'Buku Kas']);
    if (!sheet) {
      setupSpreadsheet();
      sheet = findSheet(ss, ['KAS']);
    }

    masuk = parseNumber(masuk);
    keluar = parseNumber(keluar);
    const rows = sheet.getDataRange().getValues();
    let refIdx = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][3]).trim() === String(refId).trim()) {
        refIdx = i;
        break;
      }
    }

    if (refIdx > 0) {
      sheet.getRange(refIdx + 1, 2).setValue(tanggal);
      sheet.getRange(refIdx + 1, 3).setValue(jenis);
      sheet.getRange(refIdx + 1, 5).setValue(keterangan);
      sheet.getRange(refIdx + 1, 6).setValue(masuk);
      sheet.getRange(refIdx + 1, 7).setValue(keluar);
    } else {
      const id = generateId('KAS');
      sheet.appendRow([id, tanggal, jenis, refId, keterangan, masuk, keluar, 0, new Date()]);
    }

    recalcKasSaldo(sheet);
  } catch (e) {
    console.error('Gagal sinkron kas:', e);
  }
}

function removeKasByRefId(refId) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, ['KAS', 'Kas', 'Buku Kas']);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][3]).trim() === String(refId).trim()) {
        sheet.deleteRow(i + 1);
        recalcKasSaldo(sheet);
        return;
      }
    }
  } catch (e) {}
}

function recalcKasSaldo(sheet) {
  try {
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    let running = 0;
    const saldoCol = 8;
    for (let i = 1; i < rows.length; i++) {
      const m = parseNumber(rows[i][5]);
      const k = parseNumber(rows[i][6]);
      running += (m - k);
      sheet.getRange(i + 1, saldoCol).setValue(running);
    }
  } catch (e) {}
}

// ============================================================
// DASHBOARD OPTIMIZED (Universal, In-Memory, Mendukung Semua Tahun)
// ============================================================
function getDashboardData(tahun) {
  try {
    const ss = getSS();

    const allPend  = getPendapatan({});
    const allBv    = getBiayaVariabel({});
    const allBl    = getBiayaLainnya({});
    const allModal = getModal({});
    const allKas   = getKas({});
    const allUnits = getUnitUsaha();

    const grandPend  = allPend.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const grandBv    = allBv.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const grandBl    = allBl.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const grandBiaya = grandBv + grandBl;
    const grandModal = allModal.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalSemuaTransaksi = allPend.length + allBv.length + allBl.length + allModal.length;

    const isAllYears = (!tahun || tahun === 'all' || tahun === 'semua');
    let targetYear = isAllYears ? 'all' : (parseInt(tahun) || new Date().getFullYear());

    let pend = isAllYears ? allPend : applyFilter(allPend, { tahun: targetYear });
    let bv   = isAllYears ? allBv   : applyFilter(allBv,   { tahun: targetYear });
    let bl   = isAllYears ? allBl   : applyFilter(allBl,   { tahun: targetYear });
    let modal= isAllYears ? allModal: applyFilter(allModal,{ tahun: targetYear });

    const totalPend  = pend.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBv    = bv.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBl    = bl.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBiaya = totalBv + totalBl;
    const labaBersih = totalPend - totalBiaya;
    const totalModal = modal.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);

    let saldoKas = 0;
    if (allKas.length > 0) {
      saldoKas = parseNumber(allKas[allKas.length - 1].Saldo);
      if (saldoKas === 0 && (allKas[allKas.length - 1].Masuk || allKas[allKas.length - 1].Keluar)) {
        const totM = allKas.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0); }, 0);
        const totK = allKas.reduce(function(s, r) { return s + (parseNumber(r.Keluar) || 0); }, 0);
        saldoKas = totM - totK;
      }
    } else {
      saldoKas = (isAllYears ? grandModal + grandPend - grandBiaya : totalModal + totalPend - totalBiaya);
    }

    // Hitung tren 12 bulan
    const trend = [];
    for (let m = 1; m <= 12; m++) {
      const pM = pend.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        return dp && dp.month === m;
      }).reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);

      const bvM = bv.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        return dp && dp.month === m;
      }).reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);

      const blM = bl.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        return dp && dp.month === m;
      }).reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);

      trend.push({
        bulan: m,
        pendapatan: pM,
        biaya: bvM + blM,
        laba: pM - (bvM + blM)
      });
    }

    return {
      summary: {
        totalPendapatan: totalPend,
        totalBiayaVariabel: totalBv,
        totalBiayaLainnya: totalBl,
        totalBiaya: totalBiaya,
        labaBersih: labaBersih,
        saldoKas: saldoKas,
        totalModal: totalModal,
        unitAktif: allUnits.length,
        totalTransaksi: pend.length + bv.length + bl.length + modal.length,
        grandTotalTransaksi: totalSemuaTransaksi,
        tahunTerpilih: targetYear
      },
      trend: trend
    };
  } catch (e) {
    return {
      summary: {
        totalPendapatan: 0, totalBiaya: 0, labaBersih: 0, saldoKas: 0, totalModal: 0, unitAktif: 0, totalTransaksi: 0,
        grandTotalTransaksi: 0,
        error: e.message
      },
      trend: []
    };
  }
}

// ============================================================
// LAPORAN RESMI
// ============================================================
function getLaporanLabaRugi(tahun, bulan) {
  try {
    const isAll = (!tahun || tahun === 'all' || tahun === 'semua');
    const filter = isAll ? {} : { tahun: parseInt(tahun) };
    if (bulan && bulan !== 'all') filter.bulan = parseInt(bulan);

    const pendapatan    = getPendapatan(filter);
    const biayaVariabel = getBiayaVariabel(filter);
    const biayaLainnya  = getBiayaLainnya(filter);

    const totalPendapatan    = pendapatan.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBiayaVariabel = biayaVariabel.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBiayaLainnya  = biayaLainnya.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const labaKotor  = totalPendapatan - totalBiayaVariabel;
    const labaBersih = totalPendapatan - totalBiayaVariabel - totalBiayaLainnya;

    const kategoriPendapatan = {};
    pendapatan.forEach(function(r) {
      const k = r.Kategori || 'Lain-lain';
      kategoriPendapatan[k] = (kategoriPendapatan[k] || 0) + (parseNumber(r.Jumlah) || 0);
    });

    const kategoriBiayaVar = {};
    biayaVariabel.forEach(function(r) {
      const k = r.Kategori || 'Lain-lain';
      kategoriBiayaVar[k] = (kategoriBiayaVar[k] || 0) + (parseNumber(r.Jumlah) || 0);
    });

    const kategoriBiayaLain = {};
    biayaLainnya.forEach(function(r) {
      const k = r.Kategori || 'Lain-lain';
      kategoriBiayaLain[k] = (kategoriBiayaLain[k] || 0) + (parseNumber(r.Jumlah) || 0);
    });

    const unitUsaha = getUnitUsaha();
    const perUnit = unitUsaha.map(function(u) {
      const pU = pendapatan.filter(function(r) { return r.Unit_Usaha === u.Nama_Unit; });
      const bU = biayaVariabel.filter(function(r) { return r.Unit_Usaha === u.Nama_Unit; });
      const pTot = pU.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
      const bTot = bU.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
      return { unit: u.Nama_Unit, pendapatan: pTot, biaya: bTot, laba: pTot - bTot };
    });

    return {
      periode: { tahun: isAll ? 'Semua Tahun' : filter.tahun, bulan: filter.bulan || null },
      pendapatan: { total: totalPendapatan, detail: kategoriPendapatan },
      biayaVariabel: { total: totalBiayaVariabel, detail: kategoriBiayaVar },
      biayaLainnya:  { total: totalBiayaLainnya,  detail: kategoriBiayaLain },
      totalBiaya: totalBiayaVariabel + totalBiayaLainnya,
      labaKotor: labaKotor,
      labaBersih: labaBersih,
      perUnit: perUnit
    };
  } catch (e) {
    return {
      periode: { tahun: tahun, bulan: bulan },
      pendapatan: { total: 0, detail: {} },
      biayaVariabel: { total: 0, detail: {} },
      biayaLainnya: { total: 0, detail: {} },
      totalBiaya: 0, labaKotor: 0, labaBersih: 0, perUnit: []
    };
  }
}

function getLaporanArusKas(tahun, bulan) {
  try {
    const isAll = (!tahun || tahun === 'all' || tahun === 'semua');
    const allKas = getKas({});
    const filterKas = isAll ? allKas : getKas({ tahun: tahun, bulan: bulan || null });

    const operasi   = filterKas.filter(function(r) { return ['Pendapatan', 'Biaya', 'Biaya Variabel', 'Biaya Operasional'].indexOf(r.Jenis) >= 0; });
    const investasi  = filterKas.filter(function(r) { return r.Jenis === 'Modal'; });
    const finansiasi = filterKas.filter(function(r) { return r.Jenis === 'Alokasi PAD'; });

    const arusOperasi = operasi.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0) - (parseNumber(r.Keluar) || 0); }, 0);
    const arusInvestasi = investasi.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0) - (parseNumber(r.Keluar) || 0); }, 0);
    const arusFinansiasi = finansiasi.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0) - (parseNumber(r.Keluar) || 0); }, 0);

    const saldoAwal = 0;
    const saldoAkhir = allKas.length > 0 ? (parseNumber(allKas[allKas.length - 1].Saldo) || 0) : 0;

    const trendBulanan = [];
    for (let m = 1; m <= 12; m++) {
      const km = filterKas.filter(function(r) {
        const dp = r._dateParsed || parseDateInfo(r.Tanggal);
        return dp && dp.month === m;
      });
      const masuk = km.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0); }, 0);
      const keluar = km.reduce(function(s, r) { return s + (parseNumber(r.Keluar) || 0); }, 0);
      trendBulanan.push({ bulan: m, masuk: masuk, keluar: keluar, neto: masuk - keluar });
    }

    return {
      periode: { tahun: isAll ? 'Semua Tahun' : tahun, bulan: bulan || null },
      arusOperasi: arusOperasi,
      arusInvestasi: arusInvestasi,
      arusFinansiasi: arusFinansiasi,
      totalArus: arusOperasi + arusInvestasi + arusFinansiasi,
      saldoAwal: saldoAwal,
      saldoAkhir: saldoAkhir,
      trendBulanan: trendBulanan,
      detail: filterKas
    };
  } catch (e) {
    return {
      periode: { tahun: tahun, bulan: bulan },
      arusOperasi: 0, arusInvestasi: 0, arusFinansiasi: 0, totalArus: 0,
      saldoAwal: 0, saldoAkhir: 0, trendBulanan: [], detail: []
    };
  }
}

function getLaporanNeraca(tahun) {
  try {
    const isAll = (!tahun || tahun === 'all' || tahun === 'semua');
    const modalAll = isAll ? getModal({}) : getModal({ tahun: tahun });
    const pendAll  = isAll ? getPendapatan({}) : getPendapatan({ tahun: tahun });
    const bvAll    = isAll ? getBiayaVariabel({}) : getBiayaVariabel({ tahun: tahun });
    const blAll    = isAll ? getBiayaLainnya({}) : getBiayaLainnya({ tahun: tahun });
    const kasAll   = getKas({});

    const totalModal  = modalAll.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalPend   = pendAll.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const totalBiaya  = bvAll.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0) + blAll.reduce(function(s, r) { return s + (parseNumber(r.Jumlah) || 0); }, 0);
    const labaDitahan = totalPend - totalBiaya;

    let saldoKas = 0;
    if (kasAll.length > 0) {
      saldoKas = parseNumber(kasAll[kasAll.length - 1].Saldo);
      if (saldoKas === 0 && (allKas[allKas.length - 1].Masuk || allKas[allKas.length - 1].Keluar)) {
        const totM = kasAll.reduce(function(s, r) { return s + (parseNumber(r.Masuk) || 0); }, 0);
        const totK = kasAll.reduce(function(s, r) { return s + (parseNumber(r.Keluar) || 0); }, 0);
        saldoKas = totM - totK;
      }
    } else {
      saldoKas = totalModal + labaDitahan;
    }

    return {
      aktiva: {
        lancar: { kas: saldoKas, piutang: 0, persediaan: 0, total: saldoKas },
        tetap:  { peralatan: 0, akumulasi: 0, total: 0 },
        totalAktiva: saldoKas
      },
      pasiva: {
        kewajiban: { jangkaPendek: 0, jangkaPanjang: 0, total: 0 },
        modal: { penyertaan: totalModal, labaDitahan: labaDitahan, total: totalModal + labaDitahan },
        totalPasiva: totalModal + labaDitahan
      }
    };
  } catch (e) {
    return {
      aktiva: { lancar: { kas: 0, piutang: 0, persediaan: 0, total: 0 }, tetap: { peralatan: 0, akumulasi: 0, total: 0 }, totalAktiva: 0 },
      pasiva: { kewajiban: { jangkaPendek: 0, jangkaPanjang: 0, total: 0 }, modal: { penyertaan: 0, labaDitahan: 0, total: 0 }, totalPasiva: 0 }
    };
  }
}

function deleteRowById(sheetName, id) {
  try {
    const ss = getSS();
    const sheet = findSheet(ss, [sheetName]);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: false, message: 'Sheet kosong' };

    for (let i = 1; i < rows.length; i++) {
      for (let c = 0; c < rows[i].length; c++) {
        if (String(rows[i][c]).trim() === String(id).trim()) {
          sheet.deleteRow(i + 1);
          return { success: true };
        }
      }
    }

    const matchRow = String(id).match(/^([A-Z]+)-(\d+)$/i);
    if (matchRow) {
      const rowNum = parseInt(matchRow[2], 10);
      if (rowNum >= 2 && rowNum <= rows.length) {
        sheet.deleteRow(rowNum);
        return { success: true };
      }
    }

    return { success: false, message: 'Data tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
