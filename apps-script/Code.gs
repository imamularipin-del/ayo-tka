/**
 * ============================================================================
 *  CBT TKA - SD NEGERI SEDA  |  Tahun Ajaran 2025/2026  |  Kelas 6 Semester 2
 *  Code.gs  -  100% Google Apps Script + Google Sheets (tanpa database luar)
 *
 *  Cara pakai singkat:
 *   1. Jalankan fungsi setupSheets() satu kali dari editor (beri izin akses).
 *   2. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 *   3. Login admin default:  admin / admin123  (segera ganti di menu Pengaturan).
 *
 *  Catatan keamanan: kunci jawaban TIDAK PERNAH dikirim ke HTML siswa selama
 *  ujian. Penilaian dilakukan di server (fungsi finalisasi_/cekBenar_).
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// KONSTANTA
// ---------------------------------------------------------------------------
const TZ = 'Asia/Jakarta';
const MAPEL_LIST = ['Bahasa Indonesia', 'Matematika'];
const JENIS_LIST = ['PG', 'PGK', 'PG_Kategori'];
const FRIENDLY = 'Maaf, terjadi kendala di server. Silakan hubungi proktor.';

// Struktur 7 sheet (urutan kolom = urutan header)
const SHEETS = {
  CONFIG: ['Key', 'Value'],
  SOAL: ['No', 'ID_Soal', 'Mapel', 'Kelas', 'Jenis_Soal', 'Stimulus_ID', 'Stimulus_Judul', 'Stimulus_Teks',
    'Stimulus_Gambar_URL', 'Pertanyaan', 'Pilihan_A', 'Pilihan_B', 'Pilihan_C', 'Pilihan_D', 'Pernyataan_1',
    'Pernyataan_2', 'Pernyataan_3', 'Pernyataan_4', 'Kunci_Jawaban', 'Pembahasan', 'Level', 'Status'],
  SISWA_UJIAN: ['Timestamp', 'Token', 'Nama', 'NISN', 'Kelas', 'Mapel', 'Urutan_Soal_JSON', 'Waktu_Mulai',
    'Waktu_Selesai', 'Status'],
  // kolom NISN ditambahkan di akhir JAWABAN supaya jawaban tiap siswa unik
  JAWABAN: ['Token', 'Nama', 'Mapel', 'ID_Soal', 'Jenis_Soal', 'Jawaban_Siswa', 'Benar_Salah', 'Waktu_Jawab', 'NISN'],
  NILAI: ['Timestamp', 'Token', 'Nama', 'NISN', 'Kelas', 'Mapel', 'Benar_PG', 'Benar_PGK', 'Benar_Kategori',
    'Total_Benar', 'Total_Soal', 'Nilai_Akhir', 'Predikat', 'Durasi_Menit', 'Status'],
  TOKEN: ['Token', 'Mapel', 'Kelas', 'Tanggal', 'Jam_Mulai', 'Jam_Selesai', 'Kuota', 'Terpakai', 'Aktif', 'Dibuat_Oleh'],
  ADMIN: ['Username', 'Password_Hash', 'Nama', 'Role']
};

// Kolom (1-based) yang harus berformat teks agar angka 0 di depan tidak hilang
const TEXT_COLS = {
  SOAL: [2, 4, 19], SISWA_UJIAN: [2, 4, 5], JAWABAN: [1, 4, 9], NILAI: [2, 4, 5],
  TOKEN: [1, 3, 4, 5, 6], ADMIN: [1]
};

const DEFAULT_CONFIG = [
  ['nama_sekolah', 'SD NEGERI SEDA'], ['tahun_ajaran', '2025/2026'], ['durasi_menit', '75'],
  ['jumlah_soal', '30'], ['tampilkan_pembahasan', 'YA'], ['jumlah_pg', '15'], ['jumlah_pgk', '8'],
  ['jumlah_kategori', '7'], ['acak_soal', 'YA'], ['acak_pilihan', 'YA'], ['batas_pindah_tab', '3']
];

// ---------------------------------------------------------------------------
// HELPER UMUM
// ---------------------------------------------------------------------------
/** Format balasan standar semua fungsi server. */
function ok_(data, message) { return { success: true, data: data === undefined ? null : data, message: message || '' }; }
function fail_(message, data) { return { success: false, data: data === undefined ? null : data, message: message || FRIENDLY }; }
function str_(v) { return v === null || v === undefined ? '' : String(v); }

/** Jalankan fungsi dengan try-catch; error mentah TIDAK pernah dikirim ke layar. */
function safe_(fn) {
  try { return fn(); } catch (e) {
    console.error(e && e.stack ? e.stack : e);
    return fail_(FRIENDLY);
  }
}
/** Jalankan fungsi di dalam LockService (untuk semua fungsi yang menulis). */
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_('Server sedang ramai. Coba lagi sebentar ya.', { busy: true });
  try { return fn(); } finally { try { lock.releaseLock(); } catch (e) { } }
}
function run_(fn, write) { return safe_(() => (write ? locked_(fn) : fn())); }

/** Ambil Spreadsheet: ID disimpan di PropertiesService.
 *  Urutan: (1) SS_ID tersimpan -> (2) spreadsheet induk (script terikat) -> (3) buat spreadsheet baru otomatis
 *  (untuk script standalone, supaya web app tidak mati karena "Spreadsheet tidak ditemukan"). */
function getSS_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SS_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; } }
  if (!ss) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const again = p.getProperty('SS_ID');   // cek ulang: request lain mungkin sudah membuatnya
      if (again) { try { ss = SpreadsheetApp.openById(again); } catch (e) { ss = null; } }
      if (!ss) {
        ss = SpreadsheetApp.create('CBT TKA - SD NEGERI SEDA (Database)');
        console.log('Spreadsheet database dibuat: ' + ss.getUrl());
      }
    } finally { try { lock.releaseLock(); } catch (e) { } }
  }
  if (ss && ss.getId() !== id) p.setProperty('SS_ID', ss.getId());
  return ss;
}

/** Baca seluruh data sebuah sheet (tanpa header). */
function read_(name) {
  const sheet = getSS_().getSheetByName(name);
  if (!sheet) throw new Error('Sheet ' + name + ' belum ada. Jalankan setupSheets().');
  const lr = sheet.getLastRow();
  const data = lr > 1 ? sheet.getRange(2, 1, lr - 1, SHEETS[name].length).getValues() : [];
  return { sheet: sheet, data: data };
}

/** Tambah satu baris ke sheet dengan kolom teks terjaga formatnya. */
function appendRow_(name, arr) {
  const sheet = getSS_().getSheetByName(name);
  const r = sheet.getLastRow() + 1;
  (TEXT_COLS[name] || []).forEach(c => sheet.getRange(r, c).setNumberFormat('@'));
  sheet.getRange(r, 1, 1, arr.length).setValues([arr]);
  return r;
}
/** Tambah banyak baris sekaligus (cepat). */
function appendRows_(name, rows) {
  if (!rows.length) return;
  const sheet = getSS_().getSheetByName(name);
  const r = sheet.getLastRow() + 1;
  (TEXT_COLS[name] || []).forEach(c => sheet.getRange(r, c, rows.length, 1).setNumberFormat('@'));
  sheet.getRange(r, 1, rows.length, rows[0].length).setValues(rows);
}
/** Hapus baris yang memenuhi syarat (dari bawah ke atas). Return jumlah terhapus. */
function deleteRows_(name, pred) {
  const r = read_(name);
  let n = 0;
  for (let i = r.data.length - 1; i >= 0; i--) {
    if (pred(r.data[i])) { r.sheet.deleteRow(i + 2); n++; }
  }
  return n;
}

function getConfig_() {
  const c = {};
  DEFAULT_CONFIG.forEach(d => c[d[0]] = d[1]);
  read_('CONFIG').data.forEach(r => { if (str_(r[0]) !== '') c[str_(r[0]).trim()] = str_(r[1]).trim(); });
  return c;
}

function sha256_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}
function shuffle_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}
function toMs_(v) {
  if (v instanceof Date) return v.getTime();
  const t = new Date(str_(v)).getTime();
  return isNaN(t) ? 0 : t;
}
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = str_(v).trim();
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}
function fmtTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = str_(v).match(/(\d{1,2})[:.](\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function fmtTs_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm') : str_(v); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function predikat_(n) { return n >= 90 ? 'Sangat Baik' : n >= 75 ? 'Baik' : n >= 60 ? 'Cukup' : 'Perlu Bimbingan'; }

// Cache bertingkat (nilai > 100KB dipecah otomatis)
function cachePut_(key, obj, sec) {
  const c = CacheService.getScriptCache(), s = JSON.stringify(obj), size = 30000;
  const n = Math.max(1, Math.ceil(s.length / size)), m = {};
  m[key + '_n'] = String(n);
  for (let i = 0; i < n; i++) m[key + '_' + i] = s.substr(i * size, size);
  try { c.putAll(m, sec); } catch (e) { }
}
function cacheGet_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n'));
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const got = c.getAll(keys); let s = '';
  for (let i = 0; i < n; i++) { if (got[keys[i]] === undefined) return null; s += got[keys[i]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cacheDel_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n')) || 0, keys = [key + '_n'];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  c.removeAll(keys);
}

// ---------------------------------------------------------------------------
// SETUP SHEET OTOMATIS  (jalankan sekali dari editor Apps Script)
// ---------------------------------------------------------------------------
/** Membuat 7 sheet + header + CONFIG default + akun admin default (aman dijalankan berulang). */
function setupSheets() {
  return safe_(() => {
    const ss = getSS_();
    ss.setSpreadsheetTimeZone(TZ);
    Object.keys(SHEETS).forEach(name => {
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      const head = SHEETS[name];
      sh.getRange(1, 1, 1, head.length).setValues([head])
        .setFontWeight('bold').setBackground('#1e6fd9').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      (TEXT_COLS[name] || []).forEach(c => sh.getRange(2, c, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@'));
    });
    // CONFIG default (tambah key yang belum ada)
    const cfg = read_('CONFIG');
    const have = cfg.data.map(r => str_(r[0]).trim());
    const add = DEFAULT_CONFIG.filter(d => have.indexOf(d[0]) < 0);
    if (add.length) appendRows_('CONFIG', add);
    // Admin default
    if (!read_('ADMIN').data.length) appendRow_('ADMIN', ['admin', sha256_('admin123'), 'Administrator', 'Admin']);
    // Hapus sheet bawaan kosong
    ['Sheet1', 'Lembar1'].forEach(n => {
      const s = ss.getSheetByName(n);
      if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
    PropertiesService.getScriptProperties().setProperty('SETUP_OK', '1');
    SpreadsheetApp.flush();
    return ok_(null, 'Semua sheet siap digunakan. Login admin: admin / admin123');
  });
}
function ensureSetup_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SETUP_OK')) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (props.getProperty('SETUP_OK')) return;   // sudah disiapkan request lain
    const r = setupSheets();
    if (!r || !r.success) throw new Error('setupSheets gagal');
  } finally { try { lock.releaseLock(); } catch (e) { } }
}

/** Titik masuk web app. */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CBT TKA - SD NEGERI SEDA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    console.error(e && e.stack ? e.stack : e);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px"><h2>Aplikasi belum bisa dibuka</h2>' +
      '<p>File <b>Index.html</b> tidak ditemukan di project Apps Script ini. Pastikan namanya persis <b>Index</b> ' +
      '(huruf I besar), lalu deploy ulang (Deploy &gt; Manage deployments &gt; Edit &gt; New version).</p></div>');
  }
}

/** Daftar fungsi yang boleh dipanggil dari luar Apps Script (mode Cloudflare Worker).
 *  Hanya fungsi di daftar ini yang bisa dijalankan; fungsi lain (setupSheets, diagnosa, dst.) tidak terbuka. */
const API_MAP = {
  getPublicInfo: getPublicInfo, cekToken: cekToken, mulaiUjian: mulaiUjian, getSoalUntukSiswa: getSoalUntukSiswa,
  simpanJawaban: simpanJawaban, submitUjian: submitUjian, getHasil: getHasil, getGambarSiswa: getGambarSiswa,
  loginAdmin: loginAdmin, logoutAdmin: logoutAdmin, getDashboard: getDashboard, getRekap: getRekap, getAnalisis: getAnalisis,
  hapusNilai: hapusNilai, resetSiswa: resetSiswa, kelolaPengaturan: kelolaPengaturan, kelolaSoal: kelolaSoal,
  kelolaToken: kelolaToken, exportCSV: exportCSV, uploadGambar: uploadGambar, getGambarAdmin: getGambarAdmin
};

/** Titik masuk API untuk Cloudflare Worker: body teks JSON {"fn":"nama","args":[...]} -> balasan JSON. */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(req.fn || ''), args = Array.isArray(req.args) ? req.args : [];
    out = Object.prototype.hasOwnProperty.call(API_MAP, fn) ? API_MAP[fn].apply(null, args) : fail_('Fungsi tidak dikenal.');
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    out = fail_(FRIENDLY);
  }
  return ContentService.createTextOutput(JSON.stringify(out === undefined ? null : out)).setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan SEKALI dari editor untuk memeriksa kondisi database; hasilnya ada di Execution log. */
function diagnosa() {
  const out = [];
  try {
    const ss = getSS_();
    out.push('Spreadsheet OK: ' + ss.getName() + '  ->  ' + ss.getUrl());
    const have = ss.getSheets().map(x => x.getName());
    Object.keys(SHEETS).forEach(n => out.push((have.indexOf(n) >= 0 ? 'OK     ' : 'HILANG ') + 'sheet ' + n));
    out.push('SETUP_OK = ' + PropertiesService.getScriptProperties().getProperty('SETUP_OK'));
    out.push('Web app URL = ' + (ScriptApp.getService().getUrl() || '(belum di-deploy)'));
  } catch (e) { out.push('GAGAL: ' + e.message); }
  console.log(out.join('\n'));
  return out.join('\n');
}

/** Info publik (tanpa rahasia) untuk layar login. */
function getPublicInfo() {
  return run_(() => {
    let url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
    let c = {}, siap = true;
    DEFAULT_CONFIG.forEach(d => c[d[0]] = d[1]);
    try { ensureSetup_(); c = getConfig_(); }
    catch (e) { siap = false; console.error('getPublicInfo: database belum siap -> ' + (e && e.stack ? e.stack : e)); }
    return ok_({
      siap: siap, url: url, sekolah: c.nama_sekolah, tahun: c.tahun_ajaran, durasi: Number(c.durasi_menit) || 75,
      jumlah: Number(c.jumlah_soal) || 30, pg: Number(c.jumlah_pg) || 15, pgk: Number(c.jumlah_pgk) || 8,
      kat: Number(c.jumlah_kategori) || 7, batasTab: Number(c.batas_pindah_tab) || 3, mapel: MAPEL_LIST
    });
  });
}

// ===========================================================================
//  BANK SOAL (dengan cache 5 menit supaya 30 siswa barengan tidak lemot)
// ===========================================================================
/** Membaca semua soal dari sheet SOAL (hasil di-cache 5 menit). Hanya dipakai di server. */
function loadSoal_() {
  let bank = cacheGet_('soal_all');
  if (bank) return bank;
  const H = SHEETS.SOAL;
  bank = read_('SOAL').data.filter(v => str_(v[1]) !== '').map(v => {
    const o = {}; H.forEach((h, i) => o[h] = str_(v[i]).trim ? str_(v[i]).trim() : str_(v[i])); return o;
  });
  cachePut_('soal_all', bank, 300);
  return bank;
}
function clearSoalCache_() { cacheDel_('soal_all'); }
function mapById_(arr) { const m = {}; arr.forEach(s => m[s.ID_Soal] = s); return m; }

// ===========================================================================
//  TOKEN
// ===========================================================================
function isAktif_(v) { return ['YA', 'TRUE', 'Y', '1', 'AKTIF'].indexOf(str_(v).trim().toUpperCase()) >= 0; }

/** Daftar token sebagai objek. */
function tokenList_() {
  return read_('TOKEN').data.map((v, i) => ({
    row: i + 2, token: str_(v[0]).trim().toUpperCase(), mapel: str_(v[1]).trim(), kelas: str_(v[2]).trim(),
    tanggal: fmtDate_(v[3]), mulai: fmtTime_(v[4]), selesai: fmtTime_(v[5]), kuota: Number(v[6]) || 0,
    terpakai: Number(v[7]) || 0, aktif: isAktif_(v[8]), dibuat: str_(v[9])
  })).filter(t => t.token);
}
function cariToken_(token) { return tokenList_().filter(t => t.token === token)[0] || null; }

/** Validasi token: aktif, mapel/kelas cocok, tanggal & jam berlaku. Return pesan error atau ''. */
function validasiToken_(tk, mapel, kelas) {
  if (!tk.aktif) return 'Token belum aktif atau sudah dinonaktifkan. Tanyakan ke proktor.';
  if (tk.mapel && tk.mapel.toLowerCase() !== 'semua' && tk.mapel !== mapel) return 'Token ini untuk mapel ' + tk.mapel + '. Pilih mapel yang sesuai.';
  if (tk.kelas && tk.kelas.toLowerCase() !== 'semua' && tk.kelas !== String(kelas)) return 'Token ini untuk kelas ' + tk.kelas + '.';
  if (tk.tanggal && tk.tanggal !== today_()) return 'Token ini berlaku pada tanggal ' + tk.tanggal + '.';
  const now = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  if (tk.mulai && now < tk.mulai) return 'Ujian belum dimulai. Jam mulai ' + tk.mulai + '.';
  if (tk.selesai && now > tk.selesai) return 'Waktu token sudah berakhir (jam ' + tk.selesai + ').';
  return '';
}

function cekInput_(d) {
  d = d || {};
  const o = {
    nama: str_(d.nama).replace(/\s+/g, ' ').trim(), nisn: str_(d.nisn).trim(), kelas: str_(d.kelas || '6').trim(),
    mapel: str_(d.mapel).trim(), token: str_(d.token).trim().toUpperCase()
  };
  if (o.nama.length < 3) o.err = 'Nama lengkap belum diisi.';
  else if (!/^\d{10}$/.test(o.nisn)) o.err = 'NISN harus 10 angka.';
  else if (!/^[1-6]$/.test(o.kelas)) o.err = 'Pilih kelas 1 sampai 6.';
  else if (MAPEL_LIST.indexOf(o.mapel) < 0) o.err = 'Pilih mata pelajaran dulu ya.';
  else if (!o.token) o.err = 'Token belum diisi.';
  return o;
}

/** Cari sesi ujian sebuah NISN pada mapel tertentu. */
function cariSesiNisn_(nisn, mapel) {
  const r = read_('SISWA_UJIAN'); let jalan = null, selesai = false;
  r.data.forEach(v => {
    if (str_(v[3]) === nisn && str_(v[5]) === mapel) { if (str_(v[9]) === 'Berjalan') jalan = v; else selesai = true; }
  });
  return { jalan: jalan, selesai: selesai };
}

/** Validasi data login siswa + token (belum memulai ujian). */
function cekToken(d) {
  return run_(() => {
    ensureSetup_();
    const v = cekInput_(d); if (v.err) return fail_(v.err);
    const tk = cariToken_(v.token);
    if (!tk) return fail_('Token tidak ditemukan. Periksa lagi huruf dan angkanya ya.');
    const ses = cariSesiNisn_(v.nisn, v.mapel);
    if (ses.selesai) return fail_('NISN ini sudah mengerjakan ujian ' + v.mapel + '. Hubungi proktor jika perlu ujian ulang.');
    let resume = false;
    if (ses.jalan) {
      if (str_(ses.jalan[1]).toUpperCase() !== v.token) return fail_('Kamu masih punya ujian yang berjalan. Pakai token yang sama seperti tadi.');
      resume = true;
    } else {
      const e = validasiToken_(tk, v.mapel, v.kelas); if (e) return fail_(e);
      if (tk.terpakai >= tk.kuota) return fail_('Kuota token sudah habis. Minta token baru ke proktor.');
    }
    const c = getConfig_();
    return ok_({
      nama: v.nama, nisn: v.nisn, kelas: v.kelas, mapel: v.mapel, token: v.token, resume: resume,
      durasi: Number(c.durasi_menit) || 75, jumlah: Number(c.jumlah_soal) || 30,
      pg: Number(c.jumlah_pg) || 15, pgk: Number(c.jumlah_pgk) || 8, kat: Number(c.jumlah_kategori) || 7
    }, 'Token valid');
  });
}

// ===========================================================================
//  MEMBUAT PAKET SOAL PER SISWA (acak soal + acak pilihan, disimpan permanen)
// ===========================================================================
/** Pilih soal sesuai kuota jenis; soal grup (satu stimulus) selalu utuh & berurutan. */
function buatPaket_(mapel, kelas, cfg) {
  const bank = loadSoal_().filter(s => s.Mapel === mapel && s.Status.toLowerCase() === 'aktif' &&
    (s.Kelas === '' || s.Kelas === String(kelas)) && JENIS_LIST.indexOf(s.Jenis_Soal) >= 0);
  if (!bank.length) return null;
  const total = Number(cfg.jumlah_soal) || 30;
  const quota = { PG: Number(cfg.jumlah_pg) || 0, PGK: Number(cfg.jumlah_pgk) || 0, PG_Kategori: Number(cfg.jumlah_kategori) || 0 };
  const groups = {}, units = [];
  bank.forEach(s => {
    if (s.Stimulus_ID) { if (!groups[s.Stimulus_ID]) { groups[s.Stimulus_ID] = []; units.push(groups[s.Stimulus_ID]); } groups[s.Stimulus_ID].push(s); }
    else units.push([s]);
  });
  const acakSoal = str_(cfg.acak_soal).toUpperCase() !== 'TIDAK';
  const acakPil = str_(cfg.acak_pilihan).toUpperCase() !== 'TIDAK';
  const order = acakSoal ? shuffle_(units) : units;
  const pick = [], cnt = { PG: 0, PGK: 0, PG_Kategori: 0 }, used = [];
  let n = 0;
  const tambah = (u, ketat) => {
    if (n + u.length > total) return;
    if (ketat) {
      const need = {}; u.forEach(s => need[s.Jenis_Soal] = (need[s.Jenis_Soal] || 0) + 1);
      for (const k in need) if (cnt[k] + need[k] > quota[k]) return;
    }
    u.forEach(s => cnt[s.Jenis_Soal]++); pick.push(u); used.push(u); n += u.length;
  };
  order.forEach(u => tambah(u, true));                          // putaran 1: ikuti kuota jenis
  order.forEach(u => { if (used.indexOf(u) < 0) tambah(u, false); }); // putaran 2: isi sisa slot
  const finalUnits = acakSoal ? shuffle_(pick) : pick;
  const ids = [], opsi = {};
  finalUnits.forEach(u => u.forEach(s => {
    ids.push(s.ID_Soal);
    if (s.Jenis_Soal === 'PG' || s.Jenis_Soal === 'PGK') {
      const L = ['A', 'B', 'C', 'D'].filter(l => s['Pilihan_' + l] !== '');
      opsi[s.ID_Soal] = acakPil ? shuffle_(L) : L;
    }
  }));
  return { ids: ids, opsi: opsi, dur: Number(cfg.durasi_menit) || 75 };
}

/** Mulai ujian: membuat baris SISWA_UJIAN + mengurangi kuota token. Aman dipanggil ulang (lanjut ujian). */
function mulaiUjian(d) {
  return run_(() => {
    const v = cekInput_(d); if (v.err) return fail_(v.err);
    const tk = cariToken_(v.token);
    if (!tk) return fail_('Token tidak ditemukan.');
    const ses = cariSesiNisn_(v.nisn, v.mapel);
    if (ses.selesai) return fail_('NISN ini sudah mengerjakan ujian ' + v.mapel + '.');
    if (ses.jalan) {
      if (str_(ses.jalan[1]).toUpperCase() !== v.token) return fail_('Pakai token yang sama seperti sebelumnya.');
      let u = {}; try { u = JSON.parse(ses.jalan[6]); } catch (e) { }
      return ok_({ token: v.token, nisn: v.nisn, mapel: v.mapel, nama: str_(ses.jalan[2]), sk: u.sk }, 'Melanjutkan ujian');
    }
    const e = validasiToken_(tk, v.mapel, v.kelas); if (e) return fail_(e);
    if (tk.terpakai >= tk.kuota) return fail_('Kuota token sudah habis.');
    const paket = buatPaket_(v.mapel, v.kelas, getConfig_());
    if (!paket) return fail_('Soal ' + v.mapel + ' untuk kelas ' + v.kelas + ' belum tersedia. Hubungi proktor.');
    paket.sk = Utilities.getUuid().replace(/-/g, '').substr(0, 16);
    const now = new Date();
    appendRow_('SISWA_UJIAN', [now, v.token, v.nama, v.nisn, v.kelas, v.mapel, JSON.stringify(paket), now, '', 'Berjalan']);
    getSS_().getSheetByName('TOKEN').getRange(tk.row, 8).setValue(tk.terpakai + 1);
    SpreadsheetApp.flush();
    return ok_({ token: v.token, nisn: v.nisn, mapel: v.mapel, nama: v.nama, sk: paket.sk }, 'Ujian dimulai');
  }, true);
}

// ===========================================================================
//  SESI SISWA
// ===========================================================================
/** Cari sesi (baca penuh dari sheet). Return {idx,row,urut,sheet} atau null. */
function sesi_(sess) {
  if (!sess || !sess.nisn || !sess.token || !sess.sk) return null;
  const r = read_('SISWA_UJIAN');
  for (let i = 0; i < r.data.length; i++) {
    const v = r.data[i];
    if (str_(v[3]) === str_(sess.nisn) && str_(v[1]).toUpperCase() === str_(sess.token).toUpperCase() && str_(v[5]) === str_(sess.mapel)) {
      let u = {}; try { u = JSON.parse(v[6]); } catch (e) { }
      if (u.sk === sess.sk) return { idx: i + 2, row: v, urut: u, sheet: r.sheet };
    }
  }
  return null;
}
/** Versi ringan (cache 6 jam) untuk autosave. */
function sesiCache_(sess) {
  if (!sess || !sess.sk) return null;
  const c = CacheService.getScriptCache(), raw = c.get('ses_' + sess.sk);
  if (raw) {
    try {
      const o = JSON.parse(raw);
      if (o.nisn === str_(sess.nisn) && o.token === str_(sess.token).toUpperCase() && o.mapel === str_(sess.mapel)) return o;
    } catch (e) { }
  }
  const su = sesi_(sess); if (!su) return null;
  const o = {
    token: str_(su.row[1]).toUpperCase(), nama: str_(su.row[2]), nisn: str_(su.row[3]), mapel: str_(su.row[5]),
    ids: su.urut.ids || [], mulai: toMs_(su.row[7]), dur: su.urut.dur || 75, status: str_(su.row[9])
  };
  c.put('ses_' + sess.sk, JSON.stringify(o), 21600);
  return o;
}
function sisaDetik_(mulaiMs, durMenit) { return durMenit * 60 - (Date.now() - mulaiMs) / 1000; }

/** Kirim soal + jawaban tersimpan ke siswa. TIDAK menyertakan kunci jawaban & pembahasan. */
function getSoalUntukSiswa(sess) {
  return run_(() => {
    const su = sesi_(sess);
    if (!su) return fail_('Sesi ujian tidak ditemukan. Silakan login ulang.', { reset: true });
    if (str_(su.row[9]) !== 'Berjalan') return ok_({ selesai: true });
    const bank = mapById_(loadSoal_()), stim = {}, soal = [];
    su.urut.ids.forEach(id => {
      const s = bank[id]; if (!s) return;
      if (s.Stimulus_ID && !stim[s.Stimulus_ID]) stim[s.Stimulus_ID] = { judul: s.Stimulus_Judul, teks: s.Stimulus_Teks, gambar: s.Stimulus_Gambar_URL };
      const item = { id: id, jenis: s.Jenis_Soal, stim: s.Stimulus_ID, tanya: s.Pertanyaan };
      if (s.Jenis_Soal === 'PGK' || s.Jenis_Soal === 'PG') {
        item.opsi = (su.urut.opsi[id] || ['A', 'B', 'C', 'D']).map((L, i) => ({ label: 'ABCD'[i], kode: L, teks: s['Pilihan_' + L] }));
      } else {
        item.pernyataan = [1, 2, 3, 4].map(n => s['Pernyataan_' + n]).filter(x => x !== '');
        item.lab = [s.Pilihan_A || 'Benar', s.Pilihan_B || 'Salah'];
      }
      soal.push(item);
    });
    const jr = read_('JAWABAN'), jaw = {}, token = str_(su.row[1]), nisn = str_(su.row[3]), mapel = str_(su.row[5]);
    jr.data.forEach(v => { if (str_(v[8]) === nisn && str_(v[0]) === token && str_(v[2]) === mapel && str_(v[5]) !== '') jaw[str_(v[3])] = str_(v[5]); });
    const dur = su.urut.dur || 75;
    return ok_({
      nama: str_(su.row[2]), mapel: mapel, durasi: dur, soal: soal, stim: stim, jawaban: jaw,
      sisa: Math.max(0, Math.round(sisaDetik_(toMs_(su.row[7]), dur)))
    });
  });
}

function normSet_(s) {
  const a = str_(s).split(',').map(x => x.trim().toUpperCase()).filter(x => x);
  return a.filter((x, i) => a.indexOf(x) === i).sort();
}
function normJawab_(jenis, j) {
  j = str_(j).trim().substr(0, 200);
  if (jenis === 'PG') return j.toUpperCase().substr(0, 1);
  if (jenis === 'PGK') return normSet_(j).join(',');
  return j;
}

/** Cari nomor baris jawaban sebuah soal milik siswa (dengan cache nomor baris). */
function cariBarisJawaban_(sh, o, idSoal, sk) {
  const c = CacheService.getScriptCache();
  let map = {}; try { map = JSON.parse(c.get('jr_' + sk) || '{}'); } catch (e) { }
  const cocok = row => {
    if (!row || row > sh.getLastRow()) return false;
    const v = sh.getRange(row, 1, 1, 9).getValues()[0];
    return str_(v[0]) === o.token && str_(v[3]) === idSoal && str_(v[8]) === o.nisn;
  };
  if (map[idSoal] && cocok(map[idSoal])) return map[idSoal];
  map = {};
  const lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, 9).getValues().forEach((v, i) => {
    if (str_(v[8]) === o.nisn && str_(v[0]) === o.token && str_(v[2]) === o.mapel) map[str_(v[3])] = i + 2;
  });
  c.put('jr_' + sk, JSON.stringify(map), 21600);
  return map[idSoal] || 0;
}

/** AUTOSAVE: simpan satu jawaban ke sheet JAWABAN. Mengembalikan sisa waktu dari server. */
function simpanJawaban(sess, idSoal, jawaban) {
  return run_(() => {
    const o = sesiCache_(sess);
    if (!o) return fail_('Sesi ujian tidak ditemukan. Hubungi proktor.');
    if (o.status !== 'Berjalan') return fail_('Ujian sudah dikumpulkan.', { selesai: true });
    const sisa = sisaDetik_(o.mulai, o.dur);
    if (sisa < -30) return fail_('Waktu ujian sudah habis.', { habis: true });
    if (o.ids.indexOf(idSoal) < 0) return fail_('Soal tidak dikenali.');
    const s = mapById_(loadSoal_())[idSoal];
    const jenis = s ? s.Jenis_Soal : '';
    const ans = normJawab_(jenis, jawaban);
    const sh = getSS_().getSheetByName('JAWABAN');
    const row = cariBarisJawaban_(sh, o, idSoal, sess.sk);
    if (row) {
      sh.getRange(row, 6).setValue(ans);
      sh.getRange(row, 8).setValue(new Date());
    } else {
      const r = appendRow_('JAWABAN', [o.token, o.nama, o.mapel, idSoal, jenis, ans, '', new Date(), o.nisn]);
      const c = CacheService.getScriptCache();
      let map = {}; try { map = JSON.parse(c.get('jr_' + sess.sk) || '{}'); } catch (e) { }
      map[idSoal] = r; c.put('jr_' + sess.sk, JSON.stringify(map), 21600);
    }
    SpreadsheetApp.flush();
    return ok_({ sisa: Math.max(0, Math.round(sisa)) });
  }, true);
}

// ===========================================================================
//  PENILAIAN (HANYA DI SERVER)
// ===========================================================================
/** True jika jawaban benar. PG: sama. PGK: himpunan SAMA PERSIS. Kategori: SEMUA baris benar. */
function cekBenar_(s, ans) {
  ans = str_(ans).trim();
  if (!ans) return false;
  const k = str_(s.Kunci_Jawaban).trim();
  if (s.Jenis_Soal === 'PG') return ans.toUpperCase() === k.toUpperCase();
  if (s.Jenis_Soal === 'PGK') {
    const a = normSet_(ans), b = normSet_(k);
    return a.length === b.length && a.length > 0 && a.every((x, i) => x === b[i]);
  }
  const a = ans.split(',').map(x => x.trim().toLowerCase()), b = k.split(',').map(x => x.trim().toLowerCase());
  return a.length === b.length && a.every((x, i) => x !== '' && x === b[i]);
}

/** Hitung nilai, tulis Benar_Salah di JAWABAN, tulis NILAI, kunci sesi. Dipanggil di dalam lock. */
function finalisasi_(su, status) {
  const bank = mapById_(loadSoal_());
  const nama = str_(su.row[2]), nisn = str_(su.row[3]), token = str_(su.row[1]), mapel = str_(su.row[5]), kelas = str_(su.row[4]);
  const sh = getSS_().getSheetByName('JAWABAN'), lr = sh.getLastRow();
  const jv = lr > 1 ? sh.getRange(2, 1, lr - 1, 9).getValues() : [];
  const pos = {};
  jv.forEach((v, i) => { if (str_(v[8]) === nisn && str_(v[0]) === token && str_(v[2]) === mapel) pos[str_(v[3])] = i; });
  const cnt = { PG: 0, PGK: 0, PG_Kategori: 0 }, baru = [];
  let berubah = false;
  su.urut.ids.forEach(id => {
    const s = bank[id], jenis = s ? s.Jenis_Soal : '', i = pos[id];
    const ans = i !== undefined ? str_(jv[i][5]) : '';
    const b = s ? cekBenar_(s, ans) : false;
    if (b) cnt[jenis]++;
    if (i !== undefined) { jv[i][6] = b ? 'Benar' : 'Salah'; berubah = true; }
    else baru.push([token, nama, mapel, id, jenis, '', 'Salah', new Date(), nisn]);
  });
  if (berubah) sh.getRange(2, 7, jv.length, 1).setValues(jv.map(v => [v[6]]));
  appendRows_('JAWABAN', baru);
  const total = su.urut.ids.length, benar = cnt.PG + cnt.PGK + cnt.PG_Kategori;
  const nilai = total ? Math.round(benar / total * 1000) / 10 : 0;
  const mulai = toMs_(su.row[7]), dur = (su.urut.dur || 75) * 60000;
  const durMenit = Math.round((Math.min(Date.now(), mulai + dur) - mulai) / 6000) / 10;
  appendRow_('NILAI', [new Date(), token, nama, nisn, kelas, mapel, cnt.PG, cnt.PGK, cnt.PG_Kategori, benar, total, nilai,
    predikat_(nilai), durMenit, status]);
  su.sheet.getRange(su.idx, 9).setValue(new Date());
  su.sheet.getRange(su.idx, 10).setValue(status === 'Curang' ? 'Curang' : 'Selesai');
  CacheService.getScriptCache().removeAll(['ses_' + su.urut.sk, 'jr_' + su.urut.sk]);
  SpreadsheetApp.flush();
}

/** Detail pembahasan satu soal dalam urutan huruf yang dilihat siswa. */
function detailSoal_(s, ans, urutan, benar, no) {
  const o = urutan || ['A', 'B', 'C', 'D'];
  const lab = L => { const k = o.indexOf(L); return k < 0 ? L : 'ABCD'[k]; };
  const txt = L => lab(L) + '. ' + (s['Pilihan_' + L] || '');
  const d = { no: no, jenis: s.Jenis_Soal, tanya: s.Pertanyaan, benar: benar, pembahasan: s.Pembahasan };
  if (s.Jenis_Soal === 'PG') {
    d.jawab = ans ? txt(ans) : '(tidak dijawab)'; d.kunci = txt(str_(s.Kunci_Jawaban).trim().toUpperCase());
  } else if (s.Jenis_Soal === 'PGK') {
    const a = normSet_(ans), k = normSet_(s.Kunci_Jawaban);
    d.jawab = a.length ? a.map(txt).join(' | ') : '(tidak dijawab)'; d.kunci = k.map(txt).join(' | ');
  } else {
    const a = str_(ans).split(','), k = str_(s.Kunci_Jawaban).split(',');
    d.items = [1, 2, 3, 4].map(n => s['Pernyataan_' + n]).filter(x => x !== '')
      .map((p, i) => ({ p: p, jawab: (a[i] || '').trim() || '-', kunci: (k[i] || '').trim() }));
  }
  return d;
}

/** Susun hasil lengkap untuk layar siswa (membaca NILAI + JAWABAN). */
function hasilSiswa_(su) {
  const cfg = getConfig_(), bank = mapById_(loadSoal_());
  const token = str_(su.row[1]), nisn = str_(su.row[3]), mapel = str_(su.row[5]);
  const nv = read_('NILAI').data.filter(v => str_(v[1]) === token && str_(v[3]) === nisn && str_(v[5]) === mapel).pop();
  if (!nv) return null;
  const tot = { PG: 0, PGK: 0, PG_Kategori: 0 };
  su.urut.ids.forEach(id => { const s = bank[id]; if (s) tot[s.Jenis_Soal]++; });
  const h = {
    nama: str_(nv[2]), mapel: mapel, kelas: str_(nv[4]), nilai: Number(nv[11]), predikat: str_(nv[12]),
    benar: Number(nv[9]), total: Number(nv[10]), durasi: Number(nv[13]), status: str_(nv[14]),
    rincian: { pg: { b: Number(nv[6]), t: tot.PG }, pgk: { b: Number(nv[7]), t: tot.PGK }, kat: { b: Number(nv[8]), t: tot.PG_Kategori } },
    pembahasan: null
  };
  h.salah = h.total - h.benar;
  if (str_(cfg.tampilkan_pembahasan).toUpperCase() === 'YA') {
    const jm = {};
    read_('JAWABAN').data.forEach(v => { if (str_(v[8]) === nisn && str_(v[0]) === token && str_(v[2]) === mapel) jm[str_(v[3])] = { a: str_(v[5]), b: str_(v[6]) === 'Benar' }; });
    h.pembahasan = [];
    su.urut.ids.forEach((id, i) => {
      const s = bank[id]; if (!s) return;
      const j = jm[id] || { a: '', b: false };
      h.pembahasan.push(detailSoal_(s, j.a, su.urut.opsi[id], j.b, i + 1));
    });
  }
  return h;
}

/** Kumpulkan ujian (tombol Selesai / waktu habis / terindikasi curang). */
function submitUjian(sess, alasan) {
  return run_(() => {
    const su = sesi_(sess);
    if (!su) return fail_('Sesi ujian tidak ditemukan. Hubungi proktor.');
    if (str_(su.row[9]) === 'Berjalan') finalisasi_(su, alasan === 'curang' ? 'Curang' : 'Normal');
    const h = hasilSiswa_(su);
    return h ? ok_(h, 'Ujian dikumpulkan') : fail_();
  }, true);
}

/** Ambil ulang hasil (misalnya setelah halaman di-refresh). */
function getHasil(sess) {
  return run_(() => {
    const su = sesi_(sess);
    if (!su) return fail_('Hasil tidak ditemukan.');
    if (str_(su.row[9]) === 'Berjalan') return fail_('Ujian belum dikumpulkan.');
    const h = hasilSiswa_(su);
    return h ? ok_(h) : fail_('Hasil tidak ditemukan.');
  });
}

// ===========================================================================
//  ADMIN: LOGIN & SESI
// ===========================================================================
/** Login admin: password di-hash SHA256 di server. Salah 3x = terkunci 5 menit. */
function loginAdmin(username, password) {
  return run_(() => {
    ensureSetup_();
    const u = str_(username).trim().toLowerCase().substr(0, 40);
    const c = CacheService.getScriptCache();
    const lockUntil = Number(c.get('lk_' + u) || 0);
    if (lockUntil > Date.now()) return fail_('Terlalu banyak salah. Coba lagi ' + Math.ceil((lockUntil - Date.now()) / 60000) + ' menit lagi.');
    const hash = sha256_(str_(password));
    const a = read_('ADMIN').data.filter(r => str_(r[0]).trim().toLowerCase() === u && str_(r[1]) === hash)[0];
    if (!a) {
      const n = Number(c.get('fc_' + u) || 0) + 1;
      if (n >= 3) { c.put('lk_' + u, String(Date.now() + 300000), 300); c.remove('fc_' + u); return fail_('Salah 3 kali. Akun dikunci selama 5 menit.'); }
      c.put('fc_' + u, String(n), 600);
      return fail_('Username atau password salah. Sisa percobaan: ' + (3 - n));
    }
    c.remove('fc_' + u);
    const token = Utilities.getUuid();
    c.put('adm_' + token, JSON.stringify({ u: str_(a[0]), nama: str_(a[2]), role: str_(a[3]) }), 21600);
    return ok_({ token: token, nama: str_(a[2]), role: str_(a[3]) }, 'Selamat datang');
  });
}
function logoutAdmin(t) { CacheService.getScriptCache().remove('adm_' + str_(t)); return ok_(null); }

/** Bungkus fungsi admin: cek sesi dulu. */
function admin_(t, fn, write) {
  return run_(() => {
    const c = CacheService.getScriptCache(), raw = c.get('adm_' + str_(t));
    if (!raw) return fail_('Sesi habis, silakan login ulang.', { auth: true });
    c.put('adm_' + str_(t), raw, 21600);
    return fn(JSON.parse(raw));
  }, write);
}

// ===========================================================================
//  ADMIN: DASHBOARD
// ===========================================================================
/** Otomatis mengumpulkan ujian yang waktunya sudah lewat (siswa menutup HP / mati lampu). */
function sweeper_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const r = read_('SISWA_UJIAN');
    r.data.forEach((v, i) => {
      if (str_(v[9]) !== 'Berjalan') return;
      let u = {}; try { u = JSON.parse(v[6]); } catch (e) { }
      if (sisaDetik_(toMs_(v[7]), u.dur || 75) < -60) finalisasi_({ idx: i + 2, row: v, urut: u, sheet: r.sheet }, 'Normal');
    });
  } finally { try { lock.releaseLock(); } catch (e) { } }
}

/** Data dashboard: kartu, grafik, tabel live. */
function getDashboard(t) {
  return admin_(t, () => {
    sweeper_();
    const hari = today_(), su = read_('SISWA_UJIAN').data, nl = read_('NILAI').data, jw = read_('JAWABAN').data;
    const hariIni = su.filter(v => fmtDate_(v[0]) === hari);
    const nilaiHari = nl.filter(v => fmtDate_(v[0]) === hari).map(v => Number(v[11]) || 0);
    const avg = a => a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length * 10) / 10 : null;
    const perMapel = MAPEL_LIST.map(m => ({ mapel: m, rata: avg(nl.filter(v => str_(v[5]) === m).map(v => Number(v[11]) || 0)), n: nl.filter(v => str_(v[5]) === m).length }));
    const jn = { PG: [0, 0], PGK: [0, 0], PG_Kategori: [0, 0] };
    jw.forEach(v => { const k = str_(v[4]), b = str_(v[6]); if (jn[k] && b) { jn[k][1]++; if (b === 'Benar') jn[k][0]++; } });
    const perJenis = Object.keys(jn).map(k => ({ jenis: k, persen: jn[k][1] ? Math.round(jn[k][0] / jn[k][1] * 1000) / 10 : null, n: jn[k][1] }));
    const bank = mapById_(loadSoal_()), live = [];
    su.forEach(v => {
      if (str_(v[9]) !== 'Berjalan') return;
      let u = {}; try { u = JSON.parse(v[6]); } catch (e) { }
      const nisn = str_(v[3]), tok = str_(v[1]), mp = str_(v[5]);
      const terjawab = jw.filter(x => str_(x[8]) === nisn && str_(x[0]) === tok && str_(x[2]) === mp && str_(x[5]) !== '').length;
      live.push({
        nama: str_(v[2]), nisn: nisn, kelas: str_(v[4]), mapel: mp, terjawab: terjawab, total: (u.ids || []).length,
        sisa: Math.max(0, Math.round(sisaDetik_(toMs_(v[7]), u.dur || 75)))
      });
    });
    return ok_({
      peserta: hariIni.length, sedang: su.filter(v => str_(v[9]) === 'Berjalan').length,
      selesai: hariIni.filter(v => str_(v[9]) !== 'Berjalan').length, rata: avg(nilaiHari),
      perMapel: perMapel, perJenis: perJenis, live: live, waktu: Utilities.formatDate(new Date(), TZ, 'HH:mm:ss')
    });
  });
}

// ===========================================================================
//  ADMIN: KELOLA SOAL
// ===========================================================================
function soalRowToObj_(v) { const o = {}; SHEETS.SOAL.forEach((h, i) => o[h] = str_(v[i])); return o; }
function nextNo_() { return read_('SOAL').data.filter(v => str_(v[1]) !== '').length + 1; }
function newId_(prefix) { return (prefix || 'S') + '-' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 46656).toString(36).toUpperCase(); }

/** Validasi satu soal. Return pesan error atau ''. */
function validasiSoal_(o) {
  if (MAPEL_LIST.indexOf(o.Mapel) < 0) return 'Mapel harus Bahasa Indonesia atau Matematika.';
  if (JENIS_LIST.indexOf(o.Jenis_Soal) < 0) return 'Jenis soal harus PG, PGK, atau PG_Kategori.';
  if (!str_(o.Pertanyaan).trim()) return 'Pertanyaan wajib diisi.';
  const kunci = str_(o.Kunci_Jawaban).trim();
  if (!kunci) return 'Kunci jawaban wajib diisi.';
  if (o.Jenis_Soal === 'PG' || o.Jenis_Soal === 'PGK') {
    const ada = ['A', 'B', 'C', 'D'].filter(l => str_(o['Pilihan_' + l]).trim());
    if (ada.length < 2) return 'Isi minimal 2 pilihan jawaban.';
    const k = normSet_(kunci);
    if (!k.length || k.some(x => ada.indexOf(x) < 0)) return 'Kunci harus berupa huruf pilihan yang terisi.';
    if (o.Jenis_Soal === 'PG' && k.length !== 1) return 'Soal PG hanya boleh 1 kunci.';
    if (o.Jenis_Soal === 'PGK' && k.length < 1) return 'Soal PGK minimal 1 kunci.';
  } else {
    const n = [1, 2, 3, 4].filter(i => str_(o['Pernyataan_' + i]).trim()).length;
    if (n < 2) return 'Isi minimal 2 pernyataan.';
    if (kunci.split(',').length !== n) return 'Jumlah kunci (' + kunci.split(',').length + ') harus sama dengan jumlah pernyataan (' + n + ').';
  }
  return '';
}
function rapikanSoal_(o) {
  const r = {}; SHEETS.SOAL.forEach(h => r[h] = str_(o[h]).trim());
  if (r.Jenis_Soal === 'PG') r.Kunci_Jawaban = r.Kunci_Jawaban.toUpperCase();
  if (r.Jenis_Soal === 'PGK') r.Kunci_Jawaban = normSet_(r.Kunci_Jawaban).join(',');
  if (r.Jenis_Soal === 'PG_Kategori') r.Kunci_Jawaban = r.Kunci_Jawaban.split(',').map(x => x.trim()).join(',');
  if (!r.Level) r.Level = 'Sedang';
  if (!r.Status) r.Status = 'Aktif';
  if (!r.Kelas) r.Kelas = '6';
  return r;
}
function objToSoalRow_(o) { return SHEETS.SOAL.map(h => str_(o[h])); }
function findSoalRow_(id) {
  const r = read_('SOAL');
  for (let i = 0; i < r.data.length; i++) if (str_(r.data[i][1]) === id) return { row: i + 2, sheet: r.sheet, obj: soalRowToObj_(r.data[i]) };
  return null;
}
function renomori_() {
  const sh = getSS_().getSheetByName('SOAL'), lr = sh.getLastRow();
  if (lr < 2) return;
  const ids = sh.getRange(2, 2, lr - 1, 1).getValues();
  sh.getRange(2, 1, lr - 1, 1).setValues(ids.map((x, i) => [i + 1]));
}
/** Samakan data stimulus pada semua soal yang berbagi Stimulus_ID. */
function sinkronStimulus_(o) {
  if (!o.Stimulus_ID) return;
  const r = read_('SOAL');
  r.data.forEach((v, i) => {
    if (str_(v[5]) === o.Stimulus_ID) r.sheet.getRange(i + 2, 7, 1, 3).setValues([[o.Stimulus_Judul, o.Stimulus_Teks, o.Stimulus_Gambar_URL]]);
  });
}

/** Kelola soal. aksi: list | simpan | hapus | toggle | duplikat | import | contoh */
function kelolaSoal(t, aksi, p) {
  return admin_(t, () => {
    p = p || {};
    if (aksi === 'list') {
      return ok_(read_('SOAL').data.filter(v => str_(v[1]) !== '').map(soalRowToObj_));
    }
    if (aksi === 'simpan') {
      const o = rapikanSoal_(p.soal || {});
      const e = validasiSoal_(o); if (e) return fail_(e);
      if (o.Stimulus_ID === '' && (o.Stimulus_Teks || o.Stimulus_Judul || o.Stimulus_Gambar_URL)) o.Stimulus_ID = newId_('ST');
      if (o.ID_Soal) {
        const f = findSoalRow_(o.ID_Soal); if (!f) return fail_('Soal tidak ditemukan.');
        o.No = f.obj.No;
        f.sheet.getRange(f.row, 1, 1, SHEETS.SOAL.length).setValues([objToSoalRow_(o)]);
      } else {
        o.ID_Soal = newId_('S'); o.No = nextNo_();
        appendRow_('SOAL', objToSoalRow_(o));
      }
      sinkronStimulus_(o); clearSoalCache_(); SpreadsheetApp.flush();
      return ok_({ id: o.ID_Soal }, 'Soal tersimpan');
    }
    if (aksi === 'hapus') {
      const f = findSoalRow_(str_(p.id)); if (!f) return fail_('Soal tidak ditemukan.');
      f.sheet.deleteRow(f.row); renomori_(); clearSoalCache_();
      return ok_(null, 'Soal dihapus');
    }
    if (aksi === 'toggle') {
      const f = findSoalRow_(str_(p.id)); if (!f) return fail_('Soal tidak ditemukan.');
      const baru = f.obj.Status.toLowerCase() === 'aktif' ? 'Nonaktif' : 'Aktif';
      f.sheet.getRange(f.row, SHEETS.SOAL.indexOf('Status') + 1).setValue(baru); clearSoalCache_();
      return ok_({ status: baru }, 'Status soal: ' + baru);
    }
    if (aksi === 'duplikat') {
      const f = findSoalRow_(str_(p.id)); if (!f) return fail_('Soal tidak ditemukan.');
      const o = f.obj; o.ID_Soal = newId_('S'); o.No = nextNo_(); o.Pertanyaan = o.Pertanyaan + ' (salinan)';
      appendRow_('SOAL', objToSoalRow_(o)); clearSoalCache_();
      return ok_(null, 'Soal diduplikat');
    }
    if (aksi === 'import') return importSoal_(p.rows || []);
    if (aksi === 'contoh') return importSoal_(CONTOH_SOAL);
    return fail_('Aksi tidak dikenal.');
  }, true);
}

/** Import banyak soal (array objek). ID_Soal yang sudah ada akan diperbarui, yang baru ditambahkan. */
function importSoal_(rows) {
  const r = read_('SOAL'), idx = {};
  r.data.forEach((v, i) => { if (str_(v[1])) idx[str_(v[1])] = i + 2; });
  let tambah = 0, ubah = 0; const galat = [], baru = [];
  rows.forEach((raw, n) => {
    const o = rapikanSoal_(raw);
    const e = validasiSoal_(o);
    if (e) { galat.push('Baris ' + (n + 1) + ': ' + e); return; }
    if (o.ID_Soal && idx[o.ID_Soal]) { r.sheet.getRange(idx[o.ID_Soal], 1, 1, SHEETS.SOAL.length).setValues([objToSoalRow_(Object.assign({}, o, { No: r.sheet.getRange(idx[o.ID_Soal], 1).getValue() }))]); ubah++; }
    else { if (!o.ID_Soal) o.ID_Soal = newId_('S'); idx[o.ID_Soal] = -1; baru.push(objToSoalRow_(o)); tambah++; }
  });
  appendRows_('SOAL', baru);
  renomori_(); clearSoalCache_(); SpreadsheetApp.flush();
  const msg = tambah + ' soal ditambah, ' + ubah + ' diperbarui' + (galat.length ? ', ' + galat.length + ' baris gagal' : '') + '.';
  return ok_({ tambah: tambah, ubah: ubah, galat: galat.slice(0, 10) }, msg);
}

/** Ambil gambar dari Drive sebagai data URI, SEMENTARA file tidak perlu dibagikan publik.
 *  Hanya file bergambar <= 3 MB; hasil di-cache 6 jam supaya ratusan siswa tidak membebani Drive. */
function ambilGambarDrive_(id) {
  id = str_(id);
  if (!/^[A-Za-z0-9_-]{10,}$/.test(id)) throw new Error('ID_GAMBAR_TIDAK_VALID');
  const key = 'img_' + id, hit = cacheGet_(key);
  if (hit) return hit;
  let f; try { f = DriveApp.getFileById(id); } catch (e) { throw new Error('GAMBAR_TIDAK_ADA_ATAU_TANPA_AKSES'); }
  const mime = f.getMimeType();
  if (mime.indexOf('image/') !== 0) throw new Error('BUKAN_GAMBAR');
  if (f.getSize() > 3 * 1024 * 1024) throw new Error('GAMBAR_TERLALU_BESAR');
  const uri = 'data:' + mime + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
  if (uri.length < 900000) cachePut_(key, uri, 21600);
  return uri;
}
function pesanGambar_(e) {
  const m = String(e && e.message || '');
  if (m.indexOf('TANPA_AKSES') >= 0) return 'Gambar tidak ditemukan atau akun pemilik aplikasi tidak punya akses ke file ini.';
  if (m.indexOf('BUKAN_GAMBAR') >= 0) return 'File itu bukan gambar.';
  if (m.indexOf('TERLALU_BESAR') >= 0) return 'Gambar terlalu besar (maksimal 3 MB).';
  return 'Gambar tidak dapat dimuat.';
}
/** Untuk siswa yang sedang ujian (butuh sesi valid). */
function getGambarSiswa(sess, id) {
  return safe_(() => {
    if (!sesiCache_(sess)) return fail_('Sesi ujian tidak ditemukan.');
    try { return ok_({ uri: ambilGambarDrive_(id) }); } catch (e) { return fail_(pesanGambar_(e)); }
  });
}
/** Untuk pratinjau di panel admin. */
function getGambarAdmin(t, id) {
  return admin_(t, () => {
    try { return ok_({ uri: ambilGambarDrive_(id) }); } catch (e) { return fail_(pesanGambar_(e)); }
  });
}

/** Upload gambar stimulus ke Google Drive (folder CBT_TKA_Gambar). */
function uploadGambar(t, base64, mime, nama) {
  return admin_(t, () => {
    const it = DriveApp.getFoldersByName('CBT_TKA_Gambar');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('CBT_TKA_Gambar');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime || 'image/jpeg', nama || ('gambar_' + Date.now() + '.jpg'));
    const f = folder.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return ok_({ url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1000' }, 'Gambar terunggah');
  }, true);
}

// ===========================================================================
//  ADMIN: KELOLA TOKEN
// ===========================================================================
function randomToken_(pakai) {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (let k = 0; k < 50; k++) {
    let s = ''; for (let i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
    if (pakai.indexOf(s) < 0) return s;
  }
  return 'T' + Date.now().toString(36).toUpperCase();
}

/** Kelola token. aksi: list | buat | toggle | hapus */
function kelolaToken(t, aksi, p) {
  return admin_(t, (adm) => {
    p = p || {};
    if (aksi === 'list') return ok_(tokenList_().reverse());
    if (aksi === 'buat') {
      const mapel = str_(p.mapel) || 'Semua';
      if (mapel !== 'Semua' && MAPEL_LIST.indexOf(mapel) < 0) return fail_('Mapel tidak valid.');
      const tgl = fmtDate_(p.tanggal) || today_(), jm = fmtTime_(p.jamMulai), js = fmtTime_(p.jamSelesai);
      if (!jm || !js) return fail_('Isi jam mulai dan jam selesai.');
      if (js <= jm) return fail_('Jam selesai harus setelah jam mulai.');
      const kuota = Math.max(1, Math.min(500, Number(p.kuota) || 30)), jumlah = Math.max(1, Math.min(20, Number(p.jumlah) || 1));
      const pakai = tokenList_().map(x => x.token), baris = [], dibuat = [];
      for (let i = 0; i < jumlah; i++) {
        const tk = randomToken_(pakai.concat(dibuat));
        dibuat.push(tk); baris.push([tk, mapel, str_(p.kelas) || '6', tgl, jm, js, kuota, 0, 'YA', adm.u]);
      }
      appendRows_('TOKEN', baris);
      return ok_({ token: dibuat }, 'Token dibuat: ' + dibuat.join(', '));
    }
    const tk = tokenList_().filter(x => x.token === str_(p.token).toUpperCase())[0];
    if (!tk) return fail_('Token tidak ditemukan.');
    const sh = getSS_().getSheetByName('TOKEN');
    if (aksi === 'toggle') { sh.getRange(tk.row, 9).setValue(tk.aktif ? 'TIDAK' : 'YA'); return ok_({ aktif: !tk.aktif }, tk.aktif ? 'Token dinonaktifkan' : 'Token diaktifkan'); }
    if (aksi === 'hapus') { sh.deleteRow(tk.row); return ok_(null, 'Token dihapus'); }
    return fail_('Aksi tidak dikenal.');
  }, true);
}

// ===========================================================================
//  ADMIN: REKAP NILAI, RESET, ANALISIS
// ===========================================================================
/** Rekap nilai dengan filter {mapel, kelas, tanggal, predikat}. */
function getRekap(t, f) {
  return admin_(t, () => {
    sweeper_(); f = f || {};
    const rows = read_('NILAI').data.filter(v =>
      (!f.mapel || str_(v[5]) === f.mapel) && (!f.kelas || str_(v[4]) === str_(f.kelas)) &&
      (!f.tanggal || fmtDate_(v[0]) === f.tanggal) && (!f.predikat || str_(v[12]) === f.predikat))
      .map(v => ({
        waktu: fmtTs_(v[0]), token: str_(v[1]), nama: str_(v[2]), nisn: str_(v[3]), kelas: str_(v[4]), mapel: str_(v[5]),
        pg: Number(v[6]), pgk: Number(v[7]), kat: Number(v[8]), benar: Number(v[9]), total: Number(v[10]),
        nilai: Number(v[11]), predikat: str_(v[12]), durasi: Number(v[13]), status: str_(v[14])
      }));
    return ok_(rows.reverse());
  });
}

/** Reset ujian siswa (hapus sesi, jawaban, nilai) sehingga bisa ujian ulang. */
function resetSiswa(t, p) {
  return admin_(t, () => {
    const token = str_(p.token).toUpperCase(), nisn = str_(p.nisn), mapel = str_(p.mapel);
    const su = read_('SISWA_UJIAN').data.filter(v => str_(v[1]).toUpperCase() === token && str_(v[3]) === nisn && str_(v[5]) === mapel);
    su.forEach(v => { try { const u = JSON.parse(v[6]); CacheService.getScriptCache().removeAll(['ses_' + u.sk, 'jr_' + u.sk]); } catch (e) { } });
    deleteRows_('SISWA_UJIAN', v => str_(v[1]).toUpperCase() === token && str_(v[3]) === nisn && str_(v[5]) === mapel);
    deleteRows_('JAWABAN', v => str_(v[0]).toUpperCase() === token && str_(v[8]) === nisn && str_(v[2]) === mapel);
    deleteRows_('NILAI', v => str_(v[1]).toUpperCase() === token && str_(v[3]) === nisn && str_(v[5]) === mapel);
    const tk = tokenList_().filter(x => x.token === token)[0];
    if (tk && tk.terpakai > 0) getSS_().getSheetByName('TOKEN').getRange(tk.row, 8).setValue(tk.terpakai - 1);
    SpreadsheetApp.flush();
    return ok_(null, 'Ujian siswa direset. Siswa bisa mengulang dengan token yang sama.');
  }, true);
}

/** Hapus satu baris nilai saja (jawaban & sesi tetap). */
function hapusNilai(t, p) {
  return admin_(t, () => {
    const n = deleteRows_('NILAI', v => str_(v[1]).toUpperCase() === str_(p.token).toUpperCase() && str_(v[3]) === str_(p.nisn) && str_(v[5]) === str_(p.mapel));
    return ok_(null, n + ' nilai dihapus');
  }, true);
}

/** Analisis butir soal: persentase benar tiap soal, merah bila < 40%. */
function getAnalisis(t, f) {
  return admin_(t, () => {
    f = f || {};
    const bank = mapById_(loadSoal_()), st = {};
    read_('JAWABAN').data.forEach(v => {
      const b = str_(v[6]); if (!b) return;
      const id = str_(v[3]); if (!st[id]) st[id] = [0, 0];
      st[id][1]++; if (b === 'Benar') st[id][0]++;
    });
    const out = Object.keys(st).filter(id => bank[id] && (!f.mapel || bank[id].Mapel === f.mapel)).map(id => {
      const s = bank[id], persen = Math.round(st[id][0] / st[id][1] * 1000) / 10;
      return {
        id: id, no: s.No, mapel: s.Mapel, jenis: s.Jenis_Soal, tanya: s.Pertanyaan.substr(0, 140), level: s.Level,
        benar: st[id][0], peserta: st[id][1], persen: persen, bahas: persen < 40
      };
    }).sort((a, b) => a.persen - b.persen);
    return ok_(out);
  });
}

// ===========================================================================
//  ADMIN: PENGATURAN, BACKUP, HAPUS DATA
// ===========================================================================
/** Pengaturan. aksi: getConfig | simpanConfig | gantiPassword | hapusSemua */
function kelolaPengaturan(t, aksi, p) {
  return admin_(t, (adm) => {
    p = p || {};
    if (aksi === 'getConfig') return ok_(getConfig_());
    if (aksi === 'simpanConfig') {
      const sh = getSS_().getSheetByName('CONFIG'), r = read_('CONFIG'), pos = {};
      r.data.forEach((v, i) => pos[str_(v[0]).trim()] = i + 2);
      const num = ['durasi_menit', 'jumlah_soal', 'jumlah_pg', 'jumlah_pgk', 'jumlah_kategori', 'batas_pindah_tab'];
      for (const k in p.config) {
        if (!DEFAULT_CONFIG.some(d => d[0] === k)) continue;
        const val = str_(p.config[k]).trim();
        if (num.indexOf(k) >= 0 && !(Number(val) > 0)) return fail_('Nilai ' + k + ' harus angka lebih dari 0.');
        if (pos[k]) sh.getRange(pos[k], 2).setValue(val); else appendRow_('CONFIG', [k, val]);
      }
      SpreadsheetApp.flush();
      return ok_(null, 'Pengaturan disimpan');
    }
    if (aksi === 'gantiPassword') {
      const lama = str_(p.lama), baru = str_(p.baru);
      if (baru.length < 6) return fail_('Password baru minimal 6 karakter.');
      const r = read_('ADMIN');
      for (let i = 0; i < r.data.length; i++) {
        if (str_(r.data[i][0]) === adm.u) {
          if (str_(r.data[i][1]) !== sha256_(lama)) return fail_('Password lama salah.');
          r.sheet.getRange(i + 2, 2).setValue(sha256_(baru));
          return ok_(null, 'Password berhasil diganti');
        }
      }
      return fail_('Akun tidak ditemukan.');
    }
    if (aksi === 'hapusSemua') {
      ['SISWA_UJIAN', 'JAWABAN', 'NILAI'].forEach(n => {
        const sh = getSS_().getSheetByName(n), lr = sh.getLastRow();
        if (lr > 1) sh.deleteRows(2, lr - 1);
      });
      const tk = getSS_().getSheetByName('TOKEN'), lr = tk.getLastRow();
      if (lr > 1) tk.getRange(2, 8, lr - 1, 1).setValues(Array.apply(null, Array(lr - 1)).map(() => [0]));
      CacheService.getScriptCache().removeAll(['soal_all_n']);
      SpreadsheetApp.flush();
      return ok_(null, 'Semua data ujian dihapus. Soal dan token tetap ada.');
    }
    return fail_('Aksi tidak dikenal.');
  }, true);
}

function csvCell_(v) {
  if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  v = str_(v);
  return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}
/** Export sebuah sheet menjadi teks CSV (SOAL, TOKEN, SISWA_UJIAN, JAWABAN, NILAI, CONFIG, TEMPLATE_SOAL). */
function exportCSV(t, nama) {
  return admin_(t, () => {
    nama = str_(nama);
    if (nama === 'TEMPLATE_SOAL') {
      const rows = [SHEETS.SOAL.map(csvCell_).join(',')].concat(CONTOH_SOAL.map(o => SHEETS.SOAL.map(h => csvCell_(o[h])).join(',')));
      return ok_({ filename: 'template_soal_cbt_tka.csv', csv: rows.join('\r\n') });
    }
    if (!SHEETS[nama] || nama === 'ADMIN') return fail_('Sheet tidak dikenal.');
    const sh = getSS_().getSheetByName(nama), lr = sh.getLastRow();
    const vals = lr > 0 ? sh.getRange(1, 1, lr, SHEETS[nama].length).getValues() : [SHEETS[nama]];
    return ok_({ filename: nama.toLowerCase() + '_' + today_() + '.csv', csv: vals.map(r => r.map(csvCell_).join(',')).join('\r\n') });
  });
}

// ===========================================================================
//  10 CONTOH SOAL (tombol 'Muat Contoh Soal' di menu Kelola Soal)
// ===========================================================================
const CONTOH_SOAL = [
 {
  "No": 1,
  "ID_Soal": "BIN-01",
  "Mapel": "Bahasa Indonesia",
  "Kelas": "6",
  "Jenis_Soal": "PG",
  "Stimulus_ID": "",
  "Stimulus_Judul": "",
  "Stimulus_Teks": "",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Kalimat yang menggunakan huruf kapital dengan tepat adalah ...",
  "Pilihan_A": "Ibu pergi ke Pasar Baru bersama adik.",
  "Pilihan_B": "ibu pergi ke pasar baru bersama Adik.",
  "Pilihan_C": "Ibu pergi ke pasar Baru bersama adik.",
  "Pilihan_D": "Ibu Pergi ke pasar baru Bersama adik.",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "A",
  "Pembahasan": "Nama tempat seperti Pasar Baru diawali huruf kapital pada setiap katanya, sedangkan kata umum seperti ibu (di tengah kalimat) dan adik ditulis dengan huruf kecil.",
  "Level": "Mudah",
  "Status": "Aktif"
 },
 {
  "No": 2,
  "ID_Soal": "BIN-02",
  "Mapel": "Bahasa Indonesia",
  "Kelas": "6",
  "Jenis_Soal": "PGK",
  "Stimulus_ID": "",
  "Stimulus_Judul": "",
  "Stimulus_Teks": "",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Pilihlah SEMUA jawaban yang benar! Kata baku yang ditulis dengan tepat adalah ...",
  "Pilihan_A": "apotek",
  "Pilihan_B": "aktifitas",
  "Pilihan_C": "risiko",
  "Pilihan_D": "nasehat",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "A,C",
  "Pembahasan": "Bentuk baku yang benar: apotek, risiko, aktivitas, dan nasihat. Maka kata yang benar pada pilihan adalah apotek dan risiko.",
  "Level": "Sedang",
  "Status": "Aktif"
 },
 {
  "No": 3,
  "ID_Soal": "BIN-03",
  "Mapel": "Bahasa Indonesia",
  "Kelas": "6",
  "Jenis_Soal": "PG",
  "Stimulus_ID": "ST-BIN-1",
  "Stimulus_Judul": "Sungai Bersih, Desa Sehat",
  "Stimulus_Teks": "Warga Desa Mekarsari memiliki sungai jernih yang mengalir di tepi desa. Beberapa tahun lalu, sungai itu dipenuhi sampah plastik karena warga sering membuang sampah sembarangan. Air menjadi keruh, berbau, dan ikan-ikan mulai menghilang.\n\nPak Darto, ketua RT, mengajak warga bergotong royong setiap hari Minggu. Anak-anak ikut memungut sampah, sedangkan para ibu membuat tempat sampah dari ember bekas. Sebulan kemudian, air sungai kembali jernih dan anak-anak dapat bermain di tepinya dengan gembira.\n\nKini, warga memiliki aturan baru: dilarang membuang sampah ke sungai. Siapa pun yang melanggar harus ikut membersihkan sungai selama dua minggu.",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Apa penyebab sungai di Desa Mekarsari menjadi keruh dan berbau?",
  "Pilihan_A": "Hujan turun terus-menerus",
  "Pilihan_B": "Warga membuang sampah sembarangan",
  "Pilihan_C": "Pabrik membuang limbah ke sungai",
  "Pilihan_D": "Ikan di sungai terlalu banyak",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "B",
  "Pembahasan": "Pada paragraf pertama dijelaskan bahwa warga sering membuang sampah sembarangan sehingga air menjadi keruh dan berbau.",
  "Level": "Mudah",
  "Status": "Aktif"
 },
 {
  "No": 4,
  "ID_Soal": "BIN-04",
  "Mapel": "Bahasa Indonesia",
  "Kelas": "6",
  "Jenis_Soal": "PGK",
  "Stimulus_ID": "ST-BIN-1",
  "Stimulus_Judul": "Sungai Bersih, Desa Sehat",
  "Stimulus_Teks": "Warga Desa Mekarsari memiliki sungai jernih yang mengalir di tepi desa. Beberapa tahun lalu, sungai itu dipenuhi sampah plastik karena warga sering membuang sampah sembarangan. Air menjadi keruh, berbau, dan ikan-ikan mulai menghilang.\n\nPak Darto, ketua RT, mengajak warga bergotong royong setiap hari Minggu. Anak-anak ikut memungut sampah, sedangkan para ibu membuat tempat sampah dari ember bekas. Sebulan kemudian, air sungai kembali jernih dan anak-anak dapat bermain di tepinya dengan gembira.\n\nKini, warga memiliki aturan baru: dilarang membuang sampah ke sungai. Siapa pun yang melanggar harus ikut membersihkan sungai selama dua minggu.",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Pilihlah SEMUA jawaban yang benar! Pesan yang dapat diambil dari bacaan tersebut adalah ...",
  "Pilihan_A": "Gotong royong dapat menyelesaikan masalah lingkungan.",
  "Pilihan_B": "Sampah plastik sebaiknya dibuang ke sungai.",
  "Pilihan_C": "Menjaga kebersihan sungai adalah tugas kita bersama.",
  "Pilihan_D": "Hanya ketua RT yang wajib menjaga sungai.",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "A,C",
  "Pembahasan": "Bacaan menunjukkan bahwa gotong royong menyelesaikan masalah sampah, dan menjaga sungai adalah tanggung jawab bersama seluruh warga.",
  "Level": "Sedang",
  "Status": "Aktif"
 },
 {
  "No": 5,
  "ID_Soal": "BIN-05",
  "Mapel": "Bahasa Indonesia",
  "Kelas": "6",
  "Jenis_Soal": "PG_Kategori",
  "Stimulus_ID": "ST-BIN-1",
  "Stimulus_Judul": "Sungai Bersih, Desa Sehat",
  "Stimulus_Teks": "Warga Desa Mekarsari memiliki sungai jernih yang mengalir di tepi desa. Beberapa tahun lalu, sungai itu dipenuhi sampah plastik karena warga sering membuang sampah sembarangan. Air menjadi keruh, berbau, dan ikan-ikan mulai menghilang.\n\nPak Darto, ketua RT, mengajak warga bergotong royong setiap hari Minggu. Anak-anak ikut memungut sampah, sedangkan para ibu membuat tempat sampah dari ember bekas. Sebulan kemudian, air sungai kembali jernih dan anak-anak dapat bermain di tepinya dengan gembira.\n\nKini, warga memiliki aturan baru: dilarang membuang sampah ke sungai. Siapa pun yang melanggar harus ikut membersihkan sungai selama dua minggu.",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Tentukan apakah pernyataan berikut BENAR atau SALAH sesuai isi bacaan!",
  "Pilihan_A": "",
  "Pilihan_B": "",
  "Pilihan_C": "",
  "Pilihan_D": "",
  "Pernyataan_1": "Kegiatan gotong royong dilakukan setiap hari Minggu.",
  "Pernyataan_2": "Anak-anak tidak ikut membersihkan sungai.",
  "Pernyataan_3": "Pelanggar aturan harus membersihkan sungai selama dua minggu.",
  "Pernyataan_4": "Sungai menjadi jernih kembali setelah satu tahun.",
  "Kunci_Jawaban": "Benar,Salah,Benar,Salah",
  "Pembahasan": "(1) Benar: gotong royong dilakukan setiap hari Minggu. (2) Salah: anak-anak ikut memungut sampah. (3) Benar: pelanggar membersihkan sungai dua minggu. (4) Salah: sungai jernih setelah sebulan, bukan satu tahun.",
  "Level": "Sulit",
  "Status": "Aktif"
 },
 {
  "No": 6,
  "ID_Soal": "MTK-01",
  "Mapel": "Matematika",
  "Kelas": "6",
  "Jenis_Soal": "PG",
  "Stimulus_ID": "",
  "Stimulus_Judul": "",
  "Stimulus_Teks": "",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Hasil dari 3/4 + 1/8 adalah ...",
  "Pilihan_A": "4/12",
  "Pilihan_B": "7/8",
  "Pilihan_C": "5/8",
  "Pilihan_D": "1",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "B",
  "Pembahasan": "Samakan penyebut: 3/4 = 6/8, sehingga 6/8 + 1/8 = 7/8.",
  "Level": "Mudah",
  "Status": "Aktif"
 },
 {
  "No": 7,
  "ID_Soal": "MTK-02",
  "Mapel": "Matematika",
  "Kelas": "6",
  "Jenis_Soal": "PGK",
  "Stimulus_ID": "",
  "Stimulus_Judul": "",
  "Stimulus_Teks": "",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Pilihlah SEMUA jawaban yang benar! Bilangan berikut yang merupakan kelipatan 6 adalah ...",
  "Pilihan_A": "18",
  "Pilihan_B": "26",
  "Pilihan_C": "42",
  "Pilihan_D": "64",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "A,C",
  "Pembahasan": "Kelipatan 6 adalah bilangan yang habis dibagi 6. 18 = 6 x 3 dan 42 = 6 x 7. Bilangan 26 dan 64 tidak habis dibagi 6.",
  "Level": "Sedang",
  "Status": "Aktif"
 },
 {
  "No": 8,
  "ID_Soal": "MTK-03",
  "Mapel": "Matematika",
  "Kelas": "6",
  "Jenis_Soal": "PG",
  "Stimulus_ID": "ST-MTK-1",
  "Stimulus_Judul": "Penjualan Buah Toko Bu Sari",
  "Stimulus_Teks": "Bu Sari menjual apel dan jeruk di tokonya. Harga apel Rp20.000 per kg dan harga jeruk Rp15.000 per kg. Berikut data penjualan selama tiga hari.\n\n|Hari|Apel (kg)|Jeruk (kg)|\n|Senin|12|15|\n|Selasa|18|10|\n|Rabu|15|20|",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Berapa kilogram apel yang terjual selama tiga hari?",
  "Pilihan_A": "40 kg",
  "Pilihan_B": "43 kg",
  "Pilihan_C": "45 kg",
  "Pilihan_D": "48 kg",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "C",
  "Pembahasan": "Jumlahkan kolom apel: 12 + 18 + 15 = 45 kg.",
  "Level": "Mudah",
  "Status": "Aktif"
 },
 {
  "No": 9,
  "ID_Soal": "MTK-04",
  "Mapel": "Matematika",
  "Kelas": "6",
  "Jenis_Soal": "PGK",
  "Stimulus_ID": "ST-MTK-1",
  "Stimulus_Judul": "Penjualan Buah Toko Bu Sari",
  "Stimulus_Teks": "Bu Sari menjual apel dan jeruk di tokonya. Harga apel Rp20.000 per kg dan harga jeruk Rp15.000 per kg. Berikut data penjualan selama tiga hari.\n\n|Hari|Apel (kg)|Jeruk (kg)|\n|Senin|12|15|\n|Selasa|18|10|\n|Rabu|15|20|",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Pilihlah SEMUA jawaban yang benar! Pernyataan yang sesuai dengan tabel adalah ...",
  "Pilihan_A": "Jeruk paling banyak terjual pada hari Rabu.",
  "Pilihan_B": "Apel paling banyak terjual pada hari Senin.",
  "Pilihan_C": "Pada hari Selasa, apel terjual lebih banyak daripada jeruk.",
  "Pilihan_D": "Jeruk yang terjual selama tiga hari sebanyak 40 kg.",
  "Pernyataan_1": "",
  "Pernyataan_2": "",
  "Pernyataan_3": "",
  "Pernyataan_4": "",
  "Kunci_Jawaban": "A,C",
  "Pembahasan": "Jeruk terbanyak pada hari Rabu (20 kg). Pada hari Selasa apel 18 kg lebih banyak dari jeruk 10 kg. Apel terbanyak ada di hari Selasa, dan jumlah jeruk tiga hari adalah 45 kg.",
  "Level": "Sedang",
  "Status": "Aktif"
 },
 {
  "No": 10,
  "ID_Soal": "MTK-05",
  "Mapel": "Matematika",
  "Kelas": "6",
  "Jenis_Soal": "PG_Kategori",
  "Stimulus_ID": "ST-MTK-1",
  "Stimulus_Judul": "Penjualan Buah Toko Bu Sari",
  "Stimulus_Teks": "Bu Sari menjual apel dan jeruk di tokonya. Harga apel Rp20.000 per kg dan harga jeruk Rp15.000 per kg. Berikut data penjualan selama tiga hari.\n\n|Hari|Apel (kg)|Jeruk (kg)|\n|Senin|12|15|\n|Selasa|18|10|\n|Rabu|15|20|",
  "Stimulus_Gambar_URL": "",
  "Pertanyaan": "Tentukan apakah pernyataan berikut BENAR atau SALAH berdasarkan tabel!",
  "Pilihan_A": "",
  "Pilihan_B": "",
  "Pilihan_C": "",
  "Pilihan_D": "",
  "Pernyataan_1": "Pada hari Senin, jeruk terjual lebih banyak daripada apel.",
  "Pernyataan_2": "Jumlah apel dan jeruk yang terjual pada hari Rabu adalah 30 kg.",
  "Pernyataan_3": "Pendapatan dari apel pada hari Selasa adalah Rp360.000.",
  "Pernyataan_4": "Jeruk yang terjual selama tiga hari sebanyak 50 kg.",
  "Kunci_Jawaban": "Benar,Salah,Benar,Salah",
  "Pembahasan": "(1) 15 > 12, benar. (2) 15 + 20 = 35, bukan 30, salah. (3) 18 x Rp20.000 = Rp360.000, benar. (4) Jeruk 15 + 10 + 20 = 45 kg, bukan 50 kg, salah.",
  "Level": "Sulit",
  "Status": "Aktif"
 }
];
