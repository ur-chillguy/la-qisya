/**
 * LA QISYA — Rental & Sewa Gaun
 * Backend Google Apps Script + Google Sheets + Google Drive
 * Dibuat oleh ur.chillpartner (instagram.com/ur.chillguy)
 *
 * CARA PAKAI SINGKAT
 * 1. Buat Google Spreadsheet baru → Extensions → Apps Script.
 * 2. Tempel file ini ke Code.gs, buat file HTML bernama "Index" lalu tempel Index.html.
 * 3. Jalankan fungsi setup() sekali (izinkan akses).
 * 4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 * PIN admin awal: 123456 (segera ganti di menu Pengaturan).
 */

const SHEETS = {
  GAUN: 'Gaun',
  BOOKING: 'Booking',
  TRANSAKSI: 'Transaksi',
  SETTING: 'Pengaturan'
};

const HEADERS = {
  Gaun: ['id', 'kode', 'nama', 'kategori', 'warna', 'harga', 'hargaTambahan', 'nilaiGaun', 'ukuran', 'foto', 'deskripsi', 'status', 'unggulan', 'dibuat', 'diubah'],
  Booking: ['id', 'kode', 'gaunId', 'gaunNama', 'ukuran', 'nama', 'wa', 'alamat', 'tglAmbil', 'tglKembali', 'hari', 'total', 'dp', 'dibayar', 'status', 'sumber', 'catatan', 'catatanAdmin', 'dibuat', 'diubah'],
  Transaksi: ['id', 'tanggal', 'jenis', 'kategori', 'jumlah', 'keterangan', 'bookingKode', 'metode', 'dibuat']
};

const DEFAULT_SETTINGS = {
  namaToko: 'La Qisya',
  tagline: 'Elegance for Every Occasion',
  waAdmin: '6281234567890',
  alamat: 'Alamat butik La Qisya',
  jamBuka: 'Setiap hari, 09.00 – 20.00',
  instagram: '',
  durasiPaket: 3,
  dpPersen: 50,
  jedaCuci: 1,
  batasDpJam: 24,
  rekening: 'BCA 1234567890 a.n. La Qisya',
  qrisUrl: '',
  kategori: 'Gaun Pengantin,Kebaya,Gaun Pesta,Bridesmaid,Muslimah,Anak',
  kategoriPengeluaran: 'Laundry,Perbaikan,Beli Gaun,Operasional,Promosi,Lainnya',
  syarat: 'DP 50% saat booking, pelunasan saat pengambilan gaun.\nGaun dikembalikan dalam kondisi baik pada tanggal kembali.\nPembatalan: DP tidak dapat dikembalikan.'
};

/* ===================== SETUP ===================== */

function setup() {
  const props0 = PropertiesService.getScriptProperties();
  let ss = null;
  if (!props0.getProperty('SHEET_ID')) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) { ss = SpreadsheetApp.create('La Qisya - Data'); }
    props0.setProperty('SHEET_ID', ss.getId());
  } else { ss = SpreadsheetApp.openById(props0.getProperty('SHEET_ID')); }
  _ss = ss;
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#F3DDD3');
    }
  });
  let st = ss.getSheetByName(SHEETS.SETTING);
  if (!st) st = ss.insertSheet(SHEETS.SETTING);
  if (st.getLastRow() === 0) {
    st.appendRow(['key', 'value']);
    st.setFrozenRows(1);
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) { st.appendRow([k, DEFAULT_SETTINGS[k]]); });
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN_HASH')) props.setProperty('PIN_HASH', hash_('123456'));
  if (!props.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('La Qisya - Foto Gaun');
    props.setProperty('FOLDER_ID', folder.getId());
  }
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  Logger.log('Spreadsheet: ' + ss.getUrl());
  return 'Setup selesai. PIN awal 123456. Spreadsheet: ' + ss.getUrl();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('La Qisya — Sewa Gaun')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = api(req.action, req.payload, req.token);
  } catch (err) { out = { ok: false, message: 'Permintaan tidak valid.' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== API ROUTER ===================== */

function api(action, payload, token) {
  try {
    payload = payload || {};
    const pub = {
      getCatalog: getCatalog_,
      getAvailability: getAvailability_,
      createBooking: createBooking_,
      checkBooking: checkBooking_,
      login: login_
    };
    const admin = {
      getAdminData: getAdminData_,
      saveDress: saveDress_,
      deleteDress: deleteDress_,
      uploadPhoto: uploadPhoto_,
      saveBooking: saveBookingAdmin_,
      updateBookingStatus: updateBookingStatus_,
      deleteBooking: deleteBooking_,
      saveTransaction: saveTransaction_,
      deleteTransaction: deleteTransaction_,
      saveSettings: saveSettings_,
      changePin: changePin_,
      logout: logout_
    };
    if (pub[action]) return { ok: true, data: pub[action](payload) };
    if (admin[action]) {
      if (!checkToken_(token)) return { ok: false, error: 'SESSION', message: 'Sesi berakhir, silakan masuk lagi.' };
      return { ok: true, data: admin[action](payload, token) };
    }
    return { ok: false, message: 'Aksi tidak dikenal: ' + action };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

/* ===================== HELPERS ===================== */

let _ss = null;
function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) { _ss = SpreadsheetApp.openById(id); return _ss; }
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (a) { _ss = a; return _ss; }
  throw new Error('Database belum siap. Jalankan fungsi setup() dulu di editor Apps Script.');
}

function sheet_(name) { return ss_().getSheetByName(name); }

function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  return values.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) {
      let v = r[i];
      if (v instanceof Date) {
        const tz = Session.getScriptTimeZone();
        const t = Utilities.formatDate(v, tz, 'HH:mm:ss');
        v = Utilities.formatDate(v, tz, t === '00:00:00' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      }
      o[h] = v;
    });
    return o;
  });
}

function upsert_(name, obj) {
  const sh = sheet_(name);
  const head = HEADERS[name];
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return String(r[0]); });
  const row = head.map(function (h) {
    const v = obj[h];
    if (v === undefined || v === null) return '';
    return typeof v === 'object' ? JSON.stringify(v) : v;
  });
  const idx = ids.indexOf(String(obj.id));
  if (idx > 0) sh.getRange(idx + 1, 1, 1, head.length).setValues([row]);
  else sh.appendRow(row);
  return obj;
}

function deleteRow_(name, id) {
  const sh = sheet_(name);
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  const idx = ids.indexOf(String(id));
  if (idx > 0) sh.deleteRow(idx + 1);
  return true;
}

function parseJson_(v, fb) { if (v && typeof v === 'object') return v; try { return v ? JSON.parse(v) : fb; } catch (e) { return fb; } }

function uid_() { return Utilities.getUuid().slice(0, 8); }

function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }

function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function hash_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'laqisya::' + s);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function getSettings_() {
  const sh = sheet_(SHEETS.SETTING);
  const vals = sh.getDataRange().getValues().slice(1);
  const o = Object.assign({}, DEFAULT_SETTINGS);
  vals.forEach(function (r) { if (r[0]) o[r[0]] = r[1]; });
  ['durasiPaket', 'dpPersen', 'jedaCuci', 'batasDpJam'].forEach(function (k) { o[k] = Number(o[k]) || 0; });
  return o;
}

function dressOut_(d) {
  d.ukuran = parseJson_(d.ukuran, []);
  d.foto = parseJson_(d.foto, []);
  d.harga = Number(d.harga) || 0;
  d.hargaTambahan = Number(d.hargaTambahan) || 0;
  d.nilaiGaun = Number(d.nilaiGaun) || 0;
  d.unggulan = d.unggulan === true || d.unggulan === 'TRUE' || d.unggulan === 'true';
  return d;
}

function bookingOut_(b) {
  ['hari', 'total', 'dp', 'dibayar'].forEach(function (k) { b[k] = Number(b[k]) || 0; });
  b.wa = String(b.wa);
  return b;
}

/* --- ketersediaan --- */
const ACTIVE_STATUS = ['Menunggu DP', 'Terkonfirmasi', 'Diambil'];

function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.toISOString().slice(0, 10);
}

function overlaps_(a1, a2, b1, b2) { return a1 <= b2 && b1 <= a2; }

function countUsed_(bookings, gaunId, ukuran, start, end, jeda, excludeId) {
  return bookings.filter(function (b) {
    if (b.gaunId !== gaunId || b.ukuran !== ukuran) return false;
    if (ACTIVE_STATUS.indexOf(b.status) < 0) return false;
    if (excludeId && b.id === excludeId) return false;
    return overlaps_(start, addDays_(end, jeda), b.tglAmbil, addDays_(b.tglKembali, jeda));
  }).length;
}

/* ===================== PUBLIC ===================== */

function getCatalog_() {
  const s = getSettings_();
  const dresses = readAll_(SHEETS.GAUN).map(dressOut_).filter(function (d) { return d.status === 'Aktif'; });
  delete s.syaratInternal;
  return { settings: s, dresses: dresses.map(function (d) { delete d.nilaiGaun; return d; }) };
}

function getAvailability_(p) {
  const s = getSettings_();
  const list = readAll_(SHEETS.BOOKING).map(bookingOut_).filter(function (b) {
    return b.gaunId === p.gaunId && ACTIVE_STATUS.indexOf(b.status) >= 0 && b.tglKembali >= today_();
  });
  // hanya rentang tanggal, tanpa data pribadi
  return {
    jeda: s.jedaCuci,
    ranges: list.map(function (b) { return { ukuran: b.ukuran, mulai: b.tglAmbil, selesai: addDays_(b.tglKembali, s.jedaCuci) }; })
  };
}

function calcTotal_(dress, hari, s) {
  const paket = s.durasiPaket || 3;
  return dress.harga + Math.max(0, hari - paket) * dress.hargaTambahan;
}

function createBooking_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = getSettings_();
    if (!p.nama || !p.wa || !p.gaunId || !p.ukuran || !p.tglAmbil || !p.hari) throw new Error('Data booking belum lengkap.');
    if (p.tglAmbil < today_()) throw new Error('Tanggal ambil sudah lewat.');
    const dress = readAll_(SHEETS.GAUN).map(dressOut_).filter(function (d) { return d.id === p.gaunId; })[0];
    if (!dress || dress.status !== 'Aktif') throw new Error('Gaun tidak tersedia.');
    const size = dress.ukuran.filter(function (u) { return u.ukuran === p.ukuran; })[0];
    if (!size) throw new Error('Ukuran tidak tersedia.');
    const hari = Math.max(Number(p.hari) || s.durasiPaket, 1);
    const tglKembali = addDays_(p.tglAmbil, hari - 1);
    const bookings = readAll_(SHEETS.BOOKING).map(bookingOut_);
    const used = countUsed_(bookings, dress.id, p.ukuran, p.tglAmbil, tglKembali, s.jedaCuci);
    if (used >= Number(size.stok)) throw new Error('Maaf, ukuran ' + p.ukuran + ' sudah terisi di tanggal tersebut. Silakan pilih tanggal/ukuran lain.');
    const total = calcTotal_(dress, hari, s);
    const b = {
      id: uid_(),
      kode: 'LQ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase(),
      gaunId: dress.id, gaunNama: dress.nama, ukuran: p.ukuran,
      nama: String(p.nama).slice(0, 80), wa: normWa_(p.wa), alamat: String(p.alamat || '').slice(0, 200),
      tglAmbil: p.tglAmbil, tglKembali: tglKembali, hari: hari,
      total: total, dp: Math.round(total * s.dpPersen / 100), dibayar: 0,
      status: 'Menunggu DP', sumber: 'Online', catatan: String(p.catatan || '').slice(0, 300), catatanAdmin: '',
      dibuat: now_(), diubah: now_()
    };
    upsert_(SHEETS.BOOKING, b);
    return b;
  } finally { lock.releaseLock(); }
}

function normWa_(wa) {
  let w = String(wa).replace(/\D/g, '');
  if (w.indexOf('0') === 0) w = '62' + w.slice(1);
  return w;
}

function checkBooking_(p) {
  const kode = String(p.kode || '').trim().toUpperCase();
  const wa = normWa_(p.wa || '');
  const b = readAll_(SHEETS.BOOKING).map(bookingOut_).filter(function (x) { return String(x.kode).toUpperCase() === kode && x.wa === wa; })[0];
  if (!b) throw new Error('Booking tidak ditemukan. Cek kembali kode & nomor WA.');
  delete b.catatanAdmin;
  return b;
}

/* ===================== AUTH ===================== */

function login_(p) {
  const props = PropertiesService.getScriptProperties();
  if (hash_(String(p.pin)) !== props.getProperty('PIN_HASH')) throw new Error('PIN salah.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + token, '1', 21600);
  return { token: token };
}

function checkToken_(t) { return !!t && CacheService.getScriptCache().get('tok_' + t) === '1'; }

function logout_(p, token) { CacheService.getScriptCache().remove('tok_' + token); return true; }

function changePin_(p) {
  const props = PropertiesService.getScriptProperties();
  if (hash_(String(p.lama)) !== props.getProperty('PIN_HASH')) throw new Error('PIN lama salah.');
  if (!/^\d{4,8}$/.test(String(p.baru))) throw new Error('PIN baru harus 4–8 angka.');
  props.setProperty('PIN_HASH', hash_(String(p.baru)));
  return true;
}

/* ===================== ADMIN ===================== */

function getAdminData_() {
  return {
    settings: getSettings_(),
    dresses: readAll_(SHEETS.GAUN).map(dressOut_),
    bookings: readAll_(SHEETS.BOOKING).map(bookingOut_),
    transactions: readAll_(SHEETS.TRANSAKSI).map(function (t) { t.jumlah = Number(t.jumlah) || 0; return t; })
  };
}

function saveDress_(d) {
  if (!d.nama) throw new Error('Nama gaun wajib diisi.');
  const isNew = !d.id;
  if (isNew) { d.id = uid_(); d.dibuat = now_(); }
  d.diubah = now_();
  d.harga = Number(d.harga) || 0;
  d.hargaTambahan = Number(d.hargaTambahan) || 0;
  d.nilaiGaun = Number(d.nilaiGaun) || 0;
  upsert_(SHEETS.GAUN, d);
  return dressOut_(Object.assign({}, d));
}

function deleteDress_(p) {
  const active = readAll_(SHEETS.BOOKING).filter(function (b) { return b.gaunId === p.id && ACTIVE_STATUS.indexOf(b.status) >= 0; });
  if (active.length) throw new Error('Gaun masih punya ' + active.length + ' booking aktif. Ubah status gaun ke "Arsip" saja.');
  return deleteRow_(SHEETS.GAUN, p.id);
}

function uploadPhoto_(p) {
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const data = String(p.dataUrl).split(',');
  const mime = data[0].match(/data:(.*?);/)[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(data[1]), mime, p.nama || ('foto-' + uid_() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200', fileId: file.getId() };
}

function saveBookingAdmin_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = getSettings_();
    const bookings = readAll_(SHEETS.BOOKING).map(bookingOut_);
    const old = p.id ? bookings.filter(function (b) { return b.id === p.id; })[0] : null;
    const dress = readAll_(SHEETS.GAUN).map(dressOut_).filter(function (d) { return d.id === p.gaunId; })[0];
    if (!dress) throw new Error('Gaun tidak ditemukan.');
    const size = dress.ukuran.filter(function (u) { return u.ukuran === p.ukuran; })[0];
    if (!size) throw new Error('Ukuran tidak ada di gaun ini.');
    const hari = Math.max(Number(p.hari) || 1, 1);
    const tglKembali = addDays_(p.tglAmbil, hari - 1);
    const status = p.status || (old ? old.status : 'Terkonfirmasi');
    if (ACTIVE_STATUS.indexOf(status) >= 0 && !p.paksa) {
      const used = countUsed_(bookings, dress.id, p.ukuran, p.tglAmbil, tglKembali, s.jedaCuci, p.id);
      if (used >= Number(size.stok)) throw new Error('Bentrok: stok ukuran ' + p.ukuran + ' sudah terpakai di tanggal itu.');
    }
    const total = p.total !== undefined && p.total !== '' ? Number(p.total) : calcTotal_(dress, hari, s);
    const b = Object.assign({}, old || {}, {
      id: p.id || uid_(),
      kode: (old && old.kode) || ('LQ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase()),
      gaunId: dress.id, gaunNama: dress.nama, ukuran: p.ukuran,
      nama: p.nama, wa: normWa_(p.wa), alamat: p.alamat || '',
      tglAmbil: p.tglAmbil, tglKembali: tglKembali, hari: hari,
      total: total, dp: Math.round(total * s.dpPersen / 100),
      dibayar: old ? old.dibayar : 0,
      status: status, sumber: (old && old.sumber) || p.sumber || 'Admin',
      catatan: p.catatan || '', catatanAdmin: p.catatanAdmin || '',
      dibuat: (old && old.dibuat) || now_(), diubah: now_()
    });
    upsert_(SHEETS.BOOKING, b);
    return b;
  } finally { lock.releaseLock(); }
}

/**
 * Ubah status + catat pembayaran ke keuangan (opsional)
 * p: {id, status, bayar (angka), metode}
 */
function updateBookingStatus_(p) {
  const b = readAll_(SHEETS.BOOKING).map(bookingOut_).filter(function (x) { return x.id === p.id; })[0];
  if (!b) throw new Error('Booking tidak ditemukan.');
  const bayar = Number(p.bayar) || 0;
  if (bayar > 0) {
    b.dibayar += bayar;
    upsert_(SHEETS.TRANSAKSI, {
      id: uid_(), tanggal: today_(), jenis: 'Pemasukan',
      kategori: p.status === 'Terkonfirmasi' ? 'DP Sewa' : 'Pelunasan Sewa',
      jumlah: bayar, keterangan: b.gaunNama + ' (' + b.ukuran + ') — ' + b.nama,
      bookingKode: b.kode, metode: p.metode || 'Transfer', dibuat: now_()
    });
  }
  if (p.status) b.status = p.status;
  if (p.catatanAdmin !== undefined) b.catatanAdmin = p.catatanAdmin;
  b.diubah = now_();
  upsert_(SHEETS.BOOKING, b);
  return b;
}

function deleteBooking_(p) { return deleteRow_(SHEETS.BOOKING, p.id); }

function saveTransaction_(t) {
  if (!t.jumlah || !t.jenis) throw new Error('Jenis & jumlah wajib diisi.');
  if (!t.id) { t.id = uid_(); t.dibuat = now_(); }
  t.jumlah = Number(t.jumlah);
  upsert_(SHEETS.TRANSAKSI, t);
  return t;
}

function deleteTransaction_(p) { return deleteRow_(SHEETS.TRANSAKSI, p.id); }

function saveSettings_(p) {
  const sh = sheet_(SHEETS.SETTING);
  const vals = sh.getDataRange().getValues();
  const keys = vals.map(function (r) { return r[0]; });
  Object.keys(p).forEach(function (k) {
    if (!(k in DEFAULT_SETTINGS)) return;
    const i = keys.indexOf(k);
    if (i > 0) sh.getRange(i + 1, 2).setValue(p[k]);
    else sh.appendRow([k, p[k]]);
  });
  return getSettings_();
}
