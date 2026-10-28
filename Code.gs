/**
 * PEMBUKUAN KEUANGAN KKKS Ki Hajar Dewantara — Kecamatan Reban
 * Sumber data: sheet "KewajibanSD2026"
 *
 * File: Code.gs + Index.html saja.
 * Deploy sebagai Web App (Execute as: Me, Who has access: Anyone di organisasi).
 */

var SHEET_DATA = 'KewajibanSD2026';
var SHEET_AKSES = 'Akses';
var SHEET_PENGUMUMAN = 'Pengumuman';
var DEFAULT_ADMIN_PASS = 'AdminKKKS2026';
var ORG_NAME = 'KKKS Ki Hajar Dewantara';
var ORG_WILAYAH = 'Kecamatan Reban';

/* ========== ENTRY ========== */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pembukuan KKKS — ' + ORG_WILAYAH)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ========== HELPERS ========== */
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sh_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan. Pastikan script terpasang di spreadsheet yang benar.');
  return sh;
}

function ensureSheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
  }
  return sh;
}

function clean_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}

function parseMoney_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? Math.round(v) : 0;
  var s = String(v).replace(/Rp\s*/gi, '').replace(/\./g, '').replace(/,/g, '.').replace(/[^\d.-]/g, '');
  var n = parseFloat(s);
  return isFinite(n) ? Math.round(n) : 0;
}

function stamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
}

/* ========== AKSES / LOGIN ========== */
function ensureAksesSheet_() {
  var data = readKewajiban_();
  var sh = ensureSheet_(SHEET_AKSES, ['No', 'NamaSekolah', 'Password', 'Aktif']);
  var existing = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last, 3).getValues();
    for (var i = 0; i < vals.length; i++) {
      var nm = clean_(vals[i][1]).toLowerCase();
      if (nm) existing[nm] = true;
    }
  }
  // Tambah sekolah yang belum ada; password default = nomor urut
  for (var j = 0; j < data.schools.length; j++) {
    var s = data.schools[j];
    var key = clean_(s.nama).toLowerCase();
    if (!existing[key]) {
      sh.appendRow([s.no, s.nama, String(s.no), 'Ya']);
    }
  }
  return sh;
}

function ensurePengumumanSheet_() {
  var sh = ensureSheet_(SHEET_PENGUMUMAN, ['Waktu', 'Isi', 'Aktif']);
  if (sh.getLastRow() < 2) {
    sh.appendRow([stamp_(), 'Selamat datang di portal keuangan KKKS Ki Hajar Dewantara Kecamatan Reban. Harap segera melunasi tagihan yang masih kosong.', 'Ya']);
  }
  return sh;
}

function getPengaturan_() {
  var out = {
    PasswordAdmin: DEFAULT_ADMIN_PASS,
    Ketua: 'Ketua KKKS',
    Bendahara: 'Bendahara KKKS',
    NamaOrganisasi: ORG_NAME,
    Wilayah: ORG_WILAYAH
  };
  try {
    var sh = ss_().getSheetByName('Pengaturan');
    if (!sh) return out;
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      var k = clean_(vals[i][0]);
      if (!k) continue;
      out[k] = clean_(vals[i][1]);
    }
  } catch (e) {}
  if (!out.PasswordAdmin) out.PasswordAdmin = DEFAULT_ADMIN_PASS;
  return out;
}

function getAdminPassword_() {
  var fromSheet = getPengaturan_().PasswordAdmin;
  if (fromSheet) return String(fromSheet);
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('ADMIN_PASS') || DEFAULT_ADMIN_PASS;
}

function listSekolahPublic() {
  try {
    ensureAksesSheet_();
    var data = readKewajiban_();
    return {
      ok: true,
      org: ORG_NAME,
      wilayah: ORG_WILAYAH,
      sekolah: data.schools.map(function (s) {
        return { no: s.no, nama: s.nama };
      })
    };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function loginAdmin(password) {
  try {
    var pass = clean_(password);
    if (!pass) return { ok: false, message: 'Password wajib diisi.' };
    if (pass !== getAdminPassword_()) return { ok: false, message: 'Password admin salah.' };
    ensureAksesSheet_();
    ensurePengumumanSheet_();
    log_('ADMIN', 'Login Admin', 'Berhasil');
    return {
      ok: true,
      role: 'admin',
      nama: 'Administrator / Bendahara',
      token: makeToken_('admin', 'ADMIN')
    };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function loginVisitor(namaSekolah, password) {
  try {
    var nama = clean_(namaSekolah);
    var pass = clean_(password);
    if (!nama) return { ok: false, message: 'Pilih sekolah.' };
    if (!pass) return { ok: false, message: 'Isi password.' };

    ensureAksesSheet_();
    var sh = sh_(SHEET_AKSES);
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, message: 'Data akses kosong.' };
    var vals = sh.getRange(2, 1, last, 4).getValues();
    var found = null;
    for (var i = 0; i < vals.length; i++) {
      if (clean_(vals[i][1]).toLowerCase() === nama.toLowerCase()) {
        found = { no: vals[i][0], nama: clean_(vals[i][1]), pass: clean_(vals[i][2]), aktif: clean_(vals[i][3]) };
        break;
      }
    }
    if (!found) return { ok: false, message: 'Sekolah tidak terdaftar di sheet Akses.' };
    if (found.aktif && found.aktif.toLowerCase() === 'tidak') {
      return { ok: false, message: 'Akses sekolah dinonaktifkan. Hubungi admin.' };
    }
    if (found.pass !== pass) return { ok: false, message: 'Password sekolah salah.' };

    log_(found.nama, 'Login Pengunjung', 'OK');
    return {
      ok: true,
      role: 'visitor',
      nama: found.nama,
      no: found.no,
      token: makeToken_('visitor', found.nama)
    };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function makeToken_(role, id) {
  var payload = role + '|' + id + '|' + new Date().getTime();
  return Utilities.base64EncodeWebSafe(payload);
}

function parseToken_(token) {
  try {
    var s = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    var p = s.split('|');
    return { role: p[0], id: p[1], ts: Number(p[2]) || 0 };
  } catch (e) {
    return null;
  }
}

function requireAdmin_(token) {
  var t = parseToken_(token);
  if (!t || t.role !== 'admin') throw new Error('Sesi admin tidak valid. Login ulang.');
  return t;
}

function requireVisitor_(token) {
  var t = parseToken_(token);
  if (!t || t.role !== 'visitor') throw new Error('Sesi pengunjung tidak valid. Login ulang.');
  return t;
}

function log_(user, aksi, detail) {
  try {
    var sh = ensureSheet_('LogAktivitas', ['Waktu', 'Pengguna', 'Aksi', 'Detail']);
    sh.appendRow([stamp_(), user || '', aksi || '', detail || '']);
  } catch (e) {}
}

/* ========== BACA KEWAJIBANSD2026 ==========
 * Struktur sheet (baris header, 1-based):
 *   1 judul
 *   2 nama kelompok (sel gabung: GOTONG ROYONG SISWA, DANA LOMBA, MMT, ...)
 *   3 subkegiatan (BATANG, TW.1, ASAT VI, ...)
 *   4 label TAGIHAN | BAYAR
 *   5+ data sekolah, diakhiri baris JUMLAH
 */
function findTagihanRow_(values) {
  var limit = Math.min(values.length, 10);
  for (var r = 0; r < limit; r++) {
    var n = 0;
    var row = values[r];
    for (var c = 0; c < row.length; c++) {
      if (clean_(row[c]).toUpperCase() === 'TAGIHAN') n++;
    }
    if (n >= 5) return r;
  }
  return -1;
}

function itemStatus_(tagihan, bayar) {
  if (tagihan <= 0 && bayar <= 0) return 'Kosong';
  if (tagihan > 0 && bayar >= tagihan) return 'Lunas';
  if (bayar > 0 && bayar < tagihan) return 'Sebagian';
  return 'Belum';
}

function readLayout_() {
  var sh = sh_(SHEET_DATA);
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { sheet: sh, values: values || [], activities: [], trow: -1, schools: [], totalCols: {} };
  }

  var trow = findTagihanRow_(values);
  var activities = [];
  var totalCols = { tagihan: -1, bayar: -1, kurang: -1 };

  if (trow < 0) {
    // cadangan: header 1 baris "NAMA TAGIHAN | BAYAR"
    var header = values[0];
    for (var c0 = 2; c0 < header.length; c0++) {
      var h0 = clean_(header[c0]);
      var up0 = h0.toUpperCase();
      if (up0.indexOf('TOTAL TAGIHAN') === 0) { totalCols.tagihan = c0; continue; }
      if (up0.indexOf('TOTAL TERBAYAR') === 0 || up0.indexOf('TOTAL BAYAR') === 0) { totalCols.bayar = c0; continue; }
      if (up0.indexOf('KEKURANGAN') === 0) { totalCols.kurang = c0; continue; }
      if (up0.indexOf('TAGIHAN') >= 0) {
        var nm = h0.replace(/\s*TAGIHAN\s*$/i, '').trim() || ('Kegiatan ' + (c0 + 1));
        activities.push({ name: nm, tagihanCol: c0, bayarCol: c0 + 1 });
        c0++;
      }
    }
  } else {
    var groupRow = trow >= 2 ? values[trow - 2] : [];
    var subRow = trow >= 1 ? values[trow - 1] : [];
    var labelRow = values[trow];
    var width = Math.max(labelRow.length, groupRow.length, subRow.length);
    var filled = [];
    var lastGroup = '';
    for (var c = 0; c < width; c++) {
      var rawG = clean_(groupRow[c]);
      var upG = rawG.toUpperCase();
      if (upG.indexOf('TOTAL') === 0 || upG === 'KEKURANGAN' || upG === 'NO' || upG.indexOf('SEKOLAH') === 0) {
        lastGroup = '';
        filled[c] = rawG;
      } else if (rawG) {
        lastGroup = rawG;
        filled[c] = rawG;
      } else {
        filled[c] = lastGroup;
      }
      var upLab = clean_(labelRow[c]).toUpperCase();
      if (upG.indexOf('TOTAL TAGIHAN') === 0) totalCols.tagihan = c;
      if (upG.indexOf('TOTAL TERBAYAR') === 0 || upG.indexOf('TOTAL BAYAR') === 0) totalCols.bayar = c;
      if (upG === 'KEKURANGAN' || upG.indexOf('KEKURANGAN') === 0) totalCols.kurang = c;
      if (upLab !== 'TAGIHAN') continue;

      var rawHere = clean_(groupRow[c]);
      var sub = clean_(subRow[c]);
      // kolom cadangan (MMT / CETAK SOAL tanpa subjudul) dilewati
      if (!sub && !rawHere) continue;

      var group = filled[c];
      var name = sub && group ? (group + ' ' + sub) : (sub || group || ('Kegiatan ' + (c + 1)));
      var bayarCol = c + 1;
      if (clean_(labelRow[bayarCol]).toUpperCase() !== 'BAYAR') bayarCol = -1;
      activities.push({ name: name, tagihanCol: c, bayarCol: bayarCol });
    }
  }

  var start = trow < 0 ? 1 : trow + 1;
  var schools = [];
  for (var r = start; r < values.length; r++) {
    var row = values[r];
    var no = row[0];
    var nama = clean_(row[1]);
    if (!nama) continue;
    var upNo = clean_(no).toUpperCase();
    var upNama = nama.toUpperCase();
    if (upNo.indexOf('JUMLAH') >= 0 || upNama.indexOf('JUMLAH') >= 0) continue;
    if (upNama === 'SEKOLAH' || upNama.indexOf('SEKOLAH (') === 0) continue;

    var items = [];
    var totalTagihan = 0;
    var totalBayar = 0;
    for (var a = 0; a < activities.length; a++) {
      var act = activities[a];
      var t = parseMoney_(row[act.tagihanCol]);
      var b = act.bayarCol >= 0 ? parseMoney_(row[act.bayarCol]) : 0;
      var kurang = Math.max(t - b, 0);
      totalTagihan += t;
      totalBayar += b;
      items.push({
        nama: act.name,
        tagihan: t,
        bayar: b,
        kurang: kurang,
        status: itemStatus_(t, b)
      });
    }
    schools.push({
      no: no,
      nama: nama,
      rowIndex: r + 1,
      items: items,
      totalTagihan: totalTagihan,
      totalBayar: totalBayar,
      totalKurang: Math.max(totalTagihan - totalBayar, 0)
    });
  }

  return { sheet: sh, values: values, activities: activities, trow: trow, schools: schools, totalCols: totalCols };
}

function visibleItems_(items, includeZero) {
  if (includeZero) return items;
  var out = [];
  for (var i = 0; i < items.length; i++) {
    if ((items[i].tagihan || 0) > 0 || (items[i].bayar || 0) > 0) out.push(items[i]);
  }
  return out;
}

function readKewajiban_(includeZero) {
  var layout = readLayout_();
  var schools = layout.schools.map(function (s) {
    var items = visibleItems_(s.items, !!includeZero);
    return {
      no: s.no,
      nama: s.nama,
      rowIndex: s.rowIndex,
      items: items,
      totalTagihan: s.totalTagihan,
      totalBayar: s.totalBayar,
      totalKurang: s.totalKurang
    };
  });
  return {
    activities: layout.activities.map(function (a) { return a.name; }),
    schools: schools
  };
}

/* ========== API ADMIN ========== */
function getAdminDashboard(token) {
  requireAdmin_(token);
  var data = readKewajiban_();
  var sumT = 0, sumB = 0, sumK = 0;
  var list = data.schools.map(function (s) {
    sumT += s.totalTagihan;
    sumB += s.totalBayar;
    sumK += s.totalKurang;
    return {
      no: s.no,
      nama: s.nama,
      totalTagihan: s.totalTagihan,
      totalBayar: s.totalBayar,
      totalKurang: s.totalKurang,
      jumlahItem: s.items.length,
      belumLunas: s.items.filter(function (i) { return i.status !== 'Lunas'; }).length
    };
  });
  return {
    ok: true,
    org: ORG_NAME,
    wilayah: ORG_WILAYAH,
    summary: { sekolah: list.length, totalTagihan: sumT, totalBayar: sumB, totalKurang: sumK },
    sekolah: list,
    pengumuman: getPengumumanText_(),
    activities: data.activities
  };
}

function getAdminSekolahDetail(token, namaSekolah) {
  requireAdmin_(token);
  var data = readKewajiban_(true);
  var nama = clean_(namaSekolah).toLowerCase();
  for (var i = 0; i < data.schools.length; i++) {
    if (clean_(data.schools[i].nama).toLowerCase() === nama) {
      return { ok: true, sekolah: data.schools[i] };
    }
  }
  return { ok: false, message: 'Sekolah tidak ditemukan.' };
}

function saveSchoolNominals(token, namaSekolah, items) {
  requireAdmin_(token);
  var nama = clean_(namaSekolah).toLowerCase();
  if (!nama) throw new Error('Nama sekolah wajib.');
  items = items || [];

  var layout = readLayout_();
  var school = null;
  for (var i = 0; i < layout.schools.length; i++) {
    if (clean_(layout.schools[i].nama).toLowerCase() === nama) {
      school = layout.schools[i];
      break;
    }
  }
  if (!school) throw new Error('Sekolah tidak ditemukan di KewajibanSD2026.');

  var byName = {};
  for (var a = 0; a < layout.activities.length; a++) {
    byName[layout.activities[a].name.toLowerCase()] = layout.activities[a];
  }

  var sh = layout.sheet;
  var row = school.rowIndex;
  var changed = 0;
  for (var n = 0; n < items.length; n++) {
    var it = items[n] || {};
    var act = byName[clean_(it.nama).toLowerCase()];
    if (!act) continue;
    var tagihan = parseMoney_(it.tagihan);
    var bayar = parseMoney_(it.bayar);
    sh.getRange(row, act.tagihanCol + 1).setValue(tagihan);
    if (act.bayarCol >= 0) sh.getRange(row, act.bayarCol + 1).setValue(bayar);
    changed++;
  }

  // hitung ulang total baris dari seluruh kegiatan
  var fresh = readLayout_();
  var again = null;
  for (var j = 0; j < fresh.schools.length; j++) {
    if (clean_(fresh.schools[j].nama).toLowerCase() === nama) { again = fresh.schools[j]; break; }
  }
  if (again && fresh.totalCols) {
    if (fresh.totalCols.tagihan >= 0) sh.getRange(again.rowIndex, fresh.totalCols.tagihan + 1).setValue(again.totalTagihan);
    if (fresh.totalCols.bayar >= 0) sh.getRange(again.rowIndex, fresh.totalCols.bayar + 1).setValue(again.totalBayar);
    if (fresh.totalCols.kurang >= 0) sh.getRange(again.rowIndex, fresh.totalCols.kurang + 1).setValue(again.totalKurang);
    // baris JUMLAH MASUK
    var vals = sh.getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      if (clean_(vals[r][1]).toUpperCase().indexOf('JUMLAH') >= 0) {
        var sumT = 0, sumB = 0;
        for (var s = 0; s < fresh.schools.length; s++) {
          sumT += fresh.schools[s].totalTagihan;
          sumB += fresh.schools[s].totalBayar;
        }
        if (fresh.totalCols.tagihan >= 0) sh.getRange(r + 1, fresh.totalCols.tagihan + 1).setValue(sumT);
        if (fresh.totalCols.bayar >= 0) sh.getRange(r + 1, fresh.totalCols.bayar + 1).setValue(sumB);
        if (fresh.totalCols.kurang >= 0) sh.getRange(r + 1, fresh.totalCols.kurang + 1).setValue(Math.max(sumT - sumB, 0));
        break;
      }
    }
  }

  SpreadsheetApp.flush();
  log_('ADMIN', 'Update Nominal', school.nama + ' (' + changed + ' kegiatan)');
  return { ok: true, message: 'Nominal ' + school.nama + ' disimpan ke sheet KewajibanSD2026.' };
}

function getPengumumanText_() {
  try {
    var sh = ensurePengumumanSheet_();
    var last = sh.getLastRow();
    if (last < 2) return '';
    // ambil pengumuman aktif terakhir
    var vals = sh.getRange(2, 1, last, 3).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var aktif = clean_(vals[i][2]).toLowerCase();
      if (aktif === 'ya' || aktif === 'aktif' || aktif === '') {
        return clean_(vals[i][1]);
      }
    }
    return clean_(vals[vals.length - 1][1]);
  } catch (e) {
    return '';
  }
}

function savePengumuman(token, teks) {
  requireAdmin_(token);
  var isi = clean_(teks);
  if (!isi) throw new Error('Isi pengumuman tidak boleh kosong.');
  var sh = ensurePengumumanSheet_();
  // nonaktifkan yang lama
  var last = sh.getLastRow();
  if (last >= 2) {
    var col = sh.getRange(2, 3, last, 3);
    var vals = col.getValues();
    for (var i = 0; i < vals.length; i++) vals[i][0] = 'Tidak';
    col.setValues(vals);
  }
  sh.appendRow([stamp_(), isi, 'Ya']);
  log_('ADMIN', 'Update Pengumuman', isi.substring(0, 80));
  return { ok: true, message: 'Pengumuman disimpan dan akan tampil di dashboard sekolah.', teks: isi };
}

function getAksesList(token) {
  requireAdmin_(token);
  ensureAksesSheet_();
  var sh = sh_(SHEET_AKSES);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var vals = sh.getRange(2, 1, last, 4).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    rows.push({
      no: vals[i][0],
      nama: clean_(vals[i][1]),
      password: clean_(vals[i][2]),
      aktif: clean_(vals[i][3]) || 'Ya'
    });
  }
  return { ok: true, rows: rows };
}

function updatePasswordSekolah(token, namaSekolah, passwordBaru) {
  requireAdmin_(token);
  var nama = clean_(namaSekolah);
  var pass = clean_(passwordBaru);
  if (!nama || !pass) throw new Error('Nama sekolah dan password wajib.');
  var sh = sh_(SHEET_AKSES);
  var last = sh.getLastRow();
  var vals = sh.getRange(2, 1, last, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (clean_(vals[i][1]).toLowerCase() === nama.toLowerCase()) {
      sh.getRange(i + 2, 3).setValue(pass);
      log_('ADMIN', 'Ubah Password', nama);
      return { ok: true, message: 'Password ' + nama + ' diperbarui.' };
    }
  }
  throw new Error('Sekolah tidak ditemukan di sheet Akses.');
}

/* ========== API VISITOR ========== */
function getDiagnosa_() {
  try {
    var sh = ss_().getSheetByName(SHEET_DATA);
    if (!sh) return 'Sheet KewajibanSD2026 tidak ditemukan. Script harus terpasang pada spreadsheet yang sama.';
    var values = sh.getDataRange().getDisplayValues();
    var trow = findTagihanRow_(values);
    var layout = readLayout_();
    var contoh = layout.schools.length
      ? (layout.schools[0].nama + ' = Rp ' + layout.schools[0].totalTagihan)
      : 'tidak ada baris sekolah';
    return 'Baris label TAGIHAN: ' + (trow < 0 ? 'tidak ketemu' : String(trow + 1))
      + ' · kegiatan: ' + layout.activities.length
      + ' · sekolah: ' + layout.schools.length
      + ' · contoh: ' + contoh;
  } catch (e) {
    return 'Gagal membaca sheet: ' + e.message;
  }
}

function getVisitorDashboard(token) {
  var t = requireVisitor_(token);
  var data = readKewajiban_();
  var target = null;
  for (var i = 0; i < data.schools.length; i++) {
    if (clean_(data.schools[i].nama).toLowerCase() === clean_(t.id).toLowerCase()) {
      target = data.schools[i];
      break;
    }
  }
  if (!target) throw new Error('Data sekolah tidak ditemukan di KewajibanSD2026.');

  return {
    ok: true,
    org: ORG_NAME,
    wilayah: ORG_WILAYAH,
    pengumuman: getPengumumanText_(),
    diagnosa: getDiagnosa_(),
    sekolah: {
      no: target.no,
      nama: target.nama,
      items: target.items,
      totalTagihan: target.totalTagihan,
      totalBayar: target.totalBayar,
      totalKurang: target.totalKurang
    }
  };
}

/* ========== CETAK ========== */
function getPrintData(token, namaSekolahOpt) {
  var t = parseToken_(token);
  if (!t) throw new Error('Sesi tidak valid.');
  var data = readKewajiban_();
  var nama = t.role === 'admin' ? clean_(namaSekolahOpt) : clean_(t.id);
  var target = null;
  for (var i = 0; i < data.schools.length; i++) {
    if (clean_(data.schools[i].nama).toLowerCase() === nama.toLowerCase()) {
      target = data.schools[i];
      break;
    }
  }
  if (!target) throw new Error('Sekolah tidak ditemukan.');
  var set = getPengaturan_();
  return {
    ok: true,
    org: set.NamaOrganisasi || ORG_NAME,
    wilayah: set.Wilayah || ORG_WILAYAH,
    ketua: set.Ketua || 'Ketua KKKS',
    bendahara: set.Bendahara || 'Bendahara KKKS',
    tanggal: today_(),
    sekolah: target
  };
}
