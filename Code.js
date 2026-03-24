// ============================================================
//  MANAJEMEN KOS — Code.gs  (FIXED VERSION)
// ============================================================

function doGet(e) {
  try {
    var page = 'login';
    if (e && e.parameter) page = e.parameter._page || e.parameter.page || 'login';
    var allowed = ['login','dashboard','kamar','penyewa','pembayaran','laporan','profil'];
    if (allowed.indexOf(page) === -1) page = 'login';
    return HtmlService.createHtmlOutputFromFile(page)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Manajemen Kos');
  } catch(err) {
    return HtmlService.createHtmlOutput('<h3 style="font-family:sans-serif;color:red">Error: ' + err.message + '</h3>');
  }
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

// Cache spreadsheet dalam satu request agar tidak buka ulang tiap pemanggilan
var _ss = null;
function ss() {
  if (!_ss) _ss = SpreadsheetApp.openById('19m4h-0WhKpC9U2B1rcJI1avKslcFZO0wuRWx0-jQq7U');
  return _ss;
}
// Cache sheet per nama
var _shCache = {};
function sh(name) {
  if (!_shCache[name]) _shCache[name] = ss().getSheetByName(name);
  return _shCache[name];
}
function fmtRp(n) { return Number(n).toLocaleString('id-ID'); }
function tz() { return Session.getScriptTimeZone(); }
function fmtDate(d) { return Utilities.formatDate(new Date(d), tz(), 'dd/MM/yyyy'); }
function nowDate() { return Utilities.formatDate(new Date(), tz(), 'dd/MM/yyyy'); }
function toInt(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }  // ← helper baru

// ===== AUTH =====
function login(username, password) {
  var data = sh('users').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === username && String(data[i][1]) === password) {
      return { status: true, role: String(data[i][2]), username: String(data[i][0]), nama: String(data[i][3]) };
    }
  }
  return { status: false };
}

// ===== PROFIL KOS =====
function getProfil() {
  var sheet = sh('profil');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) result[String(data[i][0])] = String(data[i][1] || '');
  }
  return result;
}

function saveProfil(d) {
  var sheet = sh('profil');
  if (!sheet) sheet = ss().insertSheet('profil');
  var data = sheet.getDataRange().getValues();
  var keys = Object.keys(d);
  keys.forEach(function(key) {
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(d[key]);
        found = true; break;
      }
    }
    if (!found) { sheet.appendRow([key, d[key]]); data.push([key, d[key]]); }
  });
  return { status: true };
}

// ===== DASHBOARD SUMMARY =====
function getDashboardSummary() {
  var kamarData   = sh('kamar').getDataRange().getValues();
  var penyewaData = sh('penyewa').getDataRange().getValues();
  var bayarData   = sh('pembayaran').getDataRange().getValues();

  var totalKamar = 0, kamarKosong = 0, kamarTerisi = 0, kamarMaint = 0;
  for (var i = 1; i < kamarData.length; i++) {
    totalKamar++;
    var st = String(kamarData[i][4]).toLowerCase();
    if (st === 'kosong') kamarKosong++;
    else if (st === 'terisi') kamarTerisi++;
    else kamarMaint++;
  }

  var totalPenyewa = 0;
  for (var i = 1; i < penyewaData.length; i++) {
    if (String(penyewaData[i][9]).toLowerCase() === 'aktif') totalPenyewa++;
  }

  var now = new Date(), bulanIni = now.getMonth() + 1, tahunIni = now.getFullYear();
  var pendapatanBulanIni = 0, tagiBelum = 0;
  for (var i = 1; i < bayarData.length; i++) {
    if (Number(bayarData[i][3]) === bulanIni && Number(bayarData[i][4]) === tahunIni) {
      if (String(bayarData[i][9]).toLowerCase() === 'lunas') pendapatanBulanIni += Number(bayarData[i][5]);
      else tagiBelum++;
    }
  }

  var jatuhTempo = [];
  for (var i = 1; i < penyewaData.length; i++) {
    if (String(penyewaData[i][9]).toLowerCase() !== 'aktif') continue;
    var pid = toInt(penyewaData[i][0]), sudahBayar = false;
    for (var j = 1; j < bayarData.length; j++) {
      if (toInt(bayarData[j][1]) === pid && Number(bayarData[j][3]) === bulanIni && Number(bayarData[j][4]) === tahunIni && String(bayarData[j][9]).toLowerCase() === 'lunas') { sudahBayar = true; break; }
    }
    if (!sudahBayar) {
      var nomorKamar = '-';
      var kamarIdCari = toInt(penyewaData[i][6]);
      for (var k = 1; k < kamarData.length; k++) { if (toInt(kamarData[k][0]) === kamarIdCari) { nomorKamar = kamarData[k][1]; break; } }
      jatuhTempo.push({ nama: penyewaData[i][1], kamar: nomorKamar, hp: penyewaData[i][3] });
    }
  }

  return { totalKamar: totalKamar, kamarKosong: kamarKosong, kamarTerisi: kamarTerisi, kamarMaint: kamarMaint, totalPenyewa: totalPenyewa, pendapatanBulanIni: pendapatanBulanIni, tagiBelum: tagiBelum, jatuhTempo: jatuhTempo, bulan: bulanIni, tahun: tahunIni };
}

// ===== KAMAR =====
function getKamar() {
  try {
    var kSheet = sh('kamar');
    var pSheet = sh('penyewa');

    // Ambil hanya kolom yang dibutuhkan, bukan getDataRange() yg baca semua kolom
    var kLastRow = kSheet.getLastRow();
    var pLastRow = pSheet.getLastRow();

    if (kLastRow < 2) return [];

    // Kamar: ambil 8 kolom saja
    var kRaw = kSheet.getRange(1, 1, kLastRow, 8).getValues();

    // Build map penghuni aktif dari kolom yg dibutuhkan saja (id,nama,hp,kamar_id,status)
    var penghuniMap = {};
    if (pLastRow >= 2) {
      // Ambil kolom: 1=id, 2=nama, 4=no_hp, 7=kamar_id, 10=status_aktif
      var pRaw = pSheet.getRange(1, 1, pLastRow, 10).getValues();
      for (var i = 1; i < pRaw.length; i++) {
        var statusAktif = pRaw[i][9];
        if (statusAktif === null || statusAktif === undefined) continue;
        if (String(statusAktif).toLowerCase() !== 'aktif') continue;
        var kid = toInt(pRaw[i][6]);
        if (!kid) continue;
        penghuniMap[kid] = {
          nama:       String(pRaw[i][1] || ''),
          no_hp:      String(pRaw[i][3] || ''),
          penyewa_id: toInt(pRaw[i][0])
        };
      }
    }

    // Konversi setiap cell ke primitif — tidak boleh ada Date/null
    var result = [];
    for (var i = 1; i < kRaw.length; i++) {
      var r = kRaw[i];
      var kid = toInt(r[0]);
      var p = penghuniMap[kid] || null;
      result.push({
        id:            kid,
        nomor:         r[1] !== null && r[1] !== undefined ? String(r[1]) : '',
        tipe:          r[2] !== null && r[2] !== undefined ? String(r[2]) : '',
        harga:         r[3] !== null && r[3] !== undefined ? Number(r[3]) : 0,
        status:        r[4] !== null && r[4] !== undefined ? String(r[4]) : 'kosong',
        fasilitas:     r[5] !== null && r[5] !== undefined ? String(r[5]) : '',
        keterangan:    r[6] !== null && r[6] !== undefined ? String(r[6]) : '',
        gambar_url:    r[7] !== null && r[7] !== undefined ? String(r[7]) : '',
        penghuni_nama: p ? p.nama : '',
        penghuni_hp:   p ? p.no_hp : '',
        penghuni_id:   p ? p.penyewa_id : 0
      });
    }
    return result;

  } catch(err) {
    return { error: err.message };
  }
}

// ===== DEBUG: test apakah getKamar bisa jalan =====
function getKamarDebug() {
  try {
    var kSheet = sh('kamar');
    var kLastRow = kSheet.getLastRow();
    return { ok: true, lastRow: kLastRow, msg: 'sheet kamar ditemukan' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

// Test serialisasi: kembalikan data mentah kamar tanpa penyewa
function getKamarRaw() {
  try {
    var kSheet = sh('kamar');
    var kLastRow = kSheet.getLastRow();
    if (kLastRow < 2) return { ok: true, rows: [] };
    var raw = kSheet.getRange(1, 1, kLastRow, 8).getValues();
    // Konversi PAKSA semua cell ke string/number
    var clean = [];
    for (var i = 1; i < raw.length; i++) {
      var r = raw[i];
      var row = [];
      for (var j = 0; j < r.length; j++) {
        var v = r[j];
        if (v === null || v === undefined) { row.push(''); }
        else if (v instanceof Date) { row.push(Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')); }
        else { row.push(String(v)); }
      }
      clean.push(row);
    }
    return { ok: true, rows: clean };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function tambahKamar(d) {
  var sheet = sh('kamar'), data = sheet.getDataRange().getValues(), newId = 1;
  for (var i = 1; i < data.length; i++) if (toInt(data[i][0]) >= newId) newId = toInt(data[i][0]) + 1;
  sheet.appendRow([newId, d.nomor, d.tipe, Number(d.harga), d.status||'kosong', d.fasilitas||'', d.keterangan||'', d.gambar_url||'']);
  return { status: true, id: newId };
}

function editKamar(d) {
  var sheet = sh('kamar'), data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(d.id)) {
      sheet.getRange(i+1,2).setValue(d.nomor);
      sheet.getRange(i+1,3).setValue(d.tipe);
      sheet.getRange(i+1,4).setValue(Number(d.harga));
      sheet.getRange(i+1,5).setValue(d.status);
      sheet.getRange(i+1,6).setValue(d.fasilitas||'');
      sheet.getRange(i+1,7).setValue(d.keterangan||'');
      sheet.getRange(i+1,8).setValue(d.gambar_url||'');
      return { status: true };
    }
  }
  return { status: false, msg: 'Kamar tidak ditemukan' };
}

function hapusKamar(id) {
  var sheet = sh('kamar'), data = sheet.getDataRange().getValues();
  var pData = sh('penyewa').getDataRange().getValues();
  for (var j = 1; j < pData.length; j++) {
    if (toInt(pData[j][6]) === toInt(id) && String(pData[j][9]).toLowerCase() === 'aktif')
      return { status: false, msg: 'Kamar masih dihuni penyewa aktif!' };
  }
  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(id)) { sheet.deleteRow(i+1); return { status: true }; }
  }
  return { status: false };
}

function uploadGambarKamar(base64Data, fileName, mimeType) {
  try {
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);

    var folderName = 'Gambar Kamar Kos';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Utilities.sleep(500); // penting!

    return { 
      status: true, 
      url: 'https://lh3.googleusercontent.com/d/' + file.getId()
    };

  } catch(e) {
    return { status: false, msg: e.message };
  }
}

// ===== PENYEWA =====
function getPenyewa() {
  try {
    var pSheet = sh('penyewa');
    var kSheet = sh('kamar');

    var pLastRow = pSheet.getLastRow();
    var kLastRow = kSheet.getLastRow();

    if (pLastRow < 2) return [];
    if (kLastRow < 2) return { error: 'Data kamar kosong' };

    var pData = pSheet.getRange(1, 1, pLastRow, 10).getValues();
    var kData = kSheet.getRange(1, 1, kLastRow, 8).getValues();

    // Build kamar map: id -> {nomor, harga}
    var kamarMap = {};
    for (var i = 1; i < kData.length; i++) {
      var kid = toInt(kData[i][0]);
      if (kid) kamarMap[kid] = { nomor: String(kData[i][1] || '-'), harga: Number(kData[i][3] || 0) };
    }

    // ✅ FIX: return array of OBJECTS (bukan array of arrays)
    // GAS tidak bisa serialize Date/null dalam nested array → return null
    // Solusi: konversi tiap row ke plain object dengan semua nilai string/number
    var result = [];
    for (var i = 1; i < pData.length; i++) {
      var r = pData[i];
      var kamarId = toInt(r[6]);
      var km = kamarMap[kamarId] || { nomor: '-', harga: 0 };
      result.push({
        id:          toInt(r[0]),
        nama:        String(r[1] || ''),
        nik:         String(r[2] || ''),
        no_hp:       String(r[3] || ''),
        email:       String(r[4] || ''),
        alamat_asal: String(r[5] || ''),
        kamar_id:    kamarId,
        tgl_masuk:   r[7] ? Utilities.formatDate(new Date(r[7]), tz(), 'yyyy-MM-dd') : '',
        tgl_keluar:  r[8] ? Utilities.formatDate(new Date(r[8]), tz(), 'yyyy-MM-dd') : '',
        status_aktif: String(r[9] || ''),
        nomor_kamar: km.nomor,
        harga_kamar: km.harga
      });
    }
    return result;

  } catch (err) {
    return { error: err.message };
  }
}

function getKamarKosong() {
  var data = sh('kamar').getDataRange().getValues(), result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).toLowerCase() === 'kosong')
      result.push({ id: toInt(data[i][0]), nomor: data[i][1], tipe: data[i][2], harga: Number(data[i][3]) });
  }
  return result;
}

function tambahPenyewa(d) {
  var sheet = sh('penyewa'), kamarSheet = sh('kamar');
  var data = sheet.getDataRange().getValues(), newId = 1;
  for (var i = 1; i < data.length; i++) if (toInt(data[i][0]) >= newId) newId = toInt(data[i][0]) + 1;
  sheet.appendRow([newId, d.nama, d.nik, d.no_hp, d.email||'', d.alamat_asal||'', toInt(d.kamar_id), d.tgl_masuk, d.tgl_keluar||'', 'aktif']);
  var kData = kamarSheet.getDataRange().getValues();
  for (var i = 1; i < kData.length; i++) {
    if (toInt(kData[i][0]) === toInt(d.kamar_id)) { kamarSheet.getRange(i+1,5).setValue('terisi'); break; }
  }
  return { status: true, id: newId };
}

function editPenyewa(d) {
  var sheet = sh('penyewa'), data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(d.id)) {
      sheet.getRange(i+1,2).setValue(d.nama);
      sheet.getRange(i+1,3).setValue(d.nik);
      sheet.getRange(i+1,4).setValue(d.no_hp);
      sheet.getRange(i+1,5).setValue(d.email||'');
      sheet.getRange(i+1,6).setValue(d.alamat_asal||'');
      sheet.getRange(i+1,8).setValue(d.tgl_masuk);
      sheet.getRange(i+1,9).setValue(d.tgl_keluar||'');
      return { status: true };
    }
  }
  return { status: false };
}

function penyewaKeluar(id) {
  var sheet = sh('penyewa'), data = sheet.getDataRange().getValues();
  var kamarSheet = sh('kamar'), kData = kamarSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(id)) {
      sheet.getRange(i+1,10).setValue('keluar');
      sheet.getRange(i+1,9).setValue(Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'));
      var kamarId = toInt(data[i][6]);
      for (var j = 1; j < kData.length; j++) {
        if (toInt(kData[j][0]) === kamarId) { kamarSheet.getRange(j+1,5).setValue('kosong'); break; }
      }
      return { status: true };
    }
  }
  return { status: false };
}

function hapusPenyewa(id) {
  var sheet = sh('penyewa');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(id)) {

      // ❌ CEK: hanya boleh hapus kalau status = keluar
      if (String(data[i][9]).toLowerCase() !== 'keluar') {
        return { status: false, msg: "Hanya penyewa yang sudah keluar yang bisa dihapus!" };
      }

      // ✅ HAPUS DATA
      sheet.deleteRow(i + 1);

      return { status: true };
    }
  }

  return { status: false, msg: "Data tidak ditemukan" };
}

// ===== PEMBAYARAN =====
function getPembayaran(filter) {
  var bayarData = sh('pembayaran').getDataRange().getValues();
  var pData = sh('penyewa').getDataRange().getValues();
  var kData = sh('kamar').getDataRange().getValues();
  var pMap = {}, kMap = {};
  for (var i = 1; i < pData.length; i++) pMap[toInt(pData[i][0])] = pData[i][1];
  for (var i = 1; i < kData.length; i++) kMap[toInt(kData[i][0])] = kData[i][1];
  var result = [];
  for (var i = 1; i < bayarData.length; i++) {
    var row = bayarData[i];
    if (filter && filter.bulan && Number(row[3]) !== Number(filter.bulan)) continue;
    if (filter && filter.tahun && Number(row[4]) !== Number(filter.tahun)) continue;
    result.push({
      id: row[0], penyewa_id: row[1], kamar_id: row[2],
      bulan: row[3], tahun: row[4], jumlah: row[5],
      tgl_bayar: row[6] ? fmtDate(row[6]) : '-',
      metode: row[7]||'-', keterangan: row[8]||'', status: row[9],
      nama_penyewa: pMap[toInt(row[1])] || '-',
      nomor_kamar:  kMap[toInt(row[2])] || '-'
    });
  }
  return result.reverse();
}

function catatPembayaran(d) {
  var sheet = sh('pembayaran'), data = sheet.getDataRange().getValues(), newId = 1;
  for (var i = 1; i < data.length; i++) if (toInt(data[i][0]) >= newId) newId = toInt(data[i][0]) + 1;
  sheet.appendRow([newId, toInt(d.penyewa_id), toInt(d.kamar_id), Number(d.bulan), Number(d.tahun), Number(d.jumlah), new Date(), d.metode||'Tunai', d.keterangan||'', 'lunas']);
  return { status: true, id: newId };
}

function hapusPembayaran(id) {
  var sheet = sh('pembayaran'), data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (toInt(data[i][0]) === toInt(id)) { sheet.deleteRow(i+1); return { status: true }; }
  }
  return { status: false };
}

function buatKwitansi(d) {
  var profil = getProfil();
  var namaKos = profil.nama_kos || 'Manajemen Kos';
  var alamat  = profil.alamat   || '';
  var noWa    = profil.no_wa    || '';
  var bulanNama = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

  var html = '<!DOCTYPE html><html><body style="font-family:Arial;max-width:420px;margin:auto;padding:30px;">' +
    '<div style="text-align:center;border-bottom:2px solid #2d5be3;padding-bottom:14px;margin-bottom:18px;">' +
    '<h2 style="margin:0;font-size:20px;color:#2d5be3;">'+namaKos+'</h2>' +
    (alamat?'<p style="margin:3px 0;font-size:11px;color:#666;">'+alamat+'</p>':'') +
    (noWa?'<p style="margin:3px 0;font-size:11px;color:#666;">WA: '+noWa+'</p>':'') +
    '<p style="margin:8px 0 0;font-size:14px;font-weight:bold;color:#333;">KWITANSI PEMBAYARAN SEWA</p></div>' +
    '<table width="100%" style="border-collapse:collapse;font-size:13px;">' +
    '<tr><td style="padding:5px 0;color:#666;width:140px">No. Kwitansi</td><td>: <b>#'+String(d.id).padStart(5,'0')+'</b></td></tr>' +
    '<tr><td style="padding:5px 0;color:#666">Tanggal</td><td>: '+nowDate()+'</td></tr>' +
    '<tr><td style="padding:5px 0;color:#666">Nama Penyewa</td><td>: <b>'+d.nama_penyewa+'</b></td></tr>' +
    '<tr><td style="padding:5px 0;color:#666">Nomor Kamar</td><td>: '+d.nomor_kamar+'</td></tr>' +
    '<tr><td style="padding:5px 0;color:#666">Periode Sewa</td><td>: '+bulanNama[Number(d.bulan)-1]+' '+d.tahun+'</td></tr>' +
    '<tr><td style="padding:5px 0;color:#666">Metode Bayar</td><td>: '+d.metode+'</td></tr>' +
    '</table>' +
    '<div style="background:#eef2ff;border:1px dashed #2d5be3;border-radius:10px;padding:16px;margin:18px 0;text-align:center;">' +
    '<p style="margin:0;font-size:11px;color:#666;text-transform:uppercase;letter-spacing:1px;">Jumlah Dibayar</p>' +
    '<h1 style="margin:8px 0;font-size:30px;color:#2d5be3;font-weight:bold;">Rp '+fmtRp(d.jumlah)+'</h1>' +
    '<p style="margin:0;font-size:12px;color:#059669;font-weight:bold;">✓ LUNAS</p>' +
    '</div>' +
    (d.keterangan?'<p style="font-size:12px;color:#666;margin-bottom:14px;">Keterangan: '+d.keterangan+'</p>':'') +
    '<div style="border-top:1px solid #eee;margin-top:20px;padding-top:14px;display:flex;justify-content:space-between;align-items:flex-end;">' +
    '<span style="font-size:11px;color:#999;">Dicetak: '+nowDate()+'</span>' +
    '<div style="text-align:center;"><div style="width:90px;height:55px;"></div><div style="border-top:1px solid #333;font-size:11px;padding-top:4px;">Petugas / Pemilik</div></div>' +
    '</div></body></html>';

  var blob = Utilities.newBlob(html,'text/html').getAs('application/pdf');
  blob.setName('Kwitansi-'+d.id+'-'+d.nama_penyewa+'.pdf');
  var file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=download&id=' + file.getId();
}

function getPenyewaAktif() {
  var pData = sh('penyewa').getDataRange().getValues();
  var kData = sh('kamar').getDataRange().getValues();
  var kMap  = {};
  for (var i = 1; i < kData.length; i++) kMap[toInt(kData[i][0])] = { nomor: kData[i][1], harga: Number(kData[i][3]) };
  var result = [];
  for (var i = 1; i < pData.length; i++) {
    if (String(pData[i][9]).toLowerCase() !== 'aktif') continue;
    var km = kMap[toInt(pData[i][6])] || {};
    result.push({ id: pData[i][0], nama: pData[i][1], kamar_id: pData[i][6], nomor_kamar: km.nomor||'-', harga: km.harga||0 });
  }
  return result;
}

// ===== LAPORAN =====
function getLaporan(mode) {
  var bayarData = sh('pembayaran').getDataRange().getValues();
  var kData     = sh('kamar').getDataRange().getValues();
  var pData     = sh('penyewa').getDataRange().getValues();
  var grouped = {}, totalPend = 0, totalTrx = 0;
  for (var i = 1; i < bayarData.length; i++) {
    if (String(bayarData[i][9]).toLowerCase() !== 'lunas') continue;
    var bulan = Number(bayarData[i][3]), tahun = Number(bayarData[i][4]);
    var key = (mode==='tahunan') ? String(tahun) : String(tahun)+'-'+String(bulan).padStart(2,'0');
    if (!grouped[key]) grouped[key] = { total: 0, count: 0 };
    grouped[key].total += Number(bayarData[i][5]);
    grouped[key].count++;
    totalPend += Number(bayarData[i][5]);
    totalTrx++;
  }
  var tipeMap = {};
  for (var i = 1; i < kData.length; i++) tipeMap[toInt(kData[i][0])] = kData[i][2];
  var tipeSales = {};
  for (var i = 1; i < bayarData.length; i++) {
    if (String(bayarData[i][9]).toLowerCase() !== 'lunas') continue;
    var tipe = tipeMap[toInt(bayarData[i][2])]||'Lainnya';
    if (!tipeSales[tipe]) tipeSales[tipe] = 0;
    tipeSales[tipe] += Number(bayarData[i][5]);
  }
  var labels = Object.keys(grouped).sort();
  var bulanNama = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
  var labelsFmt = labels.map(function(l){ if(mode==='tahunan') return l; var p=l.split('-'); return bulanNama[parseInt(p[1])-1]+' '+p[0]; });
  var totalKamar = kData.length-1, kamarTerisi = 0, penyewaAktif = 0;
  for (var i = 1; i < kData.length; i++) if (String(kData[i][4]).toLowerCase()==='terisi') kamarTerisi++;
  for (var i = 1; i < pData.length; i++) if (String(pData[i][9]).toLowerCase()==='aktif') penyewaAktif++;
  return { labels: labelsFmt, totals: labels.map(function(k){ return grouped[k].total; }), counts: labels.map(function(k){ return grouped[k].count; }), totalPendapatan: totalPend, totalTransaksi: totalTrx, tipeSales: tipeSales, totalKamar: totalKamar, kamarTerisi: kamarTerisi, penyewaAktif: penyewaAktif, okupansi: totalKamar?Math.round(kamarTerisi/totalKamar*100):0 };
}