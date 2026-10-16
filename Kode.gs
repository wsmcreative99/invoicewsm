/**
 * BACKEND INVOICE WSM (CARA B) — project Apps Script berdiri sendiri
 * ----------------------------------------------------------------
 * Script ini membuka spreadsheet lewat ID (SHEET_ID), jadi tidak harus
 * dibuat dari menu Ekstensi di dalam spreadsheet.
 *
 * Alur: tombol "Cetak" di invoice -> kirim data ke Web App ini
 *       -> tersimpan di sheet "Riwayat" -> Dashboard & Rekap Bulanan
 *       dihitung otomatis oleh rumus.
 *
 * TOKEN dan SHEET_ID sudah terisi. Samakan TOKEN dengan SHEETS_TOKEN di index.html.
 * Urutan: Simpan > jalankan setup() > jalankan tesKirim() > Deploy.
 * Setelah kode ini diubah: Terapkan > Kelola deployment > pensil > Versi baru > Terapkan.
 */

const TOKEN = 'wsm-4444';

// ID spreadsheet = bagian URL di antara /d/ dan /edit
const SHEET_ID = '144O7oSWjr764j-AUaZoRPMAJlFOTINkmF1iK-ME9HdM';

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

const SH_RIWAYAT = 'Riwayat';
const SH_ITEM = 'Detail Item';
const SH_BULANAN = 'Rekap Bulanan';
const SH_DASH = 'Dashboard';

const RP_FORMAT = '"Rp"#,##0';

/* ===================== SETUP (jalankan sekali) ===================== */
function setup() {
  const ss = ss_();
  // Locale harus en_US: rumus yang ditulis script memakai pemisah koma.
  // Dengan locale id_ID rumus bertanda koma jadi #ERROR!.
  ss.setSpreadsheetLocale('en_US');
  ss.setSpreadsheetTimeZone('Asia/Jakarta');

  // ---- Riwayat ----
  const rw = getOrCreate_(ss, SH_RIWAYAT);
  rw.getRange(1, 1, 1, 11).setValues([[
    'Tanggal Cetak', 'No Invoice', 'Klien', 'Tanggal Acara', 'Lokasi',
    'Subtotal', 'Diskon', 'Pajak / Biaya', 'Total Akhir', 'Status', 'Bulan'
  ]]);
  styleHeader_(rw.getRange(1, 1, 1, 11));
  rw.setFrozenRows(1);
  rw.getRange('A2:A').setNumberFormat('dd/mm/yyyy hh:mm');
  rw.getRange('D2:D').setNumberFormat('@');
  rw.getRange('F2:I').setNumberFormat(RP_FORMAT);
  rw.getRange('K2:K').setNumberFormat('MMMM yyyy');
  rw.getRange('J2:J').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Belum Dibayar', 'Lunas'], true)
      .setAllowInvalid(false).build()
  );
  rw.setColumnWidths(1, 11, 130);
  rw.setColumnWidth(3, 200);
  rw.setColumnWidth(5, 200);

  // ---- Detail Item ----
  const it = getOrCreate_(ss, SH_ITEM);
  it.getRange(1, 1, 1, 9).setValues([[
    'No Invoice', 'Tanggal Cetak', 'Klien', 'Uraian Item', 'Qty',
    'Satuan', 'Harga Satuan', 'Total', 'Keterangan'
  ]]);
  styleHeader_(it.getRange(1, 1, 1, 9));
  it.setFrozenRows(1);
  it.getRange('B2:B').setNumberFormat('dd/mm/yyyy hh:mm');
  it.getRange('G2:H').setNumberFormat(RP_FORMAT);
  it.setColumnWidths(1, 9, 130);
  it.setColumnWidth(4, 220);

  // ---- Rekap Bulanan (otomatis lewat rumus) ----
  const bl = getOrCreate_(ss, SH_BULANAN);
  bl.clear();
  bl.getRange(1, 1, 1, 5).setValues([[
    'Bulan', 'Jumlah Invoice', 'Total Tagihan', 'Sudah Lunas', 'Belum Dibayar'
  ]]);
  styleHeader_(bl.getRange(1, 1, 1, 5));
  bl.setFrozenRows(1);
  bl.getRange('A2').setFormula(
    '=IFERROR(SORT(UNIQUE(FILTER(Riwayat!K2:K,Riwayat!K2:K<>"")),1,FALSE),"")');
  bl.getRange('B2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(Riwayat!K2:K,A2:A)))');
  bl.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Riwayat!I2:I,Riwayat!K2:K,A2:A)))');
  bl.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Riwayat!I2:I,Riwayat!K2:K,A2:A,Riwayat!J2:J,"Lunas")))');
  bl.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",C2:C-D2:D))');
  bl.getRange('A2:A').setNumberFormat('MMMM yyyy');
  bl.getRange('C2:E').setNumberFormat(RP_FORMAT);
  bl.setColumnWidths(1, 5, 150);

  // ---- Dashboard ----
  const ds = getOrCreate_(ss, SH_DASH);
  ds.clear();
  ds.getRange('A1').setValue('DASHBOARD PENDAPATAN — WSM Visual')
    .setFontSize(14).setFontWeight('bold');
  const rows = [
    ['Total Pendapatan (Lunas)',
      '=SUMIFS(Riwayat!I2:I,Riwayat!J2:J,"Lunas")'],
    ['Total Tagihan (semua invoice)',
      '=SUM(Riwayat!I2:I)'],
    ['Belum Dibayar',
      '=B4-B3'],
    ['Jumlah Invoice',
      '=COUNTA(Riwayat!B2:B)'],
    ['Pendapatan Bulan Ini (Lunas)',
      '=SUMIFS(Riwayat!I2:I,Riwayat!J2:J,"Lunas",Riwayat!K2:K,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'],
    ['Tagihan Bulan Ini',
      '=SUMIFS(Riwayat!I2:I,Riwayat!K2:K,DATE(YEAR(TODAY()),MONTH(TODAY()),1))'],
    ['Pendapatan Tahun Ini (Lunas)',
      '=SUMIFS(Riwayat!I2:I,Riwayat!J2:J,"Lunas",Riwayat!K2:K,">="&DATE(YEAR(TODAY()),1,1))']
  ];
  rows.forEach(function (r, i) {
    ds.getRange(3 + i, 1).setValue(r[0]).setFontWeight('bold');
    ds.getRange(3 + i, 2).setFormula(r[1]);
  });
  ds.getRange('B3:B5').setNumberFormat(RP_FORMAT);
  ds.getRange('B6').setNumberFormat('0');
  ds.getRange('B7:B9').setNumberFormat(RP_FORMAT);
  ds.setColumnWidth(1, 260);
  ds.setColumnWidth(2, 180);

  buatPanduan_(ss);

  // urutan tab (tidak fatal kalau gagal)
  try {
    ss.setActiveSheet(ds);
    ss.moveActiveSheet(1);
  } catch (e) {}

  // hapus sheet kosong bawaan
  ['Sheet1', 'Sheet 1', 'Lembar1', 'Lembar 1'].forEach(function (n) {
    const def = ss.getSheetByName(n);
    if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  });

  Logger.log('Setup selesai. Lanjut jalankan tesKirim() lalu Deploy > Web app.');
}

/* ===================== TAB PANDUAN ===================== */
function buatPanduan_(ss) {
  const sh = getOrCreate_(ss, 'Panduan');
  sh.clear();
  const baris = [
    ['PANDUAN SISTEM INVOICE WSM'],
    [''],
    ['CARA KERJA'],
    ['Cetak invoice di halaman web, data terkirim otomatis dan tercatat di tab Riwayat. Dashboard dan Rekap Bulanan menghitung sendiri lewat rumus.'],
    [''],
    ['TAB'],
    ['Dashboard: total pendapatan, total tagihan, belum dibayar, angka bulan ini dan tahun ini.'],
    ['Riwayat: satu baris per invoice. Ubah kolom Status lewat dropdown (Belum Dibayar / Lunas).'],
    ['Detail Item: rincian barang atau jasa tiap invoice.'],
    ['Rekap Bulanan: jumlah invoice, total tagihan, lunas, dan belum dibayar per bulan.'],
    [''],
    ['KALAU INVOICE SUDAH DIBAYAR'],
    ['Buka tab Riwayat, ubah Status invoice tersebut menjadi Lunas. Pendapatan di Dashboard dan Rekap ikut berubah.'],
    [''],
    ['KALAU ADA CETAK YANG DIBATALKAN'],
    ['Hapus barisnya di tab Riwayat dan Detail Item. Semua total menghitung ulang otomatis.'],
    [''],
    ['JANGAN DIUBAH'],
    ['Baris judul (baris 1) dan rumus di tab Dashboard, Rekap Bulanan, serta kolom Bulan di Riwayat.'],
    ['Locale spreadsheet (File > Settings) jangan diganti ke Indonesia, karena rumus yang ditulis script bisa error.'],
    [''],
    ['KALAU KODE APPS SCRIPT DIUBAH'],
    ['Terapkan > Kelola deployment > ikon pensil > Versi baru > Terapkan. URL Web App tidak berubah.']
  ];
  sh.getRange(1, 1, baris.length, 1).setValues(baris);
  sh.setColumnWidth(1, 720);
  sh.getRange(1, 1, baris.length, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  [3, 6, 12, 15, 18, 23].forEach(function (r) {
    sh.getRange(r, 1).setFontWeight('bold').setFontColor('#10A77C');
  });
}

/* ===================== TES (tanpa website) ===================== */
// Jalankan sekali setelah setup() untuk memastikan backend bisa menulis.
// Setelah dicek di sheet, hapus baris "TES/001" di Riwayat dan Detail Item.
function tesKirim() {
  const contoh = {
    token: TOKEN,
    noInvoice: 'TES/001',
    klien: 'Klien Percobaan',
    tanggalAcara: '10 Oktober 2026',
    lokasi: 'Situbondo',
    subtotal: 2500000, diskon: 100000, pajak: 0, total: 2400000,
    status: 'Belum Dibayar',
    items: [
      { item: 'Sewa LED Videotron', qty: 10, satuan: 'meter', harga: 150000, total: 1500000, ket: '' },
      { item: 'Operator Kamera', qty: 2, satuan: 'orang', harga: 500000, total: 1000000, ket: '' }
    ]
  };
  const hasil = doPost({ postData: { contents: JSON.stringify(contoh) } });
  Logger.log(hasil.getContent());
}

/* ===================== ENDPOINT ===================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) return out_({ ok: true, message: 'Backend invoice WSM aktif' });
  if (p.token !== TOKEN) return out_({ ok: false, error: 'Token salah' });
  try {
    if (p.action === 'list') return out_(daftarInvoice_());
    if (p.action === 'get') return out_(ambilInvoice_(p.no));
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ===================== RIWAYAT (dibaca oleh menu Riwayat di web) ===================== */
function cariBaris_(rw, noInv) {
  const last = rw.getLastRow();
  if (last < 2) return -1;
  const nos = rw.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < nos.length; i++) {
    if (String(nos[i][0]).trim() === noInv) return i + 2;
  }
  return -1;
}

function daftarInvoice_() {
  const rw = ss_().getSheetByName(SH_RIWAYAT);
  if (!rw) return { ok: false, error: 'Jalankan setup() dulu' };
  const last = rw.getLastRow();
  if (last < 2) return { ok: true, data: [] };
  const v = rw.getRange(2, 1, last - 1, 10).getValues();
  const data = [];
  for (let i = v.length - 1; i >= 0; i--) {      // terbaru di atas
    const r = v[i];
    if (!String(r[1]).trim()) continue;
    data.push({
      tgl: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      no: String(r[1]).trim(), klien: String(r[2]), tglAcara: String(r[3]), lokasi: String(r[4]),
      subtotal: Number(r[5]) || 0, diskon: Number(r[6]) || 0, pajak: Number(r[7]) || 0,
      total: Number(r[8]) || 0, status: r[9] === 'Lunas' ? 'Lunas' : 'Belum Dibayar'
    });
  }
  return { ok: true, data: data };
}

function ambilInvoice_(no) {
  const ss = ss_();
  const rw = ss.getSheetByName(SH_RIWAYAT);
  const it = ss.getSheetByName(SH_ITEM);
  const noInv = String(no || '').trim();
  const row = cariBaris_(rw, noInv);
  if (row === -1) return { ok: false, error: 'Invoice tidak ditemukan' };
  const r = rw.getRange(row, 1, 1, 10).getValues()[0];
  const items = [];
  const lastIt = it.getLastRow();
  if (lastIt >= 2) {
    it.getRange(2, 1, lastIt - 1, 9).getValues().forEach(function (x) {
      if (String(x[0]).trim() === noInv) {
        items.push({ item: String(x[3]), qty: Number(x[4]) || 0, satuan: String(x[5]),
                     harga: Number(x[6]) || 0, total: Number(x[7]) || 0, ket: String(x[8]) });
      }
    });
  }
  return { ok: true, data: {
    no: noInv, klien: String(r[2]), tglAcara: String(r[3]), lokasi: String(r[4]),
    subtotal: Number(r[5]) || 0, diskon: Number(r[6]) || 0, pajak: Number(r[7]) || 0,
    total: Number(r[8]) || 0, status: r[9] === 'Lunas' ? 'Lunas' : 'Belum Dibayar', items: items
  } };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out_({ ok: false, error: 'Token salah' });
    if (!d.noInvoice) return out_({ ok: false, error: 'No invoice kosong' });

    // ubah status saja (dari menu Riwayat di web)
    if (d.action === 'setStatus') {
      const rw0 = ss_().getSheetByName(SH_RIWAYAT);
      const r0 = cariBaris_(rw0, String(d.noInvoice).trim());
      if (r0 === -1) return out_({ ok: false, error: 'Invoice tidak ditemukan' });
      const st = d.status === 'Lunas' ? 'Lunas' : 'Belum Dibayar';
      rw0.getRange(r0, 10).setValue(st);
      return out_({ ok: true, status: st });
    }

    const ss = ss_();
    const rw = ss.getSheetByName(SH_RIWAYAT);
    const it = ss.getSheetByName(SH_ITEM);
    if (!rw || !it) return out_({ ok: false, error: 'Jalankan setup() dulu' });

    const now = new Date();
    const noInv = String(d.noInvoice).trim();

    // cari baris invoice yang sama (upsert -> tidak dobel kalau dicetak ulang)
    const last = rw.getLastRow();
    let row = -1;
    if (last >= 2) {
      const nos = rw.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < nos.length; i++) {
        if (String(nos[i][0]).trim() === noInv) { row = i + 2; break; }
      }
    }

    const body = [
      noInv, d.klien || '', d.tanggalAcara || '', d.lokasi || '',
      Number(d.subtotal) || 0, Number(d.diskon) || 0, Number(d.pajak) || 0,
      Number(d.total) || 0, d.status === 'Lunas' ? 'Lunas' : 'Belum Dibayar'
    ];

    if (row === -1) {
      row = Math.max(last, 1) + 1;
      rw.getRange(row, 1).setValue(now);
      rw.getRange(row, 2, 1, 9).setValues([body]);
      rw.getRange(row, 11).setFormula('=DATE(YEAR(A' + row + '),MONTH(A' + row + '),1)');
    } else {
      rw.getRange(row, 2, 1, 9).setValues([body]);
    }
    rw.getRange(row, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    rw.getRange(row, 4).setNumberFormat('@');
    rw.getRange(row, 6, 1, 4).setNumberFormat(RP_FORMAT);
    rw.getRange(row, 11).setNumberFormat('MMMM yyyy');

    // detail item: hapus yang lama, tulis ulang
    const lastIt = it.getLastRow();
    if (lastIt >= 2) {
      const nosIt = it.getRange(2, 1, lastIt - 1, 1).getValues();
      for (let i = nosIt.length - 1; i >= 0; i--) {
        if (String(nosIt[i][0]).trim() === noInv) it.deleteRow(i + 2);
      }
    }
    const items = (d.items || []).map(function (x) {
      return [noInv, now, d.klien || '', x.item || '', Number(x.qty) || 0,
        x.satuan || '', Number(x.harga) || 0, Number(x.total) || 0, x.ket || ''];
    });
    if (items.length) {
      const start = Math.max(it.getLastRow(), 1) + 1;
      it.getRange(start, 1, items.length, 9).setValues(items);
      it.getRange(start, 2, items.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
      it.getRange(start, 7, items.length, 2).setNumberFormat(RP_FORMAT);
    }

    return out_({ ok: true, row: row });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ===================== HELPER ===================== */
function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#10A77C').setFontColor('#ffffff');
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
