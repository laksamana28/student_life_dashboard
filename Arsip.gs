// ============================================================
// Arsip.gs — Fungsi Arsip Tugas Selesai
// TEMPELKAN file ini ke proyek Apps Script kamu
// (atau copy semua kode ini ke bawah Code.gs yang sudah ada)
//
// Kolom sheet: TANGGAL(0) HARI(1) WAKTU(2) RUANG(3) AKTIVITAS(4)
//              TUGAS(5) INPROGRES(6) SELESAI(7) DEADLINE(8)
//
// Baris arsip dimulai dari baris ke-162 (index sheet ke-161,
// karena baris 1 = header). Ubah ARSIP_START_ROW jika berbeda.
// ============================================================

// ── KONFIGURASI ───────────────────────────────────────────────

const ARSIP_CACHE_TTL = 300; // Cache 5 menit

// ── MENU: tambahkan item Arsip ke menu Dashboard ──────────────
// Panggil addArsipMenu() dari dalam onOpen() yang sudah ada,
// atau ganti seluruh onOpen() dengan versi di bawah ini:
//
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('📊 Dashboard')
//     .addItem('🚀 Buka Dashboard Fullscreen', 'openDashboard')
//     .addItem('🗄️ Buka Arsip Tugas Selesai', 'openArsip')   // ← tambahkan ini
//     .addToUi();
// }

function openArsip() {
  var html = HtmlService.createHtmlOutputFromFile('arsip')
    .setWidth(1400)
    .setHeight(960);
  SpreadsheetApp.getUi().showModelessDialog(html, '🗄️ Arsip Tugas Selesai');
}

// ── FUNGSI UTAMA: getArsipData() ──────────────────────────────
// Dipanggil dari arsip.html via google.script.run.getArsipData()
function getArsipData() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'arsip_data';

  // Cek cache dulu
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const allRows = sheet.getDataRange().getValues();
  const totalRows = allRows.length;

  // Baris arsip = mulai dari ARSIP_START_ROW (1-indexed)
  // Index array = ARSIP_START_ROW - 1
  const startIdx = ARSIP_START_ROW - 1;

  const arsipList = [];
  let totalTugas  = 0;  // untuk hitung completion rate global
  let totalSelesai = 0;

  // ── Hitung SEMUA tugas (untuk completion rate) ────────────
  const tugasMapAll = {};
  for (let i = 1; i < totalRows; i++) {
    const r = allRows[i];
    if (!r[5] || String(r[5]).trim() === '' || String(r[5]).trim() === '-') continue;
    const aktivitas = String(r[4] || '').trim();
    if (!aktivitas || aktivitas === '-') continue;
    const key = aktivitas + '||' + String(r[5]).trim();
    if (!tugasMapAll[key]) {
      tugasMapAll[key] = {
        selesai   : String(r[7] || '').trim().toUpperCase(),
        inprogres : String(r[6] || '').trim().toUpperCase(),
        deadline  : formatDate(r[8])
      };
    } else {
      if (String(r[7] || '').trim().toUpperCase() === 'SELESAI') tugasMapAll[key].selesai = 'SELESAI';
      if (String(r[6] || '').trim().toUpperCase() === 'PROSES')  tugasMapAll[key].inprogres = 'PROSES';
    }
  }
  const today = new Date(); today.setHours(0,0,0,0);
  Object.values(tugasMapAll).forEach(function(t) {
    totalTugas++;
    if (t.selesai === 'SELESAI') {
      totalSelesai++;
    } else {
      const dl = parseDate(t.deadline);
      if (dl && dl < today) totalSelesai++;
    }
  });
  const completionRate = totalTugas > 0 ? Math.round(totalSelesai / totalTugas * 100) : 0;

  // ── Ambil data ARSIP mulai baris ARSIP_START_ROW ──────────
  // Kita juga pakai deduplikasi agar satu tugas tidak muncul 2x
  const arsipKeySet = new Set();

  for (let i = startIdx; i < totalRows; i++) {
    const r = allRows[i];

    // Skip baris kosong
    if (!r[0] && !r[1]) continue;
    const aktivitas = String(r[4] || '').trim();
    if (!aktivitas || aktivitas === '-') continue;

    const tugas    = String(r[5] || '').trim();
    const selesai  = String(r[7] || '').trim().toUpperCase();
    const inprogres= String(r[6] || '').trim().toUpperCase();
    const deadline = formatDate(r[8]);

    // ── Hanya masukkan yang SELESAI ──────────────────────────
    // Kriteria selesai:
    //   1. Kolom SELESAI = 'SELESAI'  (prioritas utama)
    //   2. Tidak ada kata tugas → skip (baris aktivitas biasa)
    //   3. Deadline sudah lewat & bukan in-progress → anggap selesai (fallback)
    if (!tugas || tugas === '-') continue;

    const isSelesai = selesai === 'SELESAI';
    const dl = parseDate(deadline);
    const isDLPast  = dl && dl < today;
    const isFallback = !isSelesai && inprogres !== 'PROSES' && isDLPast;

    if (!isSelesai && !isFallback) continue;

    // Dedup: satu tugas unik per MK
    const key = aktivitas.toUpperCase() + '||' + tugas;
    if (arsipKeySet.has(key)) continue;
    arsipKeySet.add(key);

    // Ref baris = label "162A", "163A", dst (A = Sheet1)
    const refBaris = (i + 1) + 'A';

    arsipList.push({
      tanggal  : formatDate(r[0]),
      hari     : String(r[1] || '').trim().toUpperCase(),
      waktu    : String(r[2] || '').trim(),
      ruang    : String(r[3] || '').trim(),
      mk       : toTitleCase(aktivitas),
      tugas    : tugas,
      inprogres: inprogres,
      selesai  : selesai,
      deadline : deadline,
      refBaris : refBaris,
      rowIndex : i + 1    // 1-based row number di sheet
    });
  }

  const result = {
    arsip          : arsipList,
    totalArsip     : arsipList.length,
    completionRate : completionRate,
    totalTugas     : totalTugas,
    totalSelesai   : totalSelesai,
    arsipStartRow  : ARSIP_START_ROW,
    lastUpdated    : Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm:ss'
    )
  };

  // Simpan ke cache
  try {
    const json = JSON.stringify(result);
    if (json.length < 100000) cache.put(cacheKey, json, ARSIP_CACHE_TTL);
  } catch (e) {}

  return result;
}

// ── INVALIDATE CACHE ARSIP ────────────────────────────────────
// Panggil ini saat data sheet diubah agar cache segar
function clearArsipCache() {
  CacheService.getScriptCache().remove('arsip_data');
}

// ── HELPER: formatDate (sama dengan di Code.gs) ───────────────
// Jika sudah ada di Code.gs, HAPUS fungsi ini agar tidak duplikat
/*
function formatDate(val) {
  if (!val || val === '-' || val === '') return '-';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const num = parseFloat(String(val));
  if (!isNaN(num) && num > 40000) {
    const d = new Date((num - 25569) * 86400 * 1000);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(val).trim();
}
*/

// ── HELPER: parseDate (sama dengan di Code.gs) ────────────────
// Jika sudah ada di Code.gs, HAPUS fungsi ini agar tidak duplikat
/*
function parseDate(str) {
  if (!str || str === '-' || str === '') return null;
  const s = String(str).trim();
  const p = s.split('/');
  if (p.length === 3) {
    const d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    d.setHours(0, 0, 0, 0);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}
*/

// ── HELPER: toTitleCase (sama dengan di Code.gs) ─────────────
// Jika sudah ada di Code.gs, HAPUS fungsi ini agar tidak duplikat
/*
function toTitleCase(str) {
  return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}
*/
