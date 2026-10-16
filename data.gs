// ============================================================
// Code.gs — Student Life Dashboard (UPDATED + ARSIP SYNC)
// Kolom: TANGGAL(0) HARI(1) WAKTU(2) RUANG(3) AKTIVITAS(4)
//        TUGAS(5) INPROGRES(6) SELESAI(7) DEADLINE(8) ARSIP(9)
// Auto-sync: arsip membaca dari baris 2, sinkronisasi otomatis
// ============================================================

// ── CACHE DURATION (detik) ───────────────────────────────────
const CACHE_TTL = 300; // 5 menit

// ── ARSIP CONSTANT ───────────────────────────────────────────
const ARSIP_START_ROW = 180; // Baris awal area arsip (A180)

// ── 1. MENU ───────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Dashboard')
    .addItem('🚀 Buka Dashboard Fullscreen', 'openDashboard')
    .addItem('🗄️ Buka Arsip Tugas Selesai', 'openArsip')
    .addSeparator()
    .addItem('📦 Arsipkan Tugas Selesai', 'arsipkanTugasSelesai')
    .addItem('🔄 Sync Arsip Data', 'syncArsipManual')
    .addItem('🔄 Refresh Cache', 'clearDashboardCache')
    .addToUi();
}

// ── 2. BUKA DASHBOARD ─────────────────────────────────────────
function openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('dashboard')
      .setWidth(1400)
      .setHeight(1000);
  SpreadsheetApp.getUi().showModelessDialog(html, '📊 Student Life Dashboard');
}

// ── 2b. BUKA ARSIP ────────────────────────────────────────────
function openArsip() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.activate();
  const arsipHeader = sheet.getRange('A' + ARSIP_START_ROW);
  sheet.setActiveRange(arsipHeader);
  SpreadsheetApp.getUi().alert('✅ Arsip tugas selesai ada di baris ' + ARSIP_START_ROW + ' ke bawah.\nSilakan scroll untuk melihat daftar arsip.');
}

// ── 3. GET DATA (dengan CacheService) ─────────────────────────
function getData() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'dashboard_data';

  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log('Sheet "Sheet1" tidak ditemukan!');
    return processData([]);
  }

  const lastRow = sheet.getLastRow();
  const maxData = findDataEndRow(sheet, lastRow);
  if (maxData < 2) return processData([]);

  const rows = sheet.getRange(1, 1, maxData, 10).getValues();

  const rawData = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;
    const aktivitas = String(r[4] || '').trim();
    if (!aktivitas || aktivitas === '-') continue;

    rawData.push({
      tanggal   : formatDate(r[0]),
      hari      : String(r[1] || '').trim().toUpperCase(),
      waktu     : String(r[2] || '').trim(),
      ruang     : String(r[3] || '').trim(),
      aktivitas : aktivitas.toUpperCase(),
      tugas     : String(r[5] || '').trim(),
      inprogres : String(r[6] || '').trim().toUpperCase(),
      selesai   : String(r[7] || '').trim().toUpperCase(),
      deadline  : formatDate(r[8]),
      arsip     : String(r[9] || '').trim().toUpperCase()
    });
  }

  const result = processData(rawData);

  try {
    const json = JSON.stringify(result);
    if (json.length < 100000) cache.put(cacheKey, json, CACHE_TTL);
  } catch (e) {}

  return result;
}

// ── 4. INVALIDATE CACHE ───────────────────────────────────────
function clearDashboardCache() {
  CacheService.getScriptCache().remove('dashboard_data');
  SpreadsheetApp.getUi().alert('✅ Cache dashboard berhasil di-refresh!');
}

function clearDashboardCacheSilent() {
  CacheService.getScriptCache().remove('dashboard_data');
}

// ── 5. FORMAT TANGGAL ─────────────────────────────────────────
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

// ── 6. PROSES DATA ────────────────────────────────────────────
function processData(data) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const WEEKEND    = new Set(['SABTU', 'MINGGU']);
  const SKIP       = ['TIDUR', 'MASAK', 'MANDI', 'ISTIRAHAT', 'MAKAN', 'TUDUR',
                       'BERSIH', 'SHOLAT'];
  const STUDY_KEYS = [
    'BELAJAR','BELAHAR','STATISTIKA','JARINGAN KOMPUTER','KEWARGANEGARAAN',
    'BAHASA INGGRIS','BAHSA INGGRIS','KECERDASAN BUATAN',
    'PEMROGRAMAN BERORIENTASI OBJEK','TEKNIK DIGITAL','SISTEM BASIS DATA',
    'PRAKTIKUM','JOBSHEET','SEMINAR','UTS','UAS','QUIZ'
  ];
  const KATEGORI_MAP = {
    'OLAHRAGA'          : 'Workout / Olahraga',
    'WORKOUT'           : 'Workout / Olahraga',
    'BERMAIN'           : 'Hiburan / Healing',
    'MAIN'              : 'Hiburan / Healing',
    'BELANJA'           : 'Hiburan / Healing',
    'BELAJAR'           : 'Belajar',
    'BELAHAR'           : 'Belajar',
    'MEMBERSIHKAN RUMAH': 'Organisasi',
    'BERSIH'            : 'Organisasi',
    'FREELANCE'         : 'Freelancer',
    'FREELANCER'        : 'Freelancer',
  };

  const weekRanges = [0, 1, 2, 3].map(w => getWeekRange(w));
  const curWeek    = weekRanges[0];
  const weekStudy  = [0, 0, 0, 0];

  const tugasMap    = {};
  const weekendDist = {};
  let studyH = 0, lifeH = 0;
  const arsipList   = [];

  for (const d of data) {
    const akt  = d.aktivitas;
    const dur  = parseDuration(d.waktu);
    const tgl  = parseDate(d.tanggal);
    const skip = SKIP.some(s => akt.includes(s));

    const isArsip = (d.arsip === 'ARSIP' || d.arsip === 'YA' || d.arsip === 'Y');

    // ── TUGAS ─────────────────────────────────────────────
    if (d.tugas && d.tugas !== '-' && d.tugas !== '') {
      const key = akt + '||' + d.tugas;

      if (isArsip) {
        arsipList.push({
          mk        : toTitleCase(akt),
          tugas     : d.tugas,
          deadline  : d.deadline,
          tanggal   : d.tanggal,
          selesai   : d.selesai,
          arsipDate : d.tanggal
        });
        continue;
      }

      if (!tugasMap[key]) {
        tugasMap[key] = {
          mk        : akt,
          tugas     : d.tugas,
          deadline  : d.deadline,
          tanggal   : d.tanggal,
          inprogres : d.inprogres,
          selesai   : d.selesai,
          arsip     : d.arsip
        };
      } else {
        if (d.selesai === 'SELESAI')   tugasMap[key].selesai   = 'SELESAI';
        if (d.inprogres === 'PROSES')  tugasMap[key].inprogres = 'PROSES';
      }
    }

    if (isArsip) continue;

    // ── WEEKEND distribusi ────────────────────────────────
    if (WEEKEND.has(d.hari) && !skip) {
      let kat = 'Organisasi';
      for (const [k, v] of Object.entries(KATEGORI_MAP)) {
        if (akt.includes(k)) { kat = v; break; }
      }
      weekendDist[kat] = (weekendDist[kat] || 0) + dur;
    }

    // ── BALANCE & WEEKLY TREND ────────────────────────────
    if (tgl && !skip) {
      const isStudy = STUDY_KEYS.some(k => akt.includes(k));
      if (tgl >= curWeek.start && tgl <= curWeek.end) {
        if (isStudy) studyH += dur; else lifeH += dur;
      }
      for (let w = 0; w < 4; w++) {
        if (tgl >= weekRanges[w].start && tgl <= weekRanges[w].end && isStudy) {
          weekStudy[w] += dur;
        }
      }
    }
  }

  // ── POST-PROCESS TUGAS ─────────────────────────────────
  const tugasList  = Object.values(tugasMap);
  let selesai = 0, inProgress = 0, belumMulai = 0;
  const tugasPerMK = {};

  for (const t of tugasList) {
    if (t.selesai === 'SELESAI') {
      selesai++;
    } else if (t.inprogres === 'PROSES') {
      inProgress++;
    } else {
      const dl = parseDate(t.deadline);
      if (dl && dl < today) selesai++;
      else                  belumMulai++;
    }

    const mk = toTitleCase(t.mk);
    tugasPerMK[mk] = (tugasPerMK[mk] || 0) + 1;
  }

  const totalTugas     = tugasList.length;
  const completionRate = totalTugas > 0 ? Math.round(selesai / totalTugas * 100) : 0;

  const deadlineList = tugasList
    .filter(t => t.deadline !== '-')
    .map(t => ({ mk: toTitleCase(t.mk), tugas: t.tugas, deadline: t.deadline }))
    .sort((a, b) => (parseDate(a.deadline) || 0) - (parseDate(b.deadline) || 0));

  // ── POST-PROCESS WEEKEND ──────────────────────────────
  Object.keys(weekendDist).forEach(k => {
    weekendDist[k] = Math.round(weekendDist[k] * 10) / 10;
  });
  const totalDurasiWknd = Object.values(weekendDist).reduce((s, v) => s + v, 0);
  const favEntry        = Object.entries(weekendDist).sort((a, b) => b[1] - a[1])[0] || ['Belajar', 0];

  // ── WEEKLY TREND ──────────────────────────────────────
  studyH = Math.round(studyH * 10) / 10;
  lifeH  = Math.round(lifeH  * 10) / 10;
  const total        = studyH + lifeH;
  const balanceRatio = total > 0 ? Math.round(studyH / total * 100) / 100 : 0;
  const weeklyTrend  = weekRanges.map((_, i) => ({
    label      : 'Minggu ' + (4 - i),
    studyHours : Math.round(weekStudy[i] * 10) / 10
  })).reverse();

  return {
    academic: {
      totalTugas, selesai, inProgress, belumMulai,
      completionRate, tugasPerMK, deadlineList
    },
    weekend: {
      totalAktivitas   : Object.keys(weekendDist).length,
      totalDurasi      : Math.round(totalDurasiWknd * 10) / 10,
      aktivitasFavorit : favEntry[0],
      distribusi       : weekendDist
    },
    balance: {
      studyHours: studyH, lifestyleHours: lifeH,
      balanceRatio, weeklyTrend, completionRate
    },
    arsip: {
      totalArsip : arsipList.length,
      daftar     : arsipList
    },
    lastUpdated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
  };
}

// ── 7. HELPERS ────────────────────────────────────────────────
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

function parseDuration(waktu) {
  if (!waktu || waktu === '-') return 0;
  const m = String(waktu).match(/(\d{1,2})[:.h](\d{2})\s*[-–]\s*(\d{1,2})[:.h](\d{2})/);
  if (!m) return 0;
  let dur = (parseInt(m[3]) * 60 + parseInt(m[4])) - (parseInt(m[1]) * 60 + parseInt(m[2]));
  if (dur < 0) dur += 1440;
  return dur / 60;
}

function getWeekRange(weeksAgo) {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const diffToMon = now.getDay() === 0 ? -6 : 1 - now.getDay();
  const start = new Date(now);
  start.setDate(now.getDate() + diffToMon - weeksAgo * 7);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23, 59, 59, 999);
  return { start, end };
}

function toTitleCase(str) {
  return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

/**
 * Cari baris terakhir data utama secara DINAMIS.
 * Berhenti ketika menemukan label pemisah arsip atau mendekati ARSIP_START_ROW.
 */
function findDataEndRow(sheet, lastRow) {
  if (lastRow < 2) return 0;

  // Scan dari baris 2 sampai sebelum ARSIP_START_ROW
  const scanLimit = Math.min(lastRow, ARSIP_START_ROW - 1);
  for (let i = 2; i <= scanLimit; i++) {
    const val = String(sheet.getRange(i, 1).getValue() || '').trim();
    if (val.includes('═══') || val.toUpperCase().includes('ARSIP TUGAS SELESAI')) {
      return i - 1;
    }
  }

  // Fallback: batas aman sebelum area arsip
  return Math.min(lastRow, ARSIP_START_ROW - 2);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Student Life Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function resizeDashboard(width, height) {
  var html = HtmlService.createHtmlOutputFromFile('dashboard')
      .setWidth(width).setHeight(height);
  SpreadsheetApp.getUi().showModelessDialog(html, '📊 Student Life Dashboard');
}

/**
 * Helper: cek apakah nilai dianggap "selesai"
 * Mendukung: SELESAI, V, Y, YA, DONE, SUDAH, ✓, ✔, TRUE
 */
function isSelesai(val) {
  if (!val) return false;
  const v = String(val).trim().toUpperCase();
  return ['SELESAI', 'V', 'Y', 'YA', 'DONE', 'SUDAH', '✓', '✔', 'TRUE'].includes(v);
}

// ============================================================
// ── 8. ARSIP ─────────────────────────────────────────────────
// ============================================================

/**
 * Pastikan header arsip ada di baris ARSIP_START_ROW.
 * Kolom: TANGGAL | HARI | AKTIVITAS | TUGAS | DEADLINE | SELESAI | ARSIP_DATE
 */
function ensureArsipHeader(sheet) {
  const headerCheck = String(sheet.getRange(ARSIP_START_ROW, 1).getValue() || '').trim().toUpperCase();
  if (headerCheck === 'TANGGAL') return; // sudah ada

  const headers = ['TANGGAL', 'HARI', 'AKTIVITAS', 'TUGAS', 'DEADLINE', 'SELESAI', 'ARSIP_DATE'];
  const headerRange = sheet.getRange(ARSIP_START_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');

  // Tambahkan label pemisah di baris sebelum header
  const labelRange = sheet.getRange(ARSIP_START_ROW - 1, 1, 1, 3);
  labelRange.merge();
  labelRange.setValue('═══ 🗄️ ARSIP TUGAS SELESAI ═══');
  labelRange.setFontWeight('bold');
  labelRange.setFontSize(12);
  labelRange.setBackground('#f3f3f3');
  labelRange.setHorizontalAlignment('center');
}

/**
 * Cari baris kosong pertama di area arsip (mulai dari baris ARSIP_START_ROW + 1).
 */
function getNextArsipRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ARSIP_START_ROW) return ARSIP_START_ROW + 1;
  for (let i = ARSIP_START_ROW + 1; i <= lastRow + 1; i++) {
    const val = sheet.getRange(i, 1).getValue();
    if (!val || String(val).trim() === '') return i;
  }
  return lastRow + 1;
}

/**
 * Arsipkan semua baris yang kolom SELESAI = 'SELESAI' (atau nilai lain yang valid)
 * dan kolom ARSIP masih kosong / '-'.
 */
function arsipkanTugasSelesai() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  const ui    = SpreadsheetApp.getUi();

  ensureArsipHeader(sheet);

  const lastDataRow = Math.min(sheet.getLastRow(), ARSIP_START_ROW - 2);
  if (lastDataRow < 2) {
    ui.alert('⚠️ Tidak ada data tugas di sheet.');
    return;
  }
  const rows = sheet.getRange(1, 1, lastDataRow, 10).getValues();

  let count = 0;
  let nextRow = getNextArsipRow(sheet);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const selesaiVal = String(r[7] || '').trim();
    const arsipVal   = String(r[9] || '').trim().toUpperCase();
    const tugasVal   = String(r[5] || '').trim();
    const aktVal     = String(r[4] || '').trim();

    if (isSelesai(selesaiVal) && arsipVal !== 'ARSIP' && tugasVal && tugasVal !== '-') {
      // Cek duplikat sebelum menulis
      if (!isDuplicateArsip(sheet, aktVal, tugasVal)) {
        const rowData = [
          r[0],       // TANGGAL
          r[1],       // HARI
          aktVal,     // AKTIVITAS
          tugasVal,   // TUGAS
          r[8],       // DEADLINE
          'SELESAI',  // SELESAI
          today       // ARSIP_DATE
        ];
        sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
        nextRow++;
        count++;
      }
      // Tandai kolom ARSIP (kolom J = 10) di data utama
      sheet.getRange(i + 1, 10).setValue('ARSIP');
    }
  }

  clearDashboardCacheSilent();

  if (count > 0) {
    sheet.setActiveRange(sheet.getRange('A' + ARSIP_START_ROW));
    ui.alert('✅ Berhasil mengarsipkan ' + count + ' tugas ke baris ' + ARSIP_START_ROW + '+');
  } else {
    ui.alert('ℹ️ Tidak ada tugas baru yang perlu diarsipkan.\n(Semua tugas selesai sudah diarsipkan atau belum ada tugas selesai)');
  }
}

/**
 * Cek apakah kombinasi aktivitas+tugas sudah ada di area arsip (hindari duplikat).
 */
function isDuplicateArsip(sheet, aktVal, tugasVal) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= ARSIP_START_ROW) return false;

  const numRows = lastRow - ARSIP_START_ROW;
  const arsipRows = sheet.getRange(ARSIP_START_ROW + 1, 1, numRows, 7).getValues();
  const key = aktVal.toUpperCase() + '||' + tugasVal;

  for (const ar of arsipRows) {
    const existKey = String(ar[2] || '').trim().toUpperCase() + '||' + String(ar[3] || '').trim();
    if (existKey === key) return true;
  }
  return false;
}

/**
 * Ambil data arsip — DUAL SOURCE:
 * 1. Dari data utama (baris 2+) yang kolom ARSIP = 'ARSIP'
 * 2. Dari area arsip (baris ARSIP_START_ROW+1) untuk data historis
 * Deduplikasi berdasarkan key AKTIVITAS||TUGAS.
 */
function getArsipData() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();
  const seen    = new Set();
  const result  = [];

  // ── SUMBER 1: Data utama yang sudah ditandai ARSIP ──────
  const maxData = findDataEndRow(sheet, lastRow);
  if (maxData >= 2) {
    const mainRows = sheet.getRange(2, 1, maxData - 1, 10).getValues();
    for (let i = 0; i < mainRows.length; i++) {
      const r = mainRows[i];
      const arsipVal = String(r[9] || '').trim().toUpperCase();
      const tugasVal = String(r[5] || '').trim();
      const aktVal   = String(r[4] || '').trim();

      if (arsipVal === 'ARSIP' && tugasVal && tugasVal !== '-') {
        const key = aktVal.toUpperCase() + '||' + tugasVal;
        if (!seen.has(key)) {
          seen.add(key);
          result.push({
            tanggal   : formatDate(r[0]),
            hari      : String(r[1] || '').trim(),
            aktivitas : aktVal,
            tugas     : tugasVal,
            deadline  : formatDate(r[8]),
            selesai   : String(r[7] || '').trim(),
            arsipDate : formatDate(r[0]),
            sumber    : 'data_utama'
          });
        }
      }
    }
  }

  // ── SUMBER 2: Area arsip historis ───────────────────────
  if (lastRow > ARSIP_START_ROW) {
    const numRows = lastRow - ARSIP_START_ROW;
    const arsipRows = sheet.getRange(ARSIP_START_ROW + 1, 1, numRows, 7).getValues();
    for (let i = 0; i < arsipRows.length; i++) {
      const r = arsipRows[i];
      if (!r[0] && !r[2] && !r[3]) continue;

      const aktVal   = String(r[2] || '').trim();
      const tugasVal = String(r[3] || '').trim();
      const key = aktVal.toUpperCase() + '||' + tugasVal;

      if (!seen.has(key)) {
        seen.add(key);
        result.push({
          tanggal   : formatDate(r[0]),
          hari      : String(r[1] || '').trim(),
          aktivitas : aktVal,
          tugas     : tugasVal,
          deadline  : formatDate(r[4]),
          selesai   : String(r[5] || '').trim(),
          arsipDate : String(r[6] || '').trim(),
          sumber    : 'area_arsip'
        });
      }
    }
  }

  return {
    totalArsip : result.length,
    daftar     : result
  };
}

/**
 * Hapus semua data arsip di baris ARSIP_START_ROW+ (opsional, hati-hati!)
 * Header tetap dipertahankan.
 */
function hapusDataArsip() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ Peringatan',
    'Ini akan MENGHAPUS semua data arsip di baris ' + ARSIP_START_ROW + ' ke bawah.\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) return;

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow > ARSIP_START_ROW) {
    const clearRange = sheet.getRange(ARSIP_START_ROW + 1, 1, lastRow - ARSIP_START_ROW, 7);
    clearRange.clearContent();
  }

  clearDashboardCacheSilent();
  ui.alert('✅ Data arsip berhasil dihapus. Header di baris ' + ARSIP_START_ROW + ' tetap ada.');
}

// ============================================================
// ── 9. DIAGNOSA ───────────────────────────────────────────────
// ============================================================

function cekDataSelesai() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = Math.min(sheet.getLastRow(), ARSIP_START_ROW - 2);
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Sheet kosong.');
    return;
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  const selesaiValues = {};
  const tugasCount    = { ada: 0, kosong: 0 };
  let matchCount      = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const selesaiVal = String(r[7] || '').trim();
    const tugasVal   = String(r[5] || '').trim();

    if (selesaiVal && selesaiVal !== '-') {
      selesaiValues[selesaiVal] = (selesaiValues[selesaiVal] || 0) + 1;
    }
    if (tugasVal && tugasVal !== '-') tugasCount.ada++;
    else tugasCount.kosong++;

    if (isSelesai(selesaiVal) && tugasVal && tugasVal !== '-') {
      const arsipVal = String(r[9] || '').trim().toUpperCase();
      if (arsipVal !== 'ARSIP') matchCount++;
    }
  }

  let msg = '📋 DIAGNOSA DATA (baris 2-' + lastRow + '):\n\n';
  msg += '── Kolom H (SELESAI) - Nilai yang ditemukan:\n';
  if (Object.keys(selesaiValues).length === 0) {
    msg += '   ❌ KOSONG — tidak ada isi di kolom H\n';
  } else {
    for (const [val, cnt] of Object.entries(selesaiValues)) {
      msg += '   • "' + val + '" → ' + cnt + ' baris\n';
    }
  }
  msg += '\n── Kolom F (TUGAS):\n';
  msg += '   • Ada isi: ' + tugasCount.ada + ' baris\n';
  msg += '   • Kosong: ' + tugasCount.kosong + ' baris\n';
  msg += '\n── Baris siap diarsipkan: ' + matchCount + '\n';

  if (matchCount === 0) {
    msg += '\n⚠️ Tidak ada baris yang cocok!\n';
    msg += 'Syarat arsip: kolom H harus berisi SELESAI/V/Y/DONE/✓\n';
    msg += 'DAN kolom F harus ada nama tugas.';
  }

  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// ── 10. SYNC ARSIP ───────────────────────────────────────────
// ============================================================

/**
 * Sinkronisasi data sheet dengan area arsip.
 * - Tugas SELESAI yang belum diarsipkan → arsipkan otomatis
 * - Update arsip jika data berubah di sheet utama
 */
function syncArsipData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  ensureArsipHeader(sheet);

  const lastRow = sheet.getLastRow();
  const maxData = findDataEndRow(sheet, lastRow);
  if (maxData < 2) return { synced: 0, message: 'Tidak ada data.' };

  const mainData = sheet.getRange(2, 1, maxData - 1, 10).getValues();
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  let arsipData = [];
  if (lastRow > ARSIP_START_ROW) {
    const numArsipRows = lastRow - ARSIP_START_ROW;
    arsipData = sheet.getRange(ARSIP_START_ROW + 1, 1, numArsipRows, 7).getValues();
  }

  let count = 0;
  let nextRow = getNextArsipRow(sheet);

  for (let i = 0; i < mainData.length; i++) {
    const r = mainData[i];
    const tgl      = r[0];
    const hari     = r[1];
    const akt      = String(r[4] || '').trim();
    const tugas    = String(r[5] || '').trim();
    const selesai  = String(r[7] || '').trim();
    const dl       = r[8];
    const rowIndex = i + 2; // baris aktual di sheet (baris 1 = header)

    if (!tugas || tugas === '-' || !isSelesai(selesai)) continue;

    const key = (akt + '||' + tugas).toUpperCase();

    // Cari apakah sudah ada di arsip
    const existingIndex = arsipData.findIndex(ar =>
      (String(ar[2] || '').trim().toUpperCase() + '||' + String(ar[3] || '').trim()) === key
    );

    if (existingIndex !== -1) {
      // Sudah ada di arsip → update data utama saja
      const targetRow = ARSIP_START_ROW + 1 + existingIndex;
      sheet.getRange(targetRow, 1, 1, 5).setValues([[tgl, hari, akt, tugas, dl]]);
      sheet.getRange(rowIndex, 10).setValue('ARSIP');
    } else {
      // Belum ada → tambah ke arsip
      sheet.getRange(nextRow, 1, 1, 7).setValues([[tgl, hari, akt, tugas, dl, 'SELESAI', today]]);
      sheet.getRange(rowIndex, 10).setValue('ARSIP');
      nextRow++;
      count++;
    }
  }

  clearDashboardCacheSilent();
  return { synced: count, message: 'Sinkronisasi selesai. ' + count + ' tugas baru diarsipkan.' };
}

/**
 * Trigger otomatis saat user mengedit kolom SELESAI (H/8) atau ARSIP (J/10).
 * Pasang via Edit → Current project's triggers → onEdit (from spreadsheet, on edit).
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const row   = e.range.getRow();
    const col   = e.range.getColumn();

    // Hanya trigger di sheet utama, kolom SELESAI(8) atau ARSIP(10), area data
    const isMainSheet = (sheet.getName() === 'Sheet1' || sheet.getIndex() === 1);
    if (isMainSheet && row >= 2 && row < ARSIP_START_ROW && (col === 8 || col === 10)) {
      syncArsipData();
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err.message);
  }
}

/**
 * Arsipkan satu baris tertentu ke area arsip.
 */
function autoArsipSingleRow(sheet, rowNum) {
  ensureArsipHeader(sheet);

  const r        = sheet.getRange(rowNum, 1, 1, 10).getValues()[0];
  const aktVal   = String(r[4] || '').trim();
  const tugasVal = String(r[5] || '').trim();
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  if (isDuplicateArsip(sheet, aktVal, tugasVal)) {
    sheet.getRange(rowNum, 10).setValue('ARSIP');
    return;
  }

  const nextRow = getNextArsipRow(sheet);
  const rowData = [
    r[0],       // TANGGAL
    r[1],       // HARI
    aktVal,     // AKTIVITAS
    tugasVal,   // TUGAS
    r[8],       // DEADLINE
    'SELESAI',  // SELESAI
    today       // ARSIP_DATE
  ];
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  sheet.getRange(rowNum, 10).setValue('ARSIP');
}

/**
 * Sync manual dari menu Dashboard
 */
function syncArsipManual() {
  const result = syncArsipData();
  const ui     = SpreadsheetApp.getUi();
  if (result.synced > 0) {
    ui.alert('✅ ' + result.message);
  } else {
    ui.alert('ℹ️ Semua data sudah sinkron. Tidak ada tugas baru yang perlu diarsipkan.');
  }
}
