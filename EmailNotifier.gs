// ============================================================
// EmailNotifier.gs — UPDATED untuk struktur sheet baru
// Kolom baru: TANGGAL(0) HARI(1) WAKTU(2) RUANG(3) AKTIVITAS(4)
//             TUGAS(5) INPROGRES(6) SELESAI(7) DEADLINE(8)
// ============================================================

const EMAIL_CONFIG = {
  subject_prefix: '📚 Reminder Tugas',
  warn_days : 3,
  urgent_days: 1,
};

// ── 1. FUNGSI UTAMA: KIRIM EMAIL NOTIFIKASI ───────────────────
function sendDeadlineReminder() {
  const props   = PropertiesService.getUserProperties();
  const enabled = props.getProperty('notif_enabled') !== 'false';
  if (!enabled) return;

  const email = getUserEmail();
  if (!email) return;

  const warnDays = parseInt(props.getProperty('notif_warn_days') || EMAIL_CONFIG.warn_days);
  const data     = getData();
  const today    = new Date();
  today.setHours(0, 0, 0, 0);

  const upcoming = [];
  (data.academic.deadlineList || []).forEach(function(t) {
    const dl = parseDate(t.deadline);
    if (!dl) return;
    const diffDays = Math.ceil((dl - today) / (1000 * 60 * 60 * 24));
    if (diffDays >= 0 && diffDays <= warnDays) {
      upcoming.push({
        mk      : t.mk,
        tugas   : t.tugas,
        deadline: t.deadline,
        diffDays: diffDays,
        urgent  : diffDays <= EMAIL_CONFIG.urgent_days,
      });
    }
  });

  if (upcoming.length === 0) return;

  const subject = buildSubject(upcoming, warnDays);
  const body    = buildEmailBody(upcoming, data.academic, today);
  GmailApp.sendEmail(email, subject, '', { htmlBody: body });
  console.log('Email terkirim ke ' + email + ' — ' + upcoming.length + ' tugas');
}

// ── 2. DAPATKAN EMAIL ─────────────────────────────────────────
function getUserEmail() {
  const stored = PropertiesService.getUserProperties().getProperty('notif_email');
  if (stored && stored.includes('@')) return stored;
  return Session.getActiveUser().getEmail() || null;
}

// ── 3. SET EMAIL ──────────────────────────────────────────────
function setNotifEmail(email) {
  PropertiesService.getUserProperties().setProperty('notif_email', email);
  return { success: true, email: email };
}

// ── 4. GET KONFIGURASI ────────────────────────────────────────
function getNotifConfig() {
  const props = PropertiesService.getUserProperties();
  return {
    email     : props.getProperty('notif_email') || Session.getActiveUser().getEmail(),
    enabled   : props.getProperty('notif_enabled') !== 'false',
    warnDays  : parseInt(props.getProperty('notif_warn_days') || EMAIL_CONFIG.warn_days),
    triggerSet: isTriggerSet(),
  };
}

// ── 5. SIMPAN KONFIGURASI ─────────────────────────────────────
function saveNotifConfig(cfg) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('notif_email',     cfg.email     || '');
  props.setProperty('notif_enabled',   cfg.enabled   ? 'true' : 'false');
  props.setProperty('notif_warn_days', String(cfg.warnDays || 3));

  if (cfg.enabled) setupDailyTrigger();
  else             removeTriggers();

  return { success: true };
}

// ── 6. TRIGGER HARIAN ─────────────────────────────────────────
function isTriggerSet() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'sendDeadlineReminder';
  });
}

function setupDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('sendDeadlineReminder')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  console.log('Trigger harian jam 07:00 berhasil dipasang.');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDeadlineReminder') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ── 7. TEST KIRIM EMAIL ───────────────────────────────────────
function testSendEmail() {
  const email = getUserEmail();
  if (!email) return { success: false, msg: 'Email tidak ditemukan.' };
  try {
    const data  = getData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const preview = (data.academic.deadlineList || []).filter(function(t) {
      const dl = parseDate(t.deadline);
      if (!dl) return false;
      const diff = Math.ceil((dl - today) / 86400000);
      return diff >= 0 && diff <= 14;
    }).map(function(t) {
      const dl   = parseDate(t.deadline);
      const diff = Math.ceil((dl - today) / 86400000);
      return { mk: t.mk, tugas: t.tugas, deadline: t.deadline, diffDays: diff, urgent: diff <= 1 };
    });

    const sample = preview.length ? preview : [{
      mk: 'Contoh MK', tugas: 'Tugas Contoh',
      deadline: formatDate(new Date(today.getTime() + 86400000 * 2)),
      diffDays: 2, urgent: false
    }];

    const body = buildEmailBody(sample, data.academic, today);
    GmailApp.sendEmail(email, '🧪 [TEST] ' + EMAIL_CONFIG.subject_prefix, '', { htmlBody: body });
    return { success: true, msg: 'Email test berhasil dikirim ke ' + email };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

// ── 8. BUILD SUBJECT ──────────────────────────────────────────
function buildSubject(upcoming, warnDays) {
  const urgentCount = upcoming.filter(function(t) { return t.urgent; }).length;
  if (urgentCount > 0) {
    return '🚨 ' + EMAIL_CONFIG.subject_prefix + ' — ' + urgentCount + ' tugas MENDESAK hari ini!';
  }
  return EMAIL_CONFIG.subject_prefix + ' — ' + upcoming.length + ' tugas dalam ' + warnDays + ' hari ke depan';
}

// ── 9. BUILD HTML EMAIL ───────────────────────────────────────
function buildEmailBody(upcoming, academic, today) {
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy');

  const taskRows = upcoming.map(function(t) {
    const urgentBadge = t.urgent
      ? '<span style="background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;margin-left:6px;">🚨 MENDESAK</span>'
      : '';
    const daysBadge = t.diffDays === 0
      ? '<span style="background:#fef3c7;color:#d97706;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;">Hari Ini!</span>'
      : '<span style="background:#dbeafe;color:#1d4ed8;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;">' + t.diffDays + ' hari lagi</span>';

    return '<tr>' +
      '<td style="padding:12px 16px;border-bottom:1px solid #f3f4f6;vertical-align:top;">' +
        '<div style="font-weight:700;color:#111827;font-size:13px;">' + t.mk + '</div>' +
        '<div style="color:#6b7280;font-size:12px;margin-top:2px;">' + t.tugas + '</div>' +
      '</td>' +
      '<td style="padding:12px 16px;border-bottom:1px solid #f3f4f6;vertical-align:top;text-align:center;">' +
        '<div style="font-size:12px;color:#374151;">📅 ' + t.deadline + '</div>' +
      '</td>' +
      '<td style="padding:12px 16px;border-bottom:1px solid #f3f4f6;vertical-align:top;text-align:center;">' +
        daysBadge + urgentBadge +
      '</td>' +
    '</tr>';
  }).join('');

  const urgentCount = upcoming.filter(function(t) { return t.urgent; }).length;
  const headerBg    = urgentCount > 0 ? 'linear-gradient(135deg,#991b1b,#dc2626)' : 'linear-gradient(135deg,#1e429f,#1a56db)';
  const headerEmoji = urgentCount > 0 ? '🚨' : '📚';
  const headerTitle = urgentCount > 0 ? 'TUGAS MENDESAK!' : 'Reminder Deadline Tugas';

  return `
<!DOCTYPE html>
<html lang="id">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:30px 0;">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.10);">

  <!-- HEADER -->
  <tr><td style="background:${headerBg};padding:28px 32px;">
    <div style="font-size:28px;margin-bottom:8px;">${headerEmoji}</div>
    <div style="color:#fff;font-size:22px;font-weight:800;letter-spacing:-.3px;">${headerTitle}</div>
    <div style="color:rgba(255,255,255,.75);font-size:13px;margin-top:4px;">${todayStr}</div>
  </td></tr>

  <!-- SUMMARY CHIPS -->
  <tr><td style="padding:20px 32px 0;">
    <div style="display:flex;gap:10px;flex-wrap:wrap;">
      <div style="background:#dbeafe;border-radius:8px;padding:10px 16px;text-align:center;display:inline-block;margin-right:10px;">
        <div style="font-size:22px;font-weight:800;color:#1d4ed8;">${academic.totalTugas}</div>
        <div style="font-size:11px;color:#3b82f6;font-weight:600;text-transform:uppercase;letter-spacing:.5px;">Total Tugas</div>
      </div>
      <div style="background:#d1fae5;border-radius:8px;padding:10px 16px;text-align:center;display:inline-block;margin-right:10px;">
        <div style="font-size:22px;font-weight:800;color:#065f46;">${academic.selesai}</div>
        <div style="font-size:11px;color:#059669;font-weight:600;text-transform:uppercase;letter-spacing:.5px;">Selesai</div>
      </div>
      <div style="background:#fee2e2;border-radius:8px;padding:10px 16px;text-align:center;display:inline-block;margin-right:10px;">
        <div style="font-size:22px;font-weight:800;color:#991b1b;">${upcoming.length}</div>
        <div style="font-size:11px;color:#dc2626;font-weight:600;text-transform:uppercase;letter-spacing:.5px;">Mendekati DL</div>
      </div>
      <div style="background:#ede9fe;border-radius:8px;padding:10px 16px;text-align:center;display:inline-block;">
        <div style="font-size:22px;font-weight:800;color:#5b21b6;">${academic.completionRate}%</div>
        <div style="font-size:11px;color:#7c3aed;font-weight:600;text-transform:uppercase;letter-spacing:.5px;">Completion</div>
      </div>
    </div>
  </td></tr>

  <!-- TASK TABLE -->
  <tr><td style="padding:20px 32px;">
    <div style="font-size:14px;font-weight:800;color:#111827;margin-bottom:12px;text-transform:uppercase;letter-spacing:.5px;">⚠️ Deadline Dalam Waktu Dekat</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1.5px solid #e5e7eb;border-radius:10px;overflow:hidden;">
      <tr style="background:#f9fafb;">
        <th style="padding:10px 16px;text-align:left;font-size:11px;color:#6b7280;font-weight:700;text-transform:uppercase;letter-spacing:.5px;">Mata Kuliah / Tugas</th>
        <th style="padding:10px 16px;text-align:center;font-size:11px;color:#6b7280;font-weight:700;text-transform:uppercase;letter-spacing:.5px;">Deadline</th>
        <th style="padding:10px 16px;text-align:center;font-size:11px;color:#6b7280;font-weight:700;text-transform:uppercase;letter-spacing:.5px;">Status</th>
      </tr>
      ${taskRows || '<tr><td colspan="3" style="padding:20px;text-align:center;color:#9ca3af;font-size:13px;">Tidak ada deadline dalam waktu dekat 🎉</td></tr>'}
    </table>
  </td></tr>

  <!-- MOTIVASI -->
  <tr><td style="padding:0 32px 24px;">
    <div style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:1.5px solid #86efac;border-radius:10px;padding:14px 16px;">
      <div style="font-size:12px;color:#166534;font-weight:600;line-height:1.5;">
        💪 <strong>Semangat!</strong> Selesaikan tugas satu per satu. Kamu bisa! Cek dashboard untuk progress lengkap.
      </div>
    </div>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#f9fafb;padding:16px 32px;text-align:center;border-top:1px solid #e5e7eb;">
    <div style="font-size:11px;color:#9ca3af;">
      📊 Student Life Dashboard — Email otomatis dikirim setiap pagi jam 07:00<br>
      <span style="color:#d1d5db;">Jika tidak ingin menerima email ini, nonaktifkan di panel Notifikasi dashboard.</span>
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}
