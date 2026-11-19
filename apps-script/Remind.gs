/**
 * Remind.gs – nhắc hạn hằng ngày. Nguồn hạn: Sổ trình duyệt (cột Hạn phản hồi / Tình trạng).
 * Mặc định chỉ ghi tab TQN_NHAC_HAN; gửi mail khi TQN_CONFIG remind_email_enabled = Có.
 */
var TQN_REMIND_HEADERS = ['Nhóm', 'Hạn phản hồi', 'Còn / quá (ngày)', 'Mã WF', 'Số hiệu TBC', 'Loại', 'Nội dung', 'Trình đến', 'Người thực hiện', 'Vị trí'];

/** Danh sách cần nhắc: quá hạn + đến hạn trong N ngày, chưa có Kết quả. */
function tqnRemindList_(cfg, imp, today) {
  today = today || new Date();
  var t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var lim = new Date(t0.getTime() + cfg.remindDays * 86400000);
  var tbcByWf = {};
  imp.src.details.forEach(function (d) { if (d.log) tbcByWf[d.log.row] = d.code; });
  var items = [];
  imp.src.log.forEach(function (L) {
    if (L.result) return;                                  // đã có kết quả thì thôi
    var st = tqnNoAccent_(L.status);
    if (!tqnIsDate_(L.due) && st !== 'qua han') return;    // không có hạn thì không gắn quá hạn
    var due = tqnIsDate_(L.due) ? new Date(L.due.getFullYear(), L.due.getMonth(), L.due.getDate()) : null;
    if (due && due > lim) return;
    var diff = due ? Math.round((due - t0) / 86400000) : '';
    var grp = (due && diff < 0) || st === 'qua han' ? 'QUÁ HẠN' : (diff === 0 ? 'HẠN HÔM NAY' : 'SẮP ĐẾN HẠN');
    items.push([grp, due || '', diff, L.wf, tbcByWf[L.row] || (L.tbc[0] ? L.tbc[0].code : ''), L.type, L.subject, L.to, L.owner, L.sheet + '!' + L.row]);
  });
  var ord = { 'QUÁ HẠN': 0, 'HẠN HÔM NAY': 1, 'SẮP ĐẾN HẠN': 2 };
  items.sort(function (a, b) { return ord[a[0]] - ord[b[0]] || ((a[1] || 0) - (b[1] || 0)); });
  return items;
}

function tqnWriteRemind_(cfg, items) {
  var sh = tqnOwnSheet_(TQN.SH_REMIND, TQN_REMIND_HEADERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, TQN_REMIND_HEADERS.length).clearContent();
  if (items.length) {
    sh.getRange(2, 1, items.length, TQN_REMIND_HEADERS.length).setValues(items);
    sh.getRange(2, 2, items.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  return sh;
}

function tqnRemindHtml_(cfg, items, today) {
  var esc = function (s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var color = { 'QUÁ HẠN': '#A32D2D', 'HẠN HÔM NAY': '#854F0B', 'SẮP ĐẾN HẠN': '#185FA5' };
  var rows = items.map(function (r) {
    return '<tr><td style="color:' + color[r[0]] + ';font-weight:600">' + esc(r[0]) + '</td><td>' + esc(tqnFmtDate_(r[1], cfg.timezone)) + '</td><td>' + esc(r[2]) +
      '</td><td>' + esc(r[4] || r[3]) + '</td><td>' + esc(r[6]) + '</td><td>' + esc(r[7]) + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;font-size:13px"><h3 style="color:#1B3F7A;margin:0 0 6px">TQN Register · ' + esc(cfg.project_name) + '</h3>' +
    '<p>Nhắc hạn ngày ' + esc(tqnFmtDate_(today, cfg.timezone)) + ': ' + items.length + ' hồ sơ.</p>' +
    '<table cellpadding="5" style="border-collapse:collapse" border="1"><tr style="background:#EEF3FB"><th>Nhóm</th><th>Hạn</th><th>Ngày</th><th>Số hiệu</th><th>Nội dung</th><th>Trình đến</th></tr>' +
    rows + '</table><p style="color:#888">Nguồn: ' + esc(cfg.log_sheet) + '. Tự động bởi TQN Register ' + TQN.VERSION + '.</p></div>';
}

/** Hàm cho trigger hằng ngày. Chống chạy trùng trong cùng ngày bằng Script Properties. */
function tqnDailyJob() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) { tqnAudit_('NHAC_HAN', 'Bỏ qua: đang có lần chạy khác giữ khóa', {}); return; }
  try {
    var cfg = tqnConfig_(), today = new Date();
    var dayKey = cfg.project_id + ':' + Utilities.formatDate(today, cfg.timezone, 'yyyy-MM-dd') + ':remind';
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('lastRemind') === dayKey) { tqnAudit_('NHAC_HAN', 'Bỏ qua: hôm nay đã chạy ' + dayKey, {}); return; }
    var imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), items = tqnRemindList_(cfg, imp, today);
    tqnWriteRemind_(cfg, items);
    var sent = 0;
    if (cfg.emailOn && cfg.remind_recipients && items.length) {
      MailApp.sendEmail({
        to: cfg.remind_recipients,
        subject: '[TQN] Nhắc hạn ' + cfg.project_id + ' ' + Utilities.formatDate(today, cfg.timezone, 'dd/MM') + ': ' + items.length + ' hồ sơ',
        htmlBody: tqnRemindHtml_(cfg, items, today)
      });
      sent = 1;
    }
    props.setProperty('lastRemind', dayKey);
    tqnAudit_('NHAC_HAN', dayKey, { nhac: items.length, gui_mail: sent, qc: qc.rows.length });
  } finally { lock.releaseLock(); }
}

/** Bật trigger hằng ngày (xóa trigger cũ cùng tên trước, nên bấm nhiều lần không sinh trùng). */
function tqnInstallDaily() {
  var cfg = tqnConfig_();
  tqnRemoveDaily(true);
  ScriptApp.newTrigger('tqnDailyJob').timeBased().atHour(cfg.remindHour).everyDays(1).inTimezone(cfg.timezone).create();
  tqnAudit_('BAT_NHAC_HAN', 'Hằng ngày khoảng ' + cfg.remindHour + 'h (' + cfg.timezone + ')', {});
  return cfg.remindHour;
}
function tqnRemoveDaily(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnDailyJob') { ScriptApp.deleteTrigger(t); n++; } });
  if (!silent) tqnAudit_('TAT_NHAC_HAN', 'Đã xóa ' + n + ' trigger', {});
  return n;
}
