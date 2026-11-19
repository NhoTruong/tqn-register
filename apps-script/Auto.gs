/**
 * Auto.gs – bật/tắt toàn bộ tự động bằng 1 nút. Chạy trên máy chủ Google: máy anh tắt vẫn chạy.
 *  - tqnHourlyJob : mỗi giờ nhập & kiểm lại (chỉ đọc tab gốc, ghi tab TQN_*), không gửi mail.
 *  - tqnDailyJob  : nhắc hạn ~07:00 (Remind.gs).
 *  - tqnWeatherJob: thời tiết ~06:00 (Weather.gs), chỉ khi đủ điều kiện GPS + điều khoản nguồn.
 */
var TQN_AUTO_HANDLERS = ['tqnHourlyJob', 'tqnDailyJob', 'tqnWeatherJob'];

function tqnHourlyJob() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { tqnAudit_('TU_DONG_GIO', 'Bỏ qua: đang có lần chạy khác', {}); return 'BAN'; }
  try {
    tqnEnsureConfig_();
    var cfg = tqnConfig_(), imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), rem = tqnRemindList_(cfg, imp, new Date());
    tqnWriteRemind_(cfg, rem);
    var wf = tqnBuildWorkflow_(cfg, imp, new Date());
    PropertiesService.getDocumentProperties().setProperty('lastHourly', Utilities.formatDate(new Date(), cfg.timezone, 'dd/MM/yyyy HH:mm'));
    tqnAudit_('TU_DONG_GIO', 'Nhập & kiểm tự động', { du_lieu: imp.counts, qc: qc.byRule, nhac: rem.length, quy_trinh: wf.count });
    return 'OK';
  } finally { lock.releaseLock(); }
}

function tqnRemoveHourly_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnHourlyJob') { ScriptApp.deleteTrigger(t); n++; } });
  return n;
}

/** Bật tất cả. Bấm nhiều lần không sinh lịch trùng (xóa lịch cũ cùng tên trước). Trả danh sách dòng báo cáo. */
function tqnSetupAll() {
  tqnEnsureConfig_(); tqnEnsureWeatherConfig_(); tqnEnsureCodeConfig_();
  var out = [];
  tqnRemoveHourly_();
  ScriptApp.newTrigger('tqnHourlyJob').timeBased().everyHours(1).create();
  out.push('✔ Nhập & kiểm tự động: mỗi giờ');
  out.push('✔ Nhắc hạn: hằng ngày khoảng ' + tqnInstallDaily() + ':00' + (tqnConfig_().emailOn ? ' (có gửi mail)' : ' (chưa gửi mail – bật remind_email_enabled = Có để gửi)'));
  var wc = tqnWeatherCfg_(), why = tqnWeatherBlocked_(wc);
  if (why) { tqnRemoveWeather(true); out.push('✖ Thời tiết 06:00 CHƯA bật: ' + why + ' → sửa ở tab TQN_CONFIG rồi bấm lại nút này'); }
  else out.push('✔ Thời tiết: hằng ngày khoảng ' + tqnInstallWeather() + ':00');
  try { var th = tqnApplyTheme(); out.push('✔ Giao diện dễ nhìn: ' + th.tabs.length + ' sổ (bản gốc lưu ở TQN_GIAO_DIEN_GOC)'); } catch (e) { out.push('✖ Giao diện: ' + e.message); }
  try { tqnHourlyJob(); out.push('✔ Đã chạy nhập & kiểm lần đầu'); } catch (e) { out.push('✖ Chạy lần đầu lỗi: ' + e.message); }
  tqnAudit_('BAT_TU_DONG', out.join(' | '), {});
  return out;
}

function tqnRemoveAll() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (TQN_AUTO_HANDLERS.indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); n++; } });
  tqnAudit_('TAT_TU_DONG', 'Đã xóa ' + n + ' lịch', {});
  return n;
}

/** Trạng thái các lịch đang bật + lần chạy gần nhất. */
function tqnAutoStatus() {
  var on = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { on[t.getHandlerFunction()] = (on[t.getHandlerFunction()] || 0) + 1; });
  var p = PropertiesService.getDocumentProperties();
  return [
    (on.tqnHourlyJob ? '✔' : '✖') + ' Nhập & kiểm mỗi giờ' + (p.getProperty('lastHourly') ? ' – lần gần nhất ' + p.getProperty('lastHourly') : ''),
    (on.tqnDailyJob ? '✔' : '✖') + ' Nhắc hạn hằng ngày' + (p.getProperty('lastRemind') ? ' – ' + p.getProperty('lastRemind') : ''),
    (on.tqnWeatherJob ? '✔' : '✖') + ' Thời tiết 06:00' + (p.getProperty('lastWeather') ? ' – ' + p.getProperty('lastWeather') : '')
  ];
}

/* menu */
function tqnMenuSetupAll() {
  SpreadsheetApp.getUi().alert('Bật tự động TQN Register', tqnSetupAll().join('\n') + '\n\nGoogle tự chạy trên máy chủ, máy anh tắt vẫn chạy (có thể lệch ±15 phút).', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRemoveAll() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveAll() + ' lịch tự động.'); }
function tqnMenuAutoStatus() { SpreadsheetApp.getUi().alert('Trạng thái tự động', tqnAutoStatus().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }
