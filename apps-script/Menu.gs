/**
 * Menu.gs – menu "TQN Register" và thanh bên bảng điều khiển.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('TQN Register')
    .addItem('Bảng điều khiển', 'tqnShowSidebar')
    .addSeparator()
    .addItem('1. Nhập & kiểm tra ngay', 'tqnMenuRun')
    .addItem('2. Xem nhắc hạn hôm nay', 'tqnMenuRemindPreview')
    .addSeparator()
    .addItem('Bật nhắc hạn hằng ngày', 'tqnMenuInstallDaily')
    .addItem('Tắt nhắc hạn hằng ngày', 'tqnMenuRemoveDaily')
    .addSeparator()
    .addItem('Thời tiết: cập nhật ngay', 'tqnMenuWeatherNow')
    .addItem('Thời tiết: bật 06:00 hằng ngày', 'tqnMenuWeatherOn')
    .addItem('Thời tiết: tắt', 'tqnMenuWeatherOff')
    .addSeparator()
    .addItem('Cài đặt (tạo TQN_CONFIG)', 'tqnMenuSetup')
    .addToUi();
}

/** Chạy nhập + kiểm, có khóa chống 2 người bấm cùng lúc. Trả số liệu cho sidebar. */
function tqnRunAll_() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) throw new Error('Đang có người chạy nhập/kiểm, thử lại sau ít phút.');
  try {
    tqnEnsureConfig_();
    var cfg = tqnConfig_(), imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), rem = tqnRemindList_(cfg, imp, new Date());
    tqnWriteRemind_(cfg, rem);
    tqnAudit_('NHAP_KIEM', 'Chạy tay', { du_lieu: imp.counts, qc: qc.byRule, nhac: rem.length });
    return { cfg: cfg, imp: imp, qc: qc, rem: rem };
  } finally { lock.releaseLock(); }
}

function tqnMenuRun() {
  var r = tqnRunAll_(), c = r.imp.counts;
  var loi = r.qc.rows.filter(function (x) { return x[0] === 'LỖI'; }).length;
  var cb = r.qc.rows.filter(function (x) { return x[0] === 'CẢNH BÁO'; }).length;
  SpreadsheetApp.getUi().alert('TQN Register',
    'Đã nhập (chỉ đọc tab gốc):\n' +
    '• Sổ chi tiết: ' + c.chi_tiet + ' hồ sơ, ghép được với Sổ trình duyệt: ' + c.ghep_duoc + '\n' +
    '• Sổ trình duyệt: ' + c.so_tong + ' dòng (' + c.chi_co_so_tong + ' dòng chỉ có ở sổ tổng)\n' +
    '• TQN_DATA: ' + c.tong_dong + ' dòng\n\n' +
    'Kiểm lỗi: ' + loi + ' lỗi, ' + cb + ' cảnh báo → xem tab TQN_QC\n' +
    'Nhắc hạn: ' + r.rem.length + ' hồ sơ → tab TQN_NHAC_HAN', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRemindPreview() {
  var r = tqnRunAll_();
  tqnSs_().setActiveSheet(tqnSs_().getSheetByName(TQN.SH_REMIND));
  SpreadsheetApp.getUi().alert('Nhắc hạn hôm nay: ' + r.rem.length + ' hồ sơ. Chưa gửi mail (chỉ xem).');
}
function tqnMenuInstallDaily() {
  var h = tqnInstallDaily();
  SpreadsheetApp.getUi().alert('Đã bật nhắc hạn hằng ngày khoảng ' + h + 'h. Google có thể chạy lệch ±15 phút.\nGửi mail chỉ khi TQN_CONFIG: remind_email_enabled = Có và có remind_recipients.');
}
function tqnMenuRemoveDaily() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveDaily(false) + ' lịch nhắc hạn.'); }
function tqnMenuSetup() { tqnEnsureConfig_(); tqnEnsureWeatherConfig_(); tqnSs_().setActiveSheet(tqnSs_().getSheetByName(TQN.SH_CONFIG)); }

/* ---------------- Thanh bên ---------------- */
function tqnShowSidebar() {
  var t = HtmlService.createTemplateFromFile('Sidebar');
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('TQN Register'));
}

/** Dữ liệu cho thanh bên (gọi từ google.script.run). */
function tqnSidebarData() {
  var r = tqnRunAll_(), cfg = r.cfg, tz = cfg.timezone;
  var log = r.imp.src.log, cnt = function (f) { return log.filter(f).length; };
  var st = function (L) { return tqnNoAccent_(L.status); };
  var rules = r.qc.byRule;
  return {
    project: cfg.project_name, version: TQN.VERSION, at: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'),
    kpi: {
      tong: log.length,  // số dòng Sổ trình duyệt
      cho: cnt(function (L) { return st(L) === 'dang cho'; }),
      quaHan: cnt(function (L) { return st(L) === 'qua han'; }),
      duyet: cnt(function (L) { return /da duyet|duyet co y kien/.test(st(L)); }),
      loi: r.qc.rows.filter(function (x) { return x[0] === 'LỖI'; }).length,
      canhBao: r.qc.rows.filter(function (x) { return x[0] === 'CẢNH BÁO'; }).length,
      chiTiet: r.imp.counts.chi_tiet, ghep: r.imp.counts.ghep_duoc
    },
    remindTotal: r.rem.length,
    remind: r.rem.slice(0, 12).map(function (x) { return { g: x[0], han: tqnFmtDate_(x[1], tz), d: x[2], ma: x[4] || x[3], nd: x[6], at: x[9] }; }),
    qc: tqnGroupQC_(r.qc.rows.filter(function (x) { return x[0] !== 'THÔNG TIN'; })).slice(0, 15),
    qcTotal: r.qc.rows.length,
    rules: rules
  };
}

/** Nhảy tới ô "Tab!Dòng" khi bấm trong thanh bên. */
function tqnGoto(at) {
  var m = /^(.+)!([A-Z]*)(\d+)/.exec(String(at).split(',')[0].split(' ↔ ')[0].trim());
  if (!m) return;
  var sh = tqnSs_().getSheetByName(m[1]); if (!sh) return;
  sh.activate(); sh.setActiveRange(sh.getRange(+m[3], 1, 1, Math.min(sh.getLastColumn(), 20)));
}

/** Gom lỗi cùng luật + cùng nội dung cho thanh bên (vd 7 dòng khoảng trắng M0S → 1 mục). */
function tqnGroupQC_(rows) {
  var g = {}, order = [];
  rows.forEach(function (x) {
    var k = x[0] + '|' + x[1] + '|' + (x[1] === 'Q01_TRUNG_MA' || x[1] === 'Q09_CONG_THUC' || x[1] === 'Q08_Y_KIEN_LAP' ? x[2] : x[4]);
    if (!g[k]) { g[k] = { lv: x[0], rule: x[1], at: x[2], msg: x[4], n: 0, pos: [] }; order.push(k); }
    g[k].n++; g[k].pos.push(x[2]);
  });
  return order.map(function (k) { var o = g[k]; o.where = o.pos.join(', '); delete o.pos; return o; });
}
