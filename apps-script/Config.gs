/**
 * TQN Register Universal – Apps Script S1
 * Config.gs: cấu hình theo dự án. Mọi tham số nằm ở tab TQN_CONFIG (cột A khóa, B giá trị),
 * không sửa code khi tạo dự án mới.
 */
var TQN = {
  VERSION: 'S1-0.1.0',
  SH_CONFIG: 'TQN_CONFIG',
  SH_DATA: 'TQN_DATA',
  SH_QC: 'TQN_QC',
  SH_REMIND: 'TQN_NHAC_HAN',
  SH_AUDIT: 'TQN_AUDIT',
  PREFIX: 'TQN_'           // chỉ được ghi vào tab có tiền tố này
};

/** Giá trị mặc định cho profile COASTAL_2026_V1. Ghi vào TQN_CONFIG lần đầu cài. */
var TQN_DEFAULT_CONFIG = [
  ['project_id', 'COASTAL_QNG', 'Mã dự án bất biến, không trùng giữa các dự án'],
  ['project_name', 'Khu đô thị sinh thái Coastal Quảng Ngãi', ''],
  ['code_profile', 'COASTAL_2026_V1', 'Bộ quy tắc mã hồ sơ'],
  ['timezone', 'Asia/Ho_Chi_Minh', 'Múi giờ IANA của công trường'],
  ['log_sheet', 'Sổ trình duyệt', 'Sổ tổng có Kết quả / Hạn / Tình trạng'],
  ['detail_sheets', 'SD,M0S,MAT,OTS,SCH,RFA,RFI,RFC,CRP- Báo cáo,Hồ sơ an toàn,TRN-GNHS', 'Sổ chi tiết kiểu trình duyệt (cột B = mã)'],
  ['type_alias', 'SD=SDS;M0S=MSS;MS=MSS;MAT=MAS;OTS=OTS;SCH=SCH;RFA=RFA;RFI=RFI;RFC=RFC;TRN-GNHS=TML;CRP- Báo cáo=REP;Hồ sơ an toàn=HSE', 'Tên tab → loại chuẩn. Tab gốc không đổi tên'],
  ['status_submitted', 'đã trình duyệt;đang chờ;gửi email', 'Chữ trạng thái (sổ chi tiết) coi như đã trình'],
  ['status_closed', 'đã duyệt;duyệt có ý kiến;không duyệt;đã phản hồi;hủy', 'Chữ trạng thái coi như đã có phản hồi'],
  ['remind_days_ahead', '2', 'Nhắc trước hạn bao nhiêu ngày'],
  ['remind_hour', '7', 'Giờ chạy nhắc hạn hằng ngày (giờ dự án). Apps Script có thể lệch ±15 phút'],
  ['remind_email_enabled', 'Không', 'Có = gửi mail; Không = chỉ ghi tab TQN_NHAC_HAN'],
  ['remind_recipients', '', 'Email nhận nhắc hạn, cách nhau dấu phẩy'],
  ['han_mac_dinh', '4', 'Số ngày chờ phản hồi để tự tính hạn khi Sổ trình duyệt không có hạn (ngày trình + số ngày)'],
  ['detail_first_row', '6', 'Dòng dữ liệu đầu tiên của sổ chi tiết'],
  ['header_rows', '1-5', 'Vùng dòng tiêu đề để dò cột']
];

function tqnSs_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Đọc cấu hình thành object; thiếu khóa thì dùng mặc định. */
function tqnConfig_() {
  var cfg = {};
  TQN_DEFAULT_CONFIG.forEach(function (r) { cfg[r[0]] = String(r[1]); });
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k) cfg[k] = String(r[1] === null || r[1] === undefined ? '' : r[1]).trim();
    });
  }
  cfg.detailSheets = tqnSplit_(cfg.detail_sheets, ',');
  cfg.alias = {};
  tqnSplit_(cfg.type_alias, ';').forEach(function (p) {
    var i = p.indexOf('=');
    if (i > 0) cfg.alias[p.slice(0, i).trim()] = p.slice(i + 1).trim();
  });
  cfg.statusSubmitted = tqnSplit_(cfg.status_submitted, ';').map(tqnLower_);
  cfg.statusClosed = tqnSplit_(cfg.status_closed, ';').map(tqnLower_);
  cfg.remindDays = parseInt(cfg.remind_days_ahead, 10) || 2;
  cfg.remindHour = parseInt(cfg.remind_hour, 10) || 7;
  cfg.firstRow = parseInt(cfg.detail_first_row, 10) || 6;
  var hr = String(cfg.header_rows).split('-');
  cfg.hdrFrom = parseInt(hr[0], 10) || 1; cfg.hdrTo = parseInt(hr[1], 10) || 5;
  cfg.emailOn = /^c|^y|^1/i.test(cfg.remind_email_enabled);
  return cfg;
}

/** Tạo tab TQN_CONFIG nếu chưa có (không ghi đè giá trị anh đã sửa). */
function tqnEnsureConfig_() {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN.SH_CONFIG);
  if (sh) return sh;
  sh = ss.insertSheet(TQN.SH_CONFIG);
  sh.getRange(1, 1, 1, 3).setValues([['Khóa', 'Giá trị', 'Ghi chú']]).setFontWeight('bold').setBackground('#1B3F7A').setFontColor('#ffffff');
  sh.getRange(2, 1, TQN_DEFAULT_CONFIG.length, 3).setValues(TQN_DEFAULT_CONFIG);
  sh.getRange(2, 2, TQN_DEFAULT_CONFIG.length, 1).setBackground('#FFF2CC');
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 380); sh.setColumnWidth(3, 460);
  sh.setFrozenRows(1);
  return sh;
}

/* ------------ tiện ích chung ------------ */
function tqnSplit_(s, sep) { return String(s || '').split(sep).map(function (x) { return x.trim(); }).filter(String); }
function tqnLower_(s) { return String(s || '').toLowerCase().trim(); }

/** Làm sạch chữ: tab, khoảng trắng cứng, xuống dòng, khoảng trắng kép, khoảng trắng quanh '-'. */
function tqnClean_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  return String(v).replace(/[\t \r\n]/g, ' ').replace(/ {2,}/g, ' ').trim();
}
function tqnCleanCode_(v) { return String(tqnClean_(v)).replace(/\s*-\s*/g, '-'); }

/** Khóa so khớp hồ sơ: bỏ tiền tố ngày yymmdd- và đuôi rev (giữ nguyên số hiệu gốc ở cột khác). */
function tqnDocKey_(no) {
  var s = tqnCleanCode_(no).replace(/\.pdf$/i, '');
  if (/^\d{6}-/.test(s)) s = s.slice(7);
  s = s.replace(/[-_ ]*rev ?\d{1,3}\b.*$/i, '');
  return s.toUpperCase().trim();
}
function tqnRevFromNo_(no) { var m = /rev ?(\d{1,3})\b/i.exec(String(no || '')); return m ? ('0' + parseInt(m[1], 10)).slice(-2) : ''; }
function tqnRev_(v, no) {
  var s = String(tqnClean_(v)).toUpperCase().replace(/^REV/, '').trim();
  if (!s) return tqnRevFromNo_(no);
  if (/^\d{1,3}(\.0+)?$/.test(s)) return ('0' + parseInt(s, 10)).slice(-2);
  return s;
}
function tqnIsDate_(v) { return v instanceof Date && !isNaN(v.getTime()); }
function tqnFmtDate_(d, tz) { return tqnIsDate_(d) ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : ''; }

/** Chuẩn hóa chữ để so (bỏ dấu tiếng Việt). */
function tqnNoAccent_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
}

/** Ghi nhật ký TQN_AUDIT (append-only). */
function tqnAudit_(action, detail, counts) {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN.SH_AUDIT);
  if (!sh) {
    sh = ss.insertSheet(TQN.SH_AUDIT);
    sh.appendRow(['Thời điểm', 'Người chạy', 'Việc', 'Chi tiết', 'Số liệu', 'Phiên bản']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { user = ''; }
  sh.appendRow([new Date(), user, action, detail, JSON.stringify(counts || {}), TQN.VERSION]);
}

/** Chặn mọi lệnh ghi vào tab gốc: chỉ trả về tab có tiền tố TQN_. */
function tqnOwnSheet_(name, headers) {
  if (name.indexOf(TQN.PREFIX) !== 0) throw new Error('TQN không được ghi vào tab gốc: ' + name);
  var ss = tqnSs_(), sh = ss.getSheetByName(name);
  if (!sh) {
    var dangMo = null; try { dangMo = ss.getActiveSheet(); } catch (e) { }
    sh = ss.insertSheet(name);
    try { if (dangMo && dangMo.getName() !== name) ss.setActiveSheet(dangMo); } catch (e) { }   // tạo tab phụ không làm nhảy khỏi tab đang làm
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1B3F7A').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}
