/**
 * HoanTat.gs – bước cuối cho giống bản thử (19/11/2026).
 * - Bật thời tiết: weather_enabled = Có, weather_license_ok = Có (anh Nho yêu cầu 14:15), cài lịch 06:00, chạy thử 1 lần.
 * - Ẩn các tab không liên quan register (chỉ ẨN, không xóa; hiện lại bằng tqnHienTatCaTab).
 */
var TQN_TAB_AN = ['Lịch họp', 'Cập nhật kế hoạch công việc', 'DS BCH', 'DS Trực CN T4', 'DS Trực CN T7', 'Data', 'Check'];

function tqnDatCfg_(key, val) {
  var sh = tqnEnsureConfig_(), n = sh.getLastRow();
  if (n >= 2) {
    var ks = sh.getRange(2, 1, n - 1, 1).getValues();
    for (var i = 0; i < ks.length; i++) if (String(ks[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(val); return; }
  }
  sh.appendRow([key, val, '']);
}

function tqnBatThoiTiet() {
  tqnEnsureWeatherConfig_();
  tqnDatCfg_('weather_enabled', 'Có');
  tqnDatCfg_('weather_license_ok', 'Có');
  tqnDatCfg_('weather_hours', '6 11 13 18');
  var c0 = tqnWeatherCfg_();
  if (!(Math.abs(c0.lat) <= 90) || !(Math.abs(c0.lon) <= 180)) { tqnDatCfg_('weather_lat', 15.099038); tqnDatCfg_('weather_lon', 108.894552); tqnDatCfg_('weather_gps_confirmed', 'Có'); }
  var why = tqnWeatherBlocked_(tqnWeatherCfg_());
  if (why) return 'CHAN: ' + why;
  var gio = tqnInstallWeather();
  var r = tqnWeatherJob(true);
  return 'Lịch ' + gio + ' · chạy thử: ' + r;
}

function tqnAnTabPhu() {
  var ss = tqnSs_(), an = [];
  var dangMo = ss.getActiveSheet();
  ss.getSheets().forEach(function (sh) {
    var n = sh.getName().trim();
    var phu = n.indexOf('TQN_') === 0 || TQN_TAB_AN.some(function (t) { return t.trim() === n; });
    if (!phu || sh.isSheetHidden()) return;
    if (dangMo && dangMo.getSheetId() === sh.getSheetId()) {
      var td = ss.getSheetByName('Sổ trình duyệt'); if (td) ss.setActiveSheet(td);
    }
    try { sh.hideSheet(); an.push(n); } catch (e) { }
  });
  PropertiesService.getDocumentProperties().setProperty('tqn_tab_an', JSON.stringify(an));
  return an;
}

function tqnHienTatCaTab() {
  var ss = tqnSs_(), n = 0;
  ss.getSheets().forEach(function (sh) { if (sh.isSheetHidden()) { sh.showSheet(); n++; } });
  return n;
}

function tqnHoanTatCaiDat() {
  var kq = [];
  try { kq.push('Thời tiết: ' + tqnBatThoiTiet()); } catch (e) { kq.push('Thời tiết LỖI: ' + e.message); }
  try { var an = tqnAnTabPhu(); kq.push('Đã ẩn ' + an.length + ' tab: ' + an.join(', ')); } catch (e) { kq.push('Ẩn tab LỖI: ' + e.message); }
  try { kq.push('Tự mở bảng: ' + tqnBatMoBangKhiMo()); } catch (e) { kq.push('Tự mở bảng LỖI: ' + e.message); }
  try { kq.push('Trang chủ: ' + tqnTaoTrangChu() + ' nút'); } catch (e) { kq.push('Trang chủ LỖI: ' + e.message); }
  try { kq.push('Tự cấp mã: ' + tqnBatTuDongMa().join(', ')); } catch (e) { kq.push('Tự cấp mã LỖI: ' + e.message); }
  Logger.log(kq.join('\n'));
  return kq.join('\n');
}

/** Bắt Google hỏi lại các quyền còn thiếu (thời tiết, AI cần quyền kết nối Internet). */
function tqnXinQuyen() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  Logger.log('Đủ quyền. Thử thời tiết: ' + tqnWeatherJob(true));
}

/** Chạy lại thời tiết từ 01/11/2026 đến hôm nay (anh Nho yêu cầu 14:33 19/11). */
function tqnThoiTietTu0111() { Logger.log(tqnWeatherKhoang('01/11/2026', Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy'))); }

/** Chạy 1 lần trong trình soạn: bật tự cấp mã + ghi log các dòng đang thiếu mã (KHÔNG ghi vào sổ). */
function tqnCaiTuDongMa() {
  var tabs = tqnBatTuDongMa(), thieu = tqnDongThieuMa_(false);
  var s = 'Bật tự cấp mã: ' + tabs.join(', ') + ' | Dòng thiếu mã: ' + thieu.length + (thieu.length ? ' → ' + thieu.map(function (x) { return x.tab + '!' + x.row + '=' + x.code; }).join('; ') : '');
  Logger.log(s); return s;
}

/** Chạy 1 lần trong trình soạn: tạo thư mục thả file + lịch 15 phút, ghi link thư mục ra log. */
function tqnCaiThaFile() { var u = tqnBatThaFile(); Logger.log('Thư mục thả file: ' + u); return u; }

/** Chạy trong trình soạn: tạo / làm mới tab Tiến độ ↔ hồ sơ chặn. */
function tqnCaiTienDo() { var k = tqnTaoSheetTienDo(); Logger.log('Tab tiến độ: ' + k.tong + ' dòng, thêm mới ' + k.them); return k; }
