/**
 * Weather.gs – 06:00 mỗi ngày ghi 1 dòng dự báo vào tab "Theo dõi thời tiết" của Register.
 * - Chỉ chạy khi: weather_enabled = Có, có GPS đã xác nhận, đã xác nhận điều khoản nguồn dữ liệu.
 * - Không ghi đè dòng người nhập tay; chỉ cập nhật dòng do TQN ghi (Ghi chú bắt đầu "[TQN]").
 * - Lỗi API → không ghi số giả, ghi nhật ký FAILED.
 */
var TQN_WEATHER_DEFAULTS = [
  ['weather_enabled', 'Không', 'Có = bật cập nhật thời tiết 06:00'],
  ['weather_lat', '15.099038', 'Vĩ độ công trường (vd 15.0xxxx). Phải là tọa độ thật đã xác nhận'],
  ['weather_lon', '108.894552', 'Kinh độ công trường (vd 108.8xxxx)'],
  ['weather_gps_confirmed', 'Có', 'Có = người quản lý đã xác nhận tọa độ đúng công trường (anh Nho gửi 11:12 19/11/2026)'],
  ['weather_sheet', 'Theo dõi thời tiết', 'Tab ghi thời tiết (giữ đúng format cột của Register)'],
  ['weather_hour', '6', 'Giờ chạy (giờ dự án). Apps Script có thể lệch ±15 phút'],
  ['weather_hours', '6 11 13 18', 'Các giờ cập nhật trong ngày (sáng/trưa/chiều/tối, giống bản tin). Mỗi lần cập nhật lại đúng dòng hôm nay'],
  ['weather_provider', 'OPEN_METEO', 'Nguồn dữ liệu'],
  ['weather_license_ok', 'Không', 'Có = đã kiểm điều khoản nguồn (Open-Meteo free chỉ cho phi thương mại; dùng cho công ty cần gói trả phí)'],
  ['weather_api_key', '', 'Khóa gói trả phí nếu có (để trống với gói miễn phí). Nên đặt ở Script Properties thay vì ô này']
];

var TQN_WMO_VI = {
  0: 'Trời quang', 1: 'Ít mây', 2: 'Có mây', 3: 'Nhiều mây', 45: 'Sương mù', 48: 'Sương mù đóng băng',
  51: 'Mưa phùn nhẹ', 53: 'Mưa phùn', 55: 'Mưa phùn dày', 56: 'Mưa phùn lạnh', 57: 'Mưa phùn lạnh dày',
  61: 'Mưa nhẹ', 63: 'Mưa vừa', 65: 'Mưa to', 66: 'Mưa lạnh nhẹ', 67: 'Mưa lạnh to',
  71: 'Tuyết nhẹ', 73: 'Tuyết vừa', 75: 'Tuyết dày', 77: 'Hạt tuyết',
  80: 'Mưa rào nhẹ', 81: 'Mưa rào', 82: 'Mưa rào rất to', 85: 'Mưa tuyết nhẹ', 86: 'Mưa tuyết to',
  95: 'Dông', 96: 'Dông kèm mưa đá nhẹ', 99: 'Dông kèm mưa đá'
};
var TQN_THU = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];

function tqnWeatherCfg_() {
  var cfg = tqnConfig_();
  TQN_WEATHER_DEFAULTS.forEach(function (r) { if (cfg[r[0]] === undefined) cfg[r[0]] = r[1]; });
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG);
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (String(r[0]).indexOf('weather_') === 0) cfg[String(r[0]).trim()] = String(r[1]).trim(); });
  var yes = function (v) { return /^c|^y|^1/i.test(String(v || '')); };
  cfg.wOn = yes(cfg.weather_enabled); cfg.wGps = yes(cfg.weather_gps_confirmed); cfg.wLic = yes(cfg.weather_license_ok);
  cfg.lat = parseFloat(String(cfg.weather_lat).replace(',', '.')); cfg.lon = parseFloat(String(cfg.weather_lon).replace(',', '.'));
  cfg.wHour = parseInt(cfg.weather_hour, 10) || 6;
  return cfg;
}

/** Thêm các khóa weather_* vào TQN_CONFIG nếu chưa có (không đụng giá trị anh đã sửa). */
function tqnEnsureWeatherConfig_() {
  var sh = tqnEnsureConfig_(), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_WEATHER_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}

/** Kiểm điều kiện trước khi gọi API. Trả về chuỗi lỗi hoặc ''. */
function tqnWeatherBlocked_(cfg) {
  if (!cfg.wOn) return 'weather_enabled chưa bật (Có)';
  if (isNaN(cfg.lat) || isNaN(cfg.lon) || Math.abs(cfg.lat) > 90 || Math.abs(cfg.lon) > 180) return 'Chưa có tọa độ hợp lệ (weather_lat / weather_lon)';
  if (!cfg.wGps) return 'Tọa độ chưa được xác nhận (weather_gps_confirmed = Có)';
  if (!cfg.wLic) return 'Chưa xác nhận điều khoản nguồn dữ liệu (weather_license_ok = Có)';
  return '';
}

function tqnWeatherUrl_(cfg, khoang) {
  var base = cfg.weather_api_key ? 'https://customer-api.open-meteo.com/v1/forecast' : 'https://api.open-meteo.com/v1/forecast';
  var q = {
    latitude: cfg.lat, longitude: cfg.lon, timezone: cfg.timezone, wind_speed_unit: 'kmh',
    daily: 'weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,wind_gusts_10m_max,sunrise,sunset',
    hourly: 'relative_humidity_2m,temperature_2m,apparent_temperature,precipitation_probability,precipitation,weather_code,wind_gusts_10m,wind_speed_10m'
  };
  if (khoang) { q.start_date = khoang.tu; q.end_date = khoang.den; } else q.forecast_days = 2;
  if (cfg.weather_api_key) q.apikey = cfg.weather_api_key;
  return base + '?' + Object.keys(q).map(function (k) { return k + '=' + encodeURIComponent(q[k]); }).join('&');
}

/** Đọc số liệu của ngày thứ k trong phản hồi (k=0 là ngày đầu). */
function tqnWeatherDay_(j, k) {
  var d = j.daily || {}, h = j.hourly || {}, at = function (a, i) { return a && a[i] !== undefined ? a[i] : null; };
  var hh = {}; Object.keys(h).forEach(function (key) { hh[key] = (h[key] || []).slice(k * 24, k * 24 + 24); });
  var hum = (hh.relative_humidity_2m || []).filter(function (x) { return typeof x === 'number'; });
  var feel = (hh.apparent_temperature || []).filter(function (x) { return typeof x === 'number'; });
  return {
    date: d.time[k], tmin: at(d.temperature_2m_min, k), tmax: at(d.temperature_2m_max, k), code: at(d.weather_code, k),
    rain: at(d.precipitation_sum, k), pop: at(d.precipitation_probability_max, k),
    wind: at(d.wind_speed_10m_max, k), gust: at(d.wind_gusts_10m_max, k),
    hum: hum.length ? Math.round(hum.reduce(function (a, b) { return a + b; }, 0) / hum.length) : null,
    feel: feel.length ? Math.max.apply(null, feel) : null,
    buoi: tqnWeatherBuoi_(hh), gio: hh,
    mai: (d.time.length > k + 1 && at(d.temperature_2m_max, k + 1) !== null) ? { date: d.time[k + 1], tmin: d.temperature_2m_min[k + 1], tmax: d.temperature_2m_max[k + 1], code: d.weather_code[k + 1], rain: at(d.precipitation_sum, k + 1) } : null,
    sunrise: d.sunrise ? String(d.sunrise[k] || '').slice(11, 16) : '', sunset: d.sunset ? String(d.sunset[k] || '').slice(11, 16) : ''
  };
}

/** Gọi API có thử lại; trả về {ok, data (ngày đầu), j, raw, fetchedAt}. khoang = {tu,den} 'yyyy-MM-dd' để lấy lại ngày đã qua. */
function tqnWeatherFetch_(cfg, khoang) {
  var url = tqnWeatherUrl_(cfg, khoang), last = '';
  for (var i = 0; i < 3; i++) {
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = res.getResponseCode(), raw = res.getContentText();
      if (code === 200) {
        var j = JSON.parse(raw), d = j.daily || {};
        var need = ['time', 'temperature_2m_min', 'temperature_2m_max', 'weather_code'];
        for (var n = 0; n < need.length; n++) if (!d[need[n]] || d[need[n]][0] === undefined || d[need[n]][0] === null) throw new Error('Thiếu trường ' + need[n]);
        return { ok: true, raw: raw, j: j, fetchedAt: new Date(), data: tqnWeatherDay_(j, 0) };
      }
      last = 'HTTP ' + code;
      if (code >= 400 && code < 500 && code !== 429) break;   // lỗi tham số: thử lại vô ích
    } catch (e) { last = String(e.message || e); }
    Utilities.sleep(Math.pow(2, i) * 2000);
  }
  return { ok: false, err: last };
}

/** Chia 24 giờ hôm nay thành 3 buổi giống bản tin: Sáng 06–11, Chiều 12–17, Tối 18–23. */
var TQN_BUOI = [['Sáng', 6, 10], ['Trưa', 11, 13], ['Chiều', 14, 17], ['Tối', 18, 23]];
function tqnWeatherBuoi_(h) {
  var arr = function (k) { return (h[k] || []).slice(0, 24); };   // h đã là 24 giờ của đúng ngày
  var t = arr('temperature_2m'), pp = arr('precipitation_probability'), pr = arr('precipitation'), wc = arr('weather_code'), g = arr('wind_gusts_10m');
  if (!t.length) return [];
  var num = function (x) { return typeof x === 'number'; };
  return TQN_BUOI.map(function (b) {
    var r = function (a) { return a.slice(b[1], b[2] + 1).filter(num); };
    var tt = r(t), p = r(pp), m = r(pr), c = r(wc), gg = r(g);
    if (!tt.length) return null;
    var dem = {}, chinh = c[0];
    c.forEach(function (x) { dem[x] = (dem[x] || 0) + 1; if (dem[x] > (dem[chinh] || 0) || (dem[x] === dem[chinh] && x > chinh)) chinh = x; });
    return { ten: b[0], tu: b[1], den: b[2], tmin: Math.min.apply(null, tt), tmax: Math.max.apply(null, tt),
      pop: p.length ? Math.max.apply(null, p) : null, mua: m.length ? Math.round(m.reduce(function (a, x) { return a + x; }, 0) * 10) / 10 : null,
      code: chinh, gust: gg.length ? Math.max.apply(null, gg) : null };
  }).filter(Boolean);
}
function tqnSoVi_(x) { return String(Math.round(x * 10) / 10).replace('.', ','); }
/** Ghi chú dễ đọc, mỗi buổi một dòng. Dòng cuối ghi "Tự động" để lần sau nhận ra dòng do tool ghi. */
function tqnWeatherNote_(cfg, w) {
  var d = w.data, L = [];
  (d.buoi || []).forEach(function (b) {
    L.push(b.ten + ' ' + Math.round(b.tmin) + '–' + Math.round(b.tmax) + '°: ' + (TQN_WMO_VI[b.code] || 'mã ' + b.code).toLowerCase() +
      (b.pop !== null ? ', khả năng mưa ' + b.pop + '%' : '') + (b.mua ? ' (' + tqnSoVi_(b.mua) + ' mm)' : ''));
  });
  var x = [];
  if (d.gust !== null && d.gust !== undefined) x.push('Gió giật ' + Math.round(d.gust) + ' km/h');
  if (d.sunrise) x.push('mặt trời mọc ' + d.sunrise + ', lặn ' + d.sunset);
  if (x.length) L.push(x.join(' · '));
  if (d.mai) { var p = d.mai.date.split('-'); L.push('Ngày mai ' + p[2] + '/' + p[1] + ': ' + (TQN_WMO_VI[d.mai.code] || '').toLowerCase() + ' ' + Math.round(d.mai.tmin) + '–' + Math.round(d.mai.tmax) + '°' + (d.mai.rain ? ', mưa ' + tqnSoVi_(d.mai.rain) + ' mm' : '')); }
  if (!(d.buoi || []).length && d.pop !== null) L.unshift('Khả năng mưa ' + d.pop + '%');
  return L.join('\n');
}

/** Tab thời tiết: chỉ cho phép đúng tên trong cấu hình và đúng tiêu đề, để không ghi nhầm tab khác. */
function tqnWeatherSheet_(cfg) {
  var sh = tqnSs_().getSheetByName(cfg.weather_sheet);
  if (!sh) throw new Error('Không thấy tab ' + cfg.weather_sheet);
  var hdr = sh.getRange(1, 1, 1, Math.max(11, sh.getLastColumn())).getValues()[0].map(function (x) { return tqnNoAccent_(tqnClean_(x)); });
  var col = function (k) { var key = tqnNoAccent_(k); for (var i = 0; i < hdr.length; i++) if (hdr[i].indexOf(key) === 0) return i + 1; return 0; };
  var m = { date: col('Ngày'), thu: col('Thứ'), tmin: col('Nhiệt độ thấp nhất'), tmax: col('Nhiệt độ cao nhất'), wx: col('Thời tiết'), rain: col('Lượng mưa'), hum: col('Độ ẩm'), wind: col('Tốc độ gió'), note: col('Ghi chú') };
  if (!m.date || !m.tmin || !m.tmax || !m.note) throw new Error('Tiêu đề tab ' + cfg.weather_sheet + ' không đúng format (cần Ngày / Nhiệt độ thấp nhất / Nhiệt độ cao nhất / Ghi chú)');
  return { sh: sh, m: m };
}

/** Ghi / cập nhật dòng của ngày. Trả 'THEM' | 'CAP_NHAT' | 'BO_QUA_NHAP_TAY'. */
function tqnWeatherWrite_(cfg, w) {
  var t = tqnWeatherSheet_(cfg), sh = t.sh, m = t.m;
  var p = w.data.date.split('-'), day = new Date(+p[0], +p[1] - 1, +p[2]);
  var last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var dates = sh.getRange(2, m.date, last - 1, 1).getValues(), notes = sh.getRange(2, m.note, last - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      var dv = dates[i][0];
      if (tqnIsDate_(dv) && dv.getFullYear() === day.getFullYear() && dv.getMonth() === day.getMonth() && dv.getDate() === day.getDate()) {
        if (!/^\[TQN\]|Tự động – |^📷 /.test(String(notes[i][0]))) return 'BO_QUA_NHAP_TAY';
        row = i + 2; break;
      }
    }
  }
  var action = row ? 'CAP_NHAT' : 'THEM';
  if (!row) {
    row = Math.max(2, last + 1);
    if (last >= 2) {   // chèn đúng thứ tự ngày (lấy lại ngày cũ không bị dồn xuống cuối)
      for (var q = 0; q < dates.length; q++) if (tqnIsDate_(dates[q][0]) && dates[q][0] > day) { sh.insertRowBefore(q + 2); row = q + 2; break; }
    }
  }
  var r1 = function (x) { return (x === null || x === undefined) ? '' : Math.round(x * 10) / 10; };
  var d = w.data;
  var note = tqnWeatherNote_(cfg, w);
  var set = function (c, v) { if (c) sh.getRange(row, c).setValue(v); };
  // Ghi 1 lần cho cả dòng (nhanh hơn ~10 lần gọi): giữ nguyên giá trị các cột không thuộc tool
  var cMax = Math.max(m.date, m.thu, m.tmin, m.tmax, m.wx, m.rain, m.hum, m.wind, 1);
  var dong = sh.getRange(row, 1, 1, cMax).getValues()[0];
  var dat = function (c, v) { if (c) dong[c - 1] = v; };
  dat(m.date, day); dat(m.thu, TQN_THU[day.getDay()]); dat(m.tmin, r1(d.tmin)); dat(m.tmax, r1(d.tmax));
  dat(m.wx, TQN_WMO_VI[d.code] || ('Mã WMO ' + d.code)); dat(m.rain, r1(d.rain)); dat(m.hum, d.hum === null ? '' : d.hum); dat(m.wind, r1(d.wind));
  sh.getRange(row, 1, 1, cMax).setValues([dong]);
  set(m.note, note);
  var tom = (d.buoi || []).map(function (b) { return b.ten + ' ' + (b.pop === null ? '–' : b.pop + '%'); }).join(' · ');
  if (m.note && w.banTin) {
    try {
      var nhan = '📷 Xem bản tin hình (bấm mở)', full = nhan + '\n' + note;
      sh.getRange(row, m.note).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(full).setLinkUrl(0, nhan.length, w.banTin).build());
    } catch (e) { set(m.note, '📷 Bản tin: ' + w.banTin + '\n' + note); }
  } else if (m.note) set(m.note, '📷 Bản tin hình chưa tạo được\n' + note);
  try { if (m.note) sh.getRange(row, m.note).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); sh.getRange(row, 1, 1, m.note).setVerticalAlignment('middle'); if (m.note) sh.getRange(row, m.note).setWrap(true).setFontSize(9); sh.getRange(row, 1, 1, m.note).setVerticalAlignment('middle');
    if (sh.getColumnWidth(m.note) < 380) sh.setColumnWidth(m.note, 380);
    tqnWeatherIconCells_(sh, row, m.note + 1, d);
    sh.setRowHeightsForced(row, 1, 110); } catch (e) { tqnAudit_('THOI_TIET', 'Định dạng/hình lỗi: ' + e.message, {}); }
  if (m.date) sh.getRange(row, m.date).setNumberFormat('dd/MM/yyyy');
  return action;
}

/** Lưu bản gốc phản hồi API làm bằng chứng (tab riêng TQN_). */
function tqnWeatherRaw_(cfg, w, status) {
  var sh = tqnOwnSheet_('TQN_THOI_TIET_RAW', ['Lúc lấy', 'Dự án', 'Ngày dự báo', 'Trạng thái', 'Vĩ độ', 'Kinh độ', 'Phản hồi gốc (JSON)']);
  sh.appendRow([new Date(), cfg.project_id, w.ok ? w.data.date : '', status, cfg.lat, cfg.lon, w.ok ? String(w.raw).slice(0, 45000) : String(w.err)]);
}

/** Hàm cho trigger 06:00. manual=true khi bấm tay (vẫn chống trùng: cập nhật lại dòng TQN của hôm nay). */
function tqnWeatherJob(manual) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) { tqnAudit_('THOI_TIET', 'Bỏ qua: đang có lần chạy khác', {}); return 'BAN'; }
  try {
    var cfg = tqnWeatherCfg_(), why = tqnWeatherBlocked_(cfg);
    if (why) { tqnAudit_('THOI_TIET', 'Chặn: ' + why, {}); return 'CHAN: ' + why; }
    var gio = parseInt(Utilities.formatDate(new Date(), cfg.timezone, 'HH'), 10), slot = gio < 11 ? '06' : gio < 13 ? '11' : gio < 18 ? '13' : '18';
    var key = cfg.project_id + ':' + Utilities.formatDate(new Date(), cfg.timezone, 'yyyy-MM-dd') + ':weather' + slot;
    var props = PropertiesService.getDocumentProperties();
    if (manual !== true && props.getProperty('lastWeather') === key) { tqnAudit_('THOI_TIET', 'Bỏ qua: hôm nay đã chạy ' + key, {}); return 'DA_CHAY'; }
    var w = tqnWeatherFetch_(cfg);
    if (!w.ok) { tqnWeatherRaw_(cfg, w, 'FAILED'); tqnAudit_('THOI_TIET', 'FAILED: ' + w.err, {}); return 'FAILED: ' + w.err; }
    w.banTin = tqnWeatherBanTin_(cfg, w);
    var act = tqnWeatherWrite_(cfg, w);
    try { tqnDinhDangThoiTiet(); } catch (e) { }
    tqnWeatherRaw_(cfg, w, act);
    props.setProperty('lastWeather', key);
    tqnAudit_('THOI_TIET', act + ' ' + w.data.date, w.data);
    return act;
  } finally { lock.releaseLock(); }
}

function tqnInstallWeather() {
  var cfg = tqnWeatherCfg_();
  tqnRemoveWeather(true);
  var raw = String(cfg.weather_hours || '').trim();
  var gios = /^\d{1,2}([\s,;]+\d{1,2})*$/.test(raw) ? raw.split(/[\s,;]+/).map(function (x) { return parseInt(x, 10); }).filter(function (x) { return x >= 0 && x < 24; }) : [];
  if (!gios.length) gios = [6, 11, 13, 18];
  gios.forEach(function (h) { ScriptApp.newTrigger('tqnWeatherJob').timeBased().atHour(h).nearMinute(0).everyDays(1).inTimezone(cfg.timezone).create(); });
  tqnAudit_('BAT_THOI_TIET', 'Hằng ngày khoảng ' + gios.join('h, ') + 'h (' + cfg.timezone + ')', {});
  return gios.join('h, ') + 'h';
}
function tqnRemoveWeather(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnWeatherJob') { ScriptApp.deleteTrigger(t); n++; } });
  if (!silent) tqnAudit_('TAT_THOI_TIET', 'Đã xóa ' + n + ' trigger', {});
  return n;
}

/* menu */
function tqnWeatherMsg_(r) {
  r = String(r || '');
  if (/^CHAN: /.test(r)) return 'Chưa cập nhật được thời tiết.\nLý do: ' + r.slice(6) + '\n\nCách bật: mở tab TQN_CONFIG, đổi weather_enabled = Có và weather_license_ok = Có (sau khi chốt nguồn dữ liệu), rồi bấm lại.';
  if (r === 'THEM' || r === 'CAP_NHAT') return '✔ Đã ' + (r === 'THEM' ? 'thêm' : 'cập nhật') + ' dòng thời tiết hôm nay vào tab Theo dõi thời tiết.';
  if (r === 'BO_QUA_NHAP_TAY') return 'Dòng hôm nay đã có người nhập tay – giữ nguyên, không ghi đè.';
  if (r === 'DA_CHAY') return 'Hôm nay đã cập nhật thời tiết rồi.';
  if (/^FAILED/.test(r)) return 'Lấy dữ liệu thời tiết lỗi: ' + r.replace(/^FAILED:\s*/, '') + '\nĐã thử lại 3 lần. Bấm lại sau ít phút.';
  return r;
}
function tqnMenuWeatherNow() { tqnEnsureWeatherConfig_(); SpreadsheetApp.getUi().alert('Thời tiết', tqnWeatherMsg_(tqnWeatherJob(true)), SpreadsheetApp.getUi().ButtonSet.OK); }
function tqnMenuWeatherOn() {
  tqnEnsureWeatherConfig_();
  var why = tqnWeatherBlocked_(tqnWeatherCfg_());
  if (why) { SpreadsheetApp.getUi().alert('Chưa bật được: ' + why + '\nSửa ở tab TQN_CONFIG.'); return; }
  SpreadsheetApp.getUi().alert('Đã bật cập nhật thời tiết hằng ngày lúc ' + tqnInstallWeather() + ' (Google có thể lệch ±15 phút).');
}
function tqnMenuWeatherOff() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveWeather(false) + ' lịch thời tiết.'); }


/* ===== Bản tin hình (minh chứng) – giống mẫu bản tin TQN anh Nho gửi 19/11 ===== */
function tqnWxIcon_(c, dem) {
  if (c === null || c === undefined) return '';
  if (c >= 95) return '⛈'; if (c >= 51) return '🌧'; if (c >= 45) return '🌫'; if (c >= 2) return '☁'; return dem ? '☾' : '☀';
}
function tqnWeatherBanTinHtml_(cfg, w) {
  var d = w.data, p = d.date.split('-'), day = new Date(+p[0], +p[1] - 1, +p[2]), e = tqnEsc_;
  var g = d.gio || {}, T = g.temperature_2m || [], PP = g.precipitation_probability || [], WC = g.weather_code || [];
  var ic = function (c, dem, px) { return (c === null || c === undefined) ? '' : '<img src="' + TQN_WX_ICON[tqnWxIconKey_(c, dem)] + '" width="' + px + '" height="' + px + '">'; };
  var so = function (x) { return x === null || x === undefined ? '–' : Math.round(x); };
  // Câu tóm tắt như app điện thoại
  var bMax = null; (d.buoi || []).forEach(function (b) { if (b.pop !== null && (!bMax || b.pop > bMax.pop)) bMax = b; });
  var tom = bMax && bMax.pop >= 40 ? (TQN_WMO_VI[bMax.code] || 'Có mưa') + ' vào buổi ' + bMax.ten.toLowerCase() + '.' : 'Ít khả năng mưa, ' + (TQN_WMO_VI[d.code] || '').toLowerCase() + '.';
  tom += ' Cao ' + so(d.tmax) + '° và thấp ' + so(d.tmin) + '°C.' + (d.gust ? ' Gió giật đến ' + so(d.gust) + ' km/h.' : '');
  var khoi = function (ten, a, b) {
    var h1 = '', h2 = '', h3 = '', h4 = '';
    for (var i = a; i <= b; i++) {
      h1 += '<td class="hh">' + (i < 10 ? '0' : '') + i + ':00</td>';
      h2 += '<td class="ic">' + ic(WC[i], i >= 18 || i < 6, 34) + '</td>';
      h3 += '<td class="tt">' + so(T[i]) + '°</td>';
      h4 += '<td class="pp">💧' + (PP[i] === null || PP[i] === undefined ? '–' : PP[i] + '%') + '</td>';
    }
    return '<table class="card" width="100%" cellspacing="0" bgcolor="#6E9FB0" style="background:#6E9FB0"><tr><td colspan="' + (b - a + 1) + '" class="ct">' + ten + '</td></tr><tr>' + h1 + '</tr><tr>' + h2 + '</tr><tr>' + h3 + '</tr><tr>' + h4 + '</tr></table>';
  };
  var loai = w.quaKhu ? 'SỐ LIỆU NGÀY ĐÃ QUA' : 'DỰ BÁO';
  var mua = bMax ? '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td><div class="nho">💧 Mưa sắp tới</div><div class="vua">Khả năng mưa cao nhất: buổi ' + bMax.ten.toLowerCase() + ' (' + bMax.tu + ':00–' + bMax.den + ':00)</div></td><td align="right" class="lon">' + bMax.pop + '%</td></tr></table>' : '';
  var mai = d.mai ? '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td class="vua">Ngày mai ' + d.mai.date.slice(8, 10) + '/' + d.mai.date.slice(5, 7) + ' &nbsp;' + ic(d.mai.code, false, 26) + ' ' + e(TQN_WMO_VI[d.mai.code] || '') + (d.mai.rain ? ' · ' + tqnSoVi_(d.mai.rain) + ' mm' : '') + '</td><td align="right" class="vua"><b>' + so(d.mai.tmax) + '° ' + so(d.mai.tmin) + '°</b></td></tr></table>' : '';
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#0F172A;margin:0;padding:16px;background:#DCEBFA}' +
    '.top td{font-size:10px;font-weight:bold;color:#1E3A8A}.tag{background:#0F6E66;color:#fff;padding:2px 7px;border-radius:8px}.tag2{background:#B45309;color:#fff;padding:2px 7px;border-radius:8px}' +
    '.hero{color:#fff}.big{font-size:54px;font-weight:300;line-height:1}.cond{font-size:20px}.sub{font-size:12px;line-height:1.6}' +
    '.sum{color:#fff}' +
    '.card{background:#6E9FB0;color:#fff;border-radius:12px;margin-top:10px}.card td{text-align:center;padding:3px 0}.ct{text-align:left !important;padding:8px 12px 2px !important;font-weight:bold;font-size:12px}' +
    '.hh{font-size:11px;color:#E2E8F0}.tt{font-size:17px;font-weight:bold}.pp{font-size:10px;color:#E0F2FE}' +
    '.card2{background:#4F7FC0;color:#fff;border-radius:12px;margin-top:10px}.card2 td{padding:10px 14px}.nho{font-size:11px}.vua{font-size:13px}.lon{font-size:30px;font-weight:bold}' +
    '.ft{font-size:8.5px;color:#334155;margin-top:10px;line-height:1.5}.warn{background:#1E3A8A;color:#DBEAFE;padding:1px 5px;border-radius:3px}' +
    '</style></head><body bgcolor="#DCEBFA"><table width="100%" cellspacing="0" cellpadding="12" bgcolor="#DCEBFA" style="background:#DCEBFA"><tr><td>' +
    '<table class="top" width="100%"><tr><td>TQN · BẢN TIN THỜI TIẾT CÔNG TRƯỜNG</td><td align="right"><span class="tag">' + loai + '</span> <span class="tag2">TỰ ĐỘNG</span></td></tr></table>' +
    '<table width="100%" cellspacing="0" cellpadding="14" bgcolor="#5B9BE0" style="background:#5B9BE0;margin-top:8px"><tr><td class="hero" style="color:#fff"><div style="font-size:15px;font-weight:bold">' + e(cfg.project_name || cfg.project_id || '') + '</div>' +
    '<table width="100%"><tr><td><div class="big">' + so(d.tmax) + '°</div><div class="cond">' + e(TQN_WMO_VI[d.code] || '') + '</div>' +
    '<div class="sub">↑' + so(d.tmax) + '° / ↓' + so(d.tmin) + '°' + (d.feel !== null && d.feel !== undefined ? ' · Cảm giác như ' + so(d.feel) + '°' : '') + '<br>' + TQN_THU[day.getDay()] + ', ' + p[2] + '/' + p[1] + '/' + p[0] + (d.sunrise ? ' · Mặt trời mọc ' + d.sunrise + ', lặn ' + d.sunset : '') + '</div></td>' +
    '<td align="right" valign="top">' + ic(d.code, false, 90) + '</td></tr></table></td></tr></table>' +
    '<table width="100%" cellspacing="0" cellpadding="10" bgcolor="#4F86C6" style="background:#4F86C6;margin-top:10px"><tr><td class="sum" style="color:#fff;font-size:13px">' + e(tom) + '</td></tr></table>' +
    khoi('Sáng 06:00–11:00', 6, 11) + khoi('Chiều 12:00–17:00', 12, 17) + khoi('Tối 18:00–23:00', 18, 23) +
    mua + '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td class="vua">Lượng mưa cả ngày <b>' + (d.rain === null ? '–' : tqnSoVi_(d.rain) + ' mm') + '</b> · Độ ẩm <b>' + (d.hum === null ? '–' : d.hum + '%') + '</b> · Gió <b>' + so(d.wind) + ' km/h</b>' + (d.gust ? ', giật ' + so(d.gust) + ' km/h' : '') + '</td></tr></table>' + mai +
    '<div class="ft">TQN Register · ' + e(cfg.project_id || '') + ' · GPS ' + cfg.lat + ', ' + cfg.lon + ' · Nguồn Open-Meteo (open-meteo.com, CC BY 4.0) · ' + (w.quaKhu ? 'SỐ LIỆU MÔ HÌNH' : 'DỰ BÁO') + ', không phải số đo tại công trường · tạo lúc ' + Utilities.formatDate(w.fetchedAt, cfg.timezone, 'HH:mm dd/MM/yyyy') +
    '<br><span class="warn">Bản nội bộ – chưa xác minh giấy phép thương mại</span></div></td></tr></table></body></html>';
}
function tqnEsc_(x) { return String(x === null || x === undefined ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
/** Lưu bản tin PDF vào Drive (thư mục riêng), trả link. Lỗi thì trả '' – không chặn việc ghi Register. */
function tqnWeatherBanTin_(cfg, w) {
  try {
    if (typeof DriveApp === 'undefined') return '';
    var ten = cfg.weather_folder || 'TQN Register – Bản tin thời tiết';
    var it = DriveApp.getFoldersByName(ten), f = it.hasNext() ? it.next() : DriveApp.createFolder(ten);
    var name = 'ThoiTiet_' + w.data.date + (w.quaKhu ? '_ca-ngay' : '_' + Utilities.formatDate(w.fetchedAt, cfg.timezone, 'HH') + 'h');
    var cu = f.getFilesByName(name + '.pdf'); while (cu.hasNext()) cu.next().setTrashed(true);   // cập nhật lại thì thay bản cũ
    var blob = Utilities.newBlob(tqnWeatherBanTinHtml_(cfg, w), 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
    return f.createFile(blob).getUrl();
  } catch (e) { tqnAudit_('THOI_TIET', 'Không tạo được bản tin: ' + e.message, {}); return ''; }
}

/** Cập nhật thời tiết cho khoảng ngày (vd từ 01/11 đến hôm nay). tu/den: 'yyyy-MM-dd' hoặc 'dd/MM/yyyy'. */
function tqnWeatherKhoang(tu, den) {
  var cfg = tqnWeatherCfg_(), why = tqnWeatherBlocked_(cfg);
  if (why) return 'CHAN: ' + why;
  var chuan = function (x) { x = String(x || '').trim(); var m = x.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); if (/^\d{4}-\d{2}-\d{2}$/.test(x)) return x; throw new Error('Ngày không đúng dạng dd/mm/yyyy: ' + x); };
  tu = chuan(tu); den = chuan(den);
  var homNay = Utilities.formatDate(new Date(), cfg.timezone, 'yyyy-MM-dd');
  if (den > homNay) den = homNay;
  if (tu > den) throw new Error('Ngày bắt đầu sau ngày kết thúc.');
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return 'BAN';
  try {
    var w0 = tqnWeatherFetch_(cfg, { tu: tu, den: den });
    if (!w0.ok) { tqnWeatherRaw_(cfg, w0, 'FAILED'); return 'FAILED: ' + w0.err; }
    var n = w0.j.daily.time.length, kq = { THEM: 0, CAP_NHAT: 0, BO_QUA_NHAP_TAY: 0 }, banTin = 0;
    for (var k = 0; k < n; k++) {
      var w = { ok: true, fetchedAt: w0.fetchedAt, data: tqnWeatherDay_(w0.j, k), quaKhu: w0.j.daily.time[k] < homNay };
      if (w.data.tmax === null) continue;
      w.banTin = tqnWeatherBanTin_(cfg, w); if (w.banTin) banTin++;
      var act = tqnWeatherWrite_(cfg, w); kq[act] = (kq[act] || 0) + 1;
    }
    try { tqnDinhDangThoiTiet(); } catch (e) { }
    tqnWeatherRaw_(cfg, w0, 'KHOANG ' + tu + '→' + den);
    tqnAudit_('THOI_TIET', 'Khoảng ' + tu + '→' + den, kq);
    return 'Đã thêm ' + kq.THEM + ' ngày, cập nhật ' + kq.CAP_NHAT + ' ngày' + (kq.BO_QUA_NHAP_TAY ? ', giữ nguyên ' + kq.BO_QUA_NHAP_TAY + ' ngày nhập tay' : '') + ' · ' + banTin + ' bản tin hình trong Drive';
  } finally { lock.releaseLock(); }
}
function tqnMenuWeatherKhoang() {
  var ui = SpreadsheetApp.getUi(), tz = tqnWeatherCfg_().timezone, now = new Date();
  var macDinh = '01/' + Utilities.formatDate(now, tz, 'MM/yyyy') + ' - ' + Utilities.formatDate(now, tz, 'dd/MM/yyyy');
  var r = ui.prompt('Cập nhật thời tiết theo khoảng ngày', 'Nhập: từ ngày - đến ngày (dd/mm/yyyy - dd/mm/yyyy)\nVí dụ: ' + macDinh + '\nNgày đã qua dùng số liệu mô hình Open-Meteo, ghi rõ trong Ghi chú. Dòng nhập tay không bị ghi đè.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = (r.getResponseText() || macDinh).split(/\s*[-–]\s*|\s+đến\s+/);
  try { ui.alert('Thời tiết', tqnWeatherKhoang(t[0], t[1] || t[0]), ui.ButtonSet.OK); } catch (e) { ui.alert('Lỗi: ' + e.message); }
}

/** Hình nhỏ đặt cạnh ô Ghi chú: biểu đồ Sáng/Trưa/Chiều/Tối (khả năng mưa % và nhiệt độ cao °C). Trả data URL PNG, lỗi trả ''. */
function tqnWeatherAnh_(w) {
  try {
    if (typeof Charts === 'undefined') return '';
    var b = w.data.buoi || []; if (!b.length) return '';
    var dt = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Buổi').addColumn(Charts.ColumnType.NUMBER, 'Mưa %').addColumn(Charts.ColumnType.NUMBER, 'Nhiệt °C');
    b.forEach(function (x) { dt.addRow([x.ten + ' ' + (TQN_WMO_VI[x.code] || '').split(' ')[0].toLowerCase(), x.pop === null ? 0 : x.pop, Math.round(x.tmax)]); });
    var p = w.data.date.split('-');
    var c = Charts.newColumnChart().setDataTable(dt.build()).setColors(['#3B82F6', '#F59E0B'])
      .setTitle(p[2] + '/' + p[1] + ' · ' + (TQN_WMO_VI[w.data.code] || '') + ' · ' + Math.round(w.data.tmin) + '–' + Math.round(w.data.tmax) + '°C')
      .setRange(0, 100).setLegendPosition(Charts.Position.BOTTOM).setDimensions(320, 160).build();
    return 'data:image/png;base64,' + Utilities.base64Encode(c.getAs('image/png').getBytes());
  } catch (e) { return ''; }
}

/* ===== Biểu tượng thời tiết từng buổi (đặt ngay cạnh Ghi chú: Sáng | Trưa | Chiều | Tối) ===== */
var TQN_WX_ICON = {"nang": "data:image/png;base64,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", "trang": "data:image/png;base64,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", "itmay": "data:image/png;base64,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", "itmay_dem": "data:image/png;base64,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", "may": "data:image/png;base64,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", "suong": "data:image/png;base64,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", "mua_phun": "data:image/png;base64,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", "mua": "data:image/png;base64,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", "dong": "data:image/png;base64,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"};
function tqnWxIconKey_(c, dem) {
  if (c >= 95) return 'dong'; if (c >= 61 && c <= 67 || c >= 80) return 'mua'; if (c >= 51) return 'mua_phun';
  if (c >= 45) return 'suong'; if (c === 3) return 'may'; if (c >= 1) return dem ? 'itmay_dem' : 'itmay'; return dem ? 'trang' : 'nang';
}
function tqnWeatherIconCells_(sh, row, c0, d) {
  var b = d.buoi || []; if (!b.length || typeof SpreadsheetApp.newCellImage !== 'function') return;
  b.forEach(function (x, i) {
    var c = c0 + i, tieuDe = x.ten + ' ' + (x.tu < 10 ? '0' : '') + x.tu + '–' + x.den + 'h';
    var h = sh.getRange(1, c), cu = String(h.getValue());
    if (!cu || cu === 'Hình thời tiết (tự động)' || /–\d+h$/.test(cu)) h.setValue(tieuDe).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    var mo = x.ten + ': ' + (TQN_WMO_VI[x.code] || '') + ', ' + Math.round(x.tmin) + '–' + Math.round(x.tmax) + '°C, mưa ' + (x.pop === null ? '–' : x.pop + '%');
    sh.getRange(row, c).setValue(SpreadsheetApp.newCellImage().setSourceUrl(TQN_WX_ICON[tqnWxIconKey_(x.code, x.ten === 'Tối')]).setAltTextTitle(mo).setAltTextDescription(mo).build());
    if (sh.getColumnWidth(c) !== 78) sh.setColumnWidth(c, 78);
  });
}

/** Giao diện tab Theo dõi thời tiết cho dễ nhìn (chỉ định dạng, không đổi dữ liệu). */
function tqnDinhDangThoiTiet() {
  var cfg = tqnWeatherCfg_(), t = tqnWeatherSheet_(cfg), sh = t.sh, m = t.m;
  var last = Math.max(2, sh.getLastRow()), nc = m.note + 4;
  var hdr = sh.getRange(1, 1, 1, nc);
  hdr.setBackground('#1B3F7A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 40);
  [['#F59E0B', '#1F2937'], ['#EA580C', '#FFFFFF'], ['#2563EB', '#FFFFFF'], ['#4338CA', '#FFFFFF']].forEach(function (c, i) {
    sh.getRange(1, m.note + 1 + i).setBackground(c[0]).setFontColor(c[1]);
  });
  sh.getRange(1, m.note).setBackground('#0F6E66');
  var body = sh.getRange(2, 1, last - 1, nc);
  body.setFontSize(10).setVerticalAlignment('middle').setFontColor('#1E293B')
    .setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  var nen = []; for (var r = 2; r <= last; r++) { var x = []; for (var c = 0; c < nc; c++) x.push(r % 2 ? '#F1F5F9' : '#FFFFFF'); nen.push(x); }
  body.setBackgrounds(nen);   // 1 lần gọi cho cả bảng
  sh.getRange(2, 1, last - 1, 2).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 3, last - 1, m.note - 3).setHorizontalAlignment('center');
  if (m.date) sh.getRange(2, m.date, last - 1, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, m.note, last - 1, 1).setHorizontalAlignment('left').setFontSize(9);
  sh.getRange(2, m.note + 1, last - 1, 4).setHorizontalAlignment('center').setBackground('#EEF4FB');
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  tqnAudit_('THOI_TIET', 'Định dạng giao diện tab thời tiết', { dong: last - 1 });
  return 'Đã định dạng ' + (last - 1) + ' dòng';
}

/** Thử 1 ngày để xem bản tin (không đụng các ngày khác). */
function tqnThuBanTin1Ngay() { Logger.log(tqnWeatherKhoang('18/11/2026', '18/11/2026')); }
