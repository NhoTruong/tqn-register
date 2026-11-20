/* === TQN PATCH PROD: Quản lý hồ sơ nhanh (issue #1). Đã đo DEV: 18,5s → 6,2s median, khớp JSON 100%. KHÔNG cache. ===
 * Áp dụng: (1) sao lưu PROD; (2) bật dịch vụ nâng cao Sheets v4 (không có vẫn chạy nhờ dự phòng B1 ~2,6s/chặng);
 * (3) dán toàn bộ file này vào CUỐI Mã.gs; (4) trong chuỗi HTML Manager thay call('tqnManagerData' → call('tqnManagerDataNhanh'; (5) lưu, mở nút đo lại. */
function tqnQlKhoiMang_(v, cfg, p) {
  var hdr = v.slice(cfg.hdrFrom - 1, cfg.hdrTo); if (!hdr.length) return null;
  var hr = tqnQlDongTieuDe_(hdr, cfg), row = hdr[hr - cfg.hdrFrom] || [], start = -1;
  for (var c = 0; c < row.length; c++) if (String(row[c]).indexOf(TQN_QL.NHAN) === 0) { start = c + 1; break; }
  if (start < 0) return null;
  var ten = tqnQlTenCot_(p), subRow = v[hr] || [], sub = [];
  for (var j = 0; j < ten.length; j++) sub.push(subRow[start - 1 + j] === undefined ? '' : subRow[start - 1 + j]);
  var at = function (name) { for (var i = 0; i < sub.length; i++) if (tqnQlBo_(sub[i]) === tqnQlBo_(name)) return start + i; return -1; };
  var k = { start: start, giu: at('Ai đang giữ'), hdrRow: hr };
  k.map = tqnDetailMap_(hdr);
  return k;
}
function tqnQlTheoMaTuMang_(lay) {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), out = {}, ten = [];
  cfg.detailSheets.forEach(function (n) { if (tqnQlLoaiTab_(cfg, prof, n)) ten.push(n); });
  var all = lay(ten);
  ten.forEach(function (n) {
    var type = tqnQlLoaiTab_(cfg, prof, n), v = all[n];
    if (!v || v.length < cfg.firstRow) return;
    var k = tqnQlKhoiMang_(v, cfg, prof[type]); if (!k || k.giu < 0 || k.map.code == null || k.map.code < 0) return;
    for (var i = cfg.firstRow - 1; i < v.length; i++) {
      var r = v[i] || [], ma = r[k.map.code], gi = r[k.giu - 1];
      if (ma && gi) out[tqnKeyOf_(ma)] = { giu: String(gi), tab: n, row: i + 1 };
    }
  });
  return out;
}
/** B1: mỗi tab 1 lần getDataRange().getValues() */
var TQN_NHANH = { dl: null, lyDo: '' };
function tqnDocGop_() {                     // đọc gộp: tab quản lý + 5 tab văn bản, 1 lệnh; lỗi → dự phòng B1 (ghi lý do)
  if (TQN_NHANH.dl) return TQN_NHANH.dl;
  var cfg = tqnConfig_(), prof = tqnQlMau_(), ten = [];
  cfg.detailSheets.forEach(function (n) { if (tqnQlLoaiTab_(cfg, prof, n)) ten.push(n); });
  TQN_DOC_TABS.forEach(function (d) { if (ten.indexOf(d.tab) < 0) ten.push(d.tab); });
  var ss = tqnSs_(), o = {};
  try {
    var co = ten.filter(function (n) { return !!ss.getSheetByName(n); });
    var r = Sheets.Spreadsheets.Values.batchGet(ss.getId(), { ranges: co.map(function (n) { return "'" + n.replace(/'/g, "''") + "'"; }), valueRenderOption: 'FORMATTED_VALUE' });
    (r.valueRanges || []).forEach(function (vr, i) { o[co[i]] = vr.values || []; });
    TQN_NHANH.lyDo = 'batchGet';
  } catch (e) {
    TQN_NHANH.lyDo = 'B1 (batchGet lỗi: ' + String(e.message).slice(0, 120) + ')'; console.warn(TQN_NHANH.lyDo);
    ten.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) o[n] = sh.getDataRange().getValues(); });
  }
  TQN_NHANH.dl = o; return o;
}
function tqnQlTheoMaNhanh_() { return tqnQlTheoMaTuMang_(function () { return tqnDocGop_(); }); }
function tqnReadDocTabMang_(def, all0, today) {
  if (!all0 || all0.length < 4) return [];
  var lastC = 0; all0.forEach(function (r) { lastC = Math.max(lastC, r.length); }); lastC = Math.min(lastC, 30);
  var all = all0.map(function (r) { var x = []; for (var j = 0; j < lastC; j++) x.push(r[j] === undefined ? '' : r[j]); return x; });
  var hdrTo = Math.min(6, all.length);
  var find = function (keys) {
    if (!keys || !keys.length) return -1;
    for (var k = 0; k < keys.length; k++) { var key = tqnNoAccent_(keys[k]); for (var r = 0; r < hdrTo; r++) for (var c = 0; c < lastC; c++) if (tqnNoAccent_(tqnClean_(all[r][c])).indexOf(key) === 0) return c; }
    return -1;
  };
  var c = { code: find(def.code), subj: find(def.subj), date: find(def.date), reply: find(def.reply), note: find(def.note), who: find(def.who), status: find(def.status) };
  if (c.code < 0) return [];
  var first = 0;
  for (var r = 0; r < hdrTo; r++) if (all[r].some(function (x) { return /^stt$/i.test(String(tqnClean_(x))); })) first = r + 1;
  var t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate()), out = [];
  for (var i = first; i < all.length; i++) {
    var row = all[i], code = String(tqnClean_(tqnCell_(row, c.code))), subj = String(tqnClean_(tqnCell_(row, c.subj)));
    if (!code || !subj) continue;
    if (/^(nam \d{4}|[ivx]+\.?)$/i.test(tqnNoAccent_(code))) continue;
    var d = tqnAnyDate_(tqnCell_(row, c.date)), rp = tqnAnyDate_(tqnCell_(row, c.reply));
    var note = String(tqnClean_(tqnCell_(row, c.note))), st = String(tqnClean_(tqnCell_(row, c.status)));
    var done = !!rp || /dong|xong|hoan thanh|da xu ly|da phan hoi/.test(tqnNoAccent_(st));
    out.push({ loai: def.loai, tab: def.tab, row: i + 1, ma: code, nd: subj.replace(/\s+/g, ' '),
      ngay: d ? Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '', phanhoi: rp ? Utilities.formatDate(rp, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '',
      nguoi: String(tqnClean_(tqnCell_(row, c.who))), ykien: note, tt: done ? 'Đã phản hồi' : (st || 'Chờ phản hồi'), cho: (!done && d) ? Math.max(0, Math.round((t0 - d) / 86400000)) : '' });
  }
  return out;
}
function tqnDocsDataNhanh(today) { today = today || new Date(); var dl = tqnDocGop_(), rows = []; TQN_DOC_TABS.forEach(function (def) { try { rows = rows.concat(tqnReadDocTabMang_(def, dl[def.tab], today)); } catch (e) { } }); return rows; }
/** Bản nhanh, không cache */
function tqnManagerDataNhanh(nhanh) {
  var g1 = tqnQlTheoMa_, g2 = tqnDocsData; TQN_NHANH.dl = null;
  tqnQlTheoMa_ = tqnQlTheoMaNhanh_; tqnDocsData = tqnDocsDataNhanh;
  try { return tqnManagerData(nhanh); } finally { tqnQlTheoMa_ = g1; tqnDocsData = g2; TQN_NHANH.dl = null; }
}
/** Bản nhanh + cache theo phiên bản dữ liệu (tqnDv đổi khi sửa sổ / bấm Làm mới) */
