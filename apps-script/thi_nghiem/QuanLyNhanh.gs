/* ===== [DEV] tqnManagerData bản nhanh: 1 lệnh batchGet cho mọi tab + cache theo phiên bản (issue #1) =====
 * Không sửa hàm gốc: tạm thay tqnQlTheoMa_ và tqnDocsData trong lúc gọi tqnManagerData, rồi trả lại. */
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
function tqnManagerDataCache(nhanh, boQuaCache) {
  var c = CacheService.getDocumentCache(), dv = PropertiesService.getDocumentProperties().getProperty('tqnDv') || '0';
  var k = 'mgr:' + tqnConfig_().project_id + ':' + dv + ':' + (nhanh ? 1 : 0);
  if (!boQuaCache) { var hit = c.get(k); if (hit) { var o = JSON.parse(hit); o.cache = 'HIT'; return o; } }
  var d = tqnManagerDataNhanh(nhanh), j = JSON.stringify(d);
  if (j.length < 95000) c.put(k, j, 120);
  d.cache = j.length < 95000 ? 'MISS' : 'QUA_LON'; return d;
}
/** So khớp sâu cũ ↔ mới (bỏ trường thời điểm 'at') */
function tqnSoKhop_(a, b, duong, loi) {
  if (loi.length > 15) return;
  if (a instanceof Date) a = a.toISOString(); if (b instanceof Date) b = b.toISOString();
  if (typeof a !== 'object' || a === null || typeof b !== 'object' || b === null) { if (a !== b) loi.push(duong + ': cũ=' + JSON.stringify(a) + ' mới=' + JSON.stringify(b)); return; }
  var k = {}; Object.keys(a).forEach(function (x) { k[x] = 1; }); Object.keys(b).forEach(function (x) { k[x] = 1; });
  Object.keys(k).forEach(function (x) { if (x === 'at' || x === 'cache') return; tqnSoKhop_(a[x], b[x], duong + '.' + x, loi); });
}
/** CHẠY: so khớp + đo 20 lần bản nhanh (không cache) + 20 lần cache trúng. Chỉ đọc sổ. */
function tqnA_DoQuanLyNhanhDEV() {
  var t0 = Date.now(), cu = tqnManagerData(true), tCu = Date.now() - t0;
  var moi = tqnManagerDataNhanh(true), loi = []; tqnSoKhop_(JSON.parse(JSON.stringify(cu)), JSON.parse(JSON.stringify(moi)), 'kq', loi);
  var n1 = [], n2 = [];
  for (var i = 0; i < 20; i++) { var a = Date.now(); tqnManagerDataNhanh(true); n1.push(Date.now() - a); }
  tqnManagerDataCache(true, true);
  for (var j = 0; j < 20; j++) { var b = Date.now(); var h = tqnManagerDataCache(true, false); n2.push(Date.now() - b); }
  Logger.log(JSON.stringify({ cu_1lan_ms: tCu, khop: loi.length === 0, soChoLech: loi.length, lech: loi, nguonDoc: TQN_NHANH.lyDo || 'batchGet',
    soDong: moi.rows.length, soVanBan: moi.docs.length, kb: Math.round(JSON.stringify(moi).length / 1024),
    nhanh_khongCache: tqnTk_(n1), cacheTrung: tqnTk_(n2), cacheTrangThaiCuoi: h.cache }));
}
