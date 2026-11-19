/* ===== [DEV] Thí nghiệm issue #1: so sánh cách đọc cho tqnQlTheoMa_ (chỉ đọc, không ghi sổ) ===== */
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
function tqnQlTheoMaB1_() {
  return tqnQlTheoMaTuMang_(function (ten) { var o = {}, ss = tqnSs_(); ten.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) o[n] = sh.getDataRange().getValues(); }); return o; });
}
/** B2: 1 lệnh Sheets API batchGet cho mọi tab (cần dịch vụ nâng cao Sheets) */
function tqnQlTheoMaB2_() {
  return tqnQlTheoMaTuMang_(function (ten) {
    var r = Sheets.Spreadsheets.Values.batchGet(tqnSs_().getId(), { ranges: ten.map(function (n) { return "'" + n.replace(/'/g, "''") + "'"; }), valueRenderOption: 'UNFORMATTED_VALUE' });
    var o = {}; (r.valueRanges || []).forEach(function (vr, i) { o[ten[i]] = vr.values || []; }); return o;
  });
}
/** A: cache JSON theo phiên bản dữ liệu */
function tqnQlTheoMaA_(fn) {
  var c = CacheService.getDocumentCache(), dv = PropertiesService.getDocumentProperties().getProperty('tqnDv') || '0', k = 'qlma:' + dv;
  var hit = c.get(k); if (hit) return JSON.parse(hit);
  var o = fn(); var j = JSON.stringify(o); if (j.length < 95000) c.put(k, j, 300); return o;
}
function tqnTk_(a) { var s = a.slice().sort(function (x, y) { return x - y; }); return { n: s.length, min: s[0], median: s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2, p95: s[Math.min(s.length - 1, Math.ceil(0.95 * s.length) - 1)], max: s[s.length - 1] }; }
function tqnSoSanh_(ten, fn, n, goc) {
  var t = [], ok = true, kq;
  for (var i = 0; i < n; i++) { var a = Date.now(); kq = fn(); t.push(Date.now() - a); }
  if (goc) { var A = Object.keys(goc).sort(), B = Object.keys(kq).sort(); ok = A.length === B.length && A.every(function (k, i) { return k === B[i] && goc[k].giu === kq[k].giu && goc[k].row === kq[k].row; }); }
  var r = { cach: ten, ms: tqnTk_(t), soMa: Object.keys(kq).length, kb: Math.round(JSON.stringify(kq).length / 1024), giongGoc: ok };
  Logger.log(JSON.stringify(r)); return r;
}
/** Chạy 1: B0 gốc ×20 */
function tqnDoB0DEV() { tqnSoSanh_('B0_goc_getRange', tqnQlTheoMa_, 20); }
/** Chạy 2: B1 ×20, B2 ×20, A (cache trúng) ×20 – có so khớp kết quả với bản gốc */
function tqnDoB1B2ADEV() {
  var goc = tqnQlTheoMa_();
  tqnSoSanh_('B1_getDataRange_moi_tab', tqnQlTheoMaB1_, 20, goc);
  try { tqnSoSanh_('B2_Sheets_batchGet_1_lenh', tqnQlTheoMaB2_, 20, goc); } catch (e) { Logger.log('B2 lỗi: ' + e.message); }
  CacheService.getDocumentCache().remove('qlma:' + (PropertiesService.getDocumentProperties().getProperty('tqnDv') || '0'));
  var a0 = Date.now(); tqnQlTheoMaA_(tqnQlTheoMaB1_); Logger.log(JSON.stringify({ cach: 'A_cache_truot_lan_dau', ms: Date.now() - a0 }));
  tqnSoSanh_('A_cache_trung', function () { return tqnQlTheoMaA_(tqnQlTheoMaB1_); }, 20, goc);
  var c0 = Date.now(); tqnConfig_(); tqnQlMau_(); Logger.log(JSON.stringify({ cach: 'chi_doc_cau_hinh', ms: Date.now() - c0 }));
}

/* [DEV] Đếm lý do lọc tqnDocsData – xem issue #1 comment 00:05 (code đầy đủ đã chạy trên bản DEV) */
