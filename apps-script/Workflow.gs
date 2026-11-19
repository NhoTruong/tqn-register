/**
 * Workflow.gs – QUẢN LÝ QUY TRÌNH hồ sơ (S2).
 *  1) Bảng quy trình: mỗi hồ sơ (bản Rev mới nhất) đang ở BƯỚC nào và AI ĐANG GIỮ việc (ball-in-court).
 *  2) Cấp mã hồ sơ mới theo mẫu của từng loại (profile COASTAL_2026_V1, sửa được ở TQN_CONFIG),
 *     có khóa tài liệu + sổ cấp mã TQN_CAP_MA: 2 người bấm cùng lúc không trùng số, số đã cấp không cấp lại.
 *  3) Trình lại: giữ số thứ tự, đổi ngày, tăng Rev.
 *  4) Ghi mã vào sổ chi tiết: CHỈ ghi vào ô trống của dòng trống (không đè dữ liệu), có xác nhận.
 */
var TQN_WF_HEADERS = ['Bước', 'Ai đang giữ việc', 'Loại', 'Số hiệu', 'Rev', 'Nội dung', 'Người thực hiện', 'Trình đến',
  'Ngày trình', 'Hạn phản hồi', 'Số ngày chờ', 'Quá hạn', 'Ý kiến gần nhất', 'Vị trí', 'Mã WF'];
var TQN_WF_STEPS = ['', '1. Chưa trình', '2. Đã trình – chờ phản hồi', '3. Chờ CĐT / TVTK', '4. Trả về – nhà thầu trình lại', '5. Đã có kết quả'];
var TQN_CAPMA_HEADERS = ['Thời điểm', 'Người cấp', 'Loại', 'Số hiệu đã cấp', 'Số thứ tự', 'Rev', 'Nội dung', 'Từ số hiệu (trình lại)', 'Ghi vào sổ', 'Trạng thái'];

/** Mẫu mã theo đúng cách các sổ đang ghi (đọc từ dữ liệu Register 19/11/2026). Sửa ở TQN_CONFIG (code_<LOẠI>). */
var TQN_CODE_DEFAULTS = [
  ['code_RFI', '{yymmdd}-COA-TBC-RFI-{nnn} rev{rr}', 'Mẫu mã RFI. {yymmdd} ngày, {nnn} số thứ tự, {rr} lần trình'],
  ['code_RFA', '{yymmdd}-COA-TBC-RFA-{nhom}-{nnn}-rev{rr}', 'Mẫu mã RFA ({nhom} vd COS)'],
  ['code_RFC', '{yymmdd}-COA-TBC-RFC-{nnn}', 'Mẫu mã RFC'],
  ['code_SDS', '{yymmdd}-COA.{goi}-NT-SDS-{nnn}', 'Mẫu mã bản vẽ triển khai ({goi} = NP / HTKT)'],
  ['code_MSS', '{yymmdd}-COA.{goi}-NT-MSS-{nnn}', 'Mẫu mã biện pháp thi công'],
  ['code_MAS', '{yymmdd}-COA.{goi}-NT-MAS-{nnn}', 'Mẫu mã vật tư'],
  ['code_OTS', '{yymmdd}-COA.{goi}-NT-OTS-{nnn}', 'Mẫu mã trình duyệt khác'],
  ['code_rev_start', '1', 'Rev của lần trình đầu khi mẫu có {rr} (sổ RFI/RFA đang bắt đầu từ rev01)'],
  ['code_digits', '3', 'Số chữ số của {nnn}']
];

function tqnEnsureCodeConfig_() {
  tqnEnsureConfig_();
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_CODE_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}
function tqnCodeCfg_() {
  var cfg = tqnConfig_(), d = {};
  TQN_CODE_DEFAULTS.forEach(function (r) { d[r[0]] = cfg[r[0]] !== undefined && cfg[r[0]] !== '' ? cfg[r[0]] : r[1]; });
  cfg.codePattern = function (type) { return d['code_' + String(type).toUpperCase()] || ''; };
  cfg.revStart = parseInt(d.code_rev_start, 10); if (isNaN(cfg.revStart)) cfg.revStart = 1;
  cfg.digits = parseInt(d.code_digits, 10) || 3;
  return cfg;
}

/* ---------------- 1. Bảng quy trình (ball-in-court) ---------------- */
function tqnStage_(r, cfg, today) {
  var C = TQN_COL, n = function (i) { return tqnNoAccent_(String(r[i] || '')); };
  var res = n(C['Kết quả']), stL = n(C['Tình trạng (Sổ trình duyệt)']), stD = n(C['Tình trạng (sổ chi tiết)']);
  var owner = r[C['Người thực hiện']] || 'Nhà thầu', to = r[C['Trình đến']] || '';
  var sub = r[C['Ngày trình']], tv = r[C['TVGS ngày']], cd = r[C['CĐT ngày']], ret = r[C['Ngày giao nhận trả về']], due = r[C['Hạn phản hồi']];
  var step, bic, from = null;
  if (/khong duyet|trinh lai|tra ve/.test(res + ' ' + stL + ' ' + stD)) { step = 4; bic = 'Nhà thầu (' + owner + ')'; from = tqnIsDate_(ret) ? ret : (tqnIsDate_(cd) ? cd : (tqnIsDate_(tv) ? tv : null)); }
  else if (res || /da duyet|duyet co y kien|da phan hoi/.test(stL) || cfg.statusClosed.some(function (s) { return s && stD.indexOf(tqnNoAccent_(s)) >= 0; }) || tqnIsDate_(cd)) { step = 5; bic = '—'; }
  else if (tqnIsDate_(tv)) { step = 3; bic = to && !/tvgs/i.test(tqnNoAccent_(to)) ? to : 'CĐT'; from = tv; }
  else if (tqnIsDate_(sub) || r[C['Mã WF (Sổ trình duyệt)']] || cfg.statusSubmitted.some(function (s) { return s && stD.indexOf(tqnNoAccent_(s)) >= 0; })) { step = 2; bic = to || 'TVGS'; from = tqnIsDate_(sub) ? sub : null; }
  else { step = 1; bic = 'Nhà thầu (' + owner + ')'; }
  var t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var days = from ? Math.max(0, Math.round((t0 - new Date(from.getFullYear(), from.getMonth(), from.getDate())) / 86400000)) : '';
  var late = (step === 2 || step === 3) && tqnIsDate_(due) && due < t0 ? 'QUÁ HẠN' : '';
  var note = r[C['CĐT ý kiến']] || r[C['TVTK ý kiến']] || r[C['TVGS ý kiến']] || '';
  return { step: step, bic: bic, days: days, late: late, note: String(note).slice(0, 200) };
}

/** Lập tab TQN_QUY_TRINH từ kết quả nhập (chỉ bản Rev mới nhất của mỗi hồ sơ). */
function tqnBuildWorkflow_(cfg, imp, today) {
  today = today || new Date();
  var C = TQN_COL, latest = {};
  imp.rows.forEach(function (r) {
    if (r[C['Còn trong nguồn']] === 'Không') return;
    var k = r[C['Khóa so khớp']] ? r[C['Loại']] + '|' + r[C['Khóa so khớp']] : r[0];
    var rv = parseInt(r[C['Rev']], 10); rv = isNaN(rv) ? -1 : rv;
    if (!latest[k] || rv > latest[k].rv) latest[k] = { r: r, rv: rv };
  });
  var out = Object.keys(latest).map(function (k) {
    var r = latest[k].r, s = tqnStage_(r, cfg, today);
    return [TQN_WF_STEPS[s.step], s.bic, r[C['Loại']], r[C['Số hiệu (nguyên văn)']], r[C['Rev']], r[C['Nội dung']], r[C['Người thực hiện']],
      r[C['Trình đến']], r[C['Ngày trình']], r[C['Hạn phản hồi']], s.days, s.late, s.note, r[C['Nguồn']], r[C['Mã WF (Sổ trình duyệt)']]];
  });
  out.sort(function (a, b) { return (a[11] ? 0 : 1) - (b[11] ? 0 : 1) || String(a[0]).localeCompare(String(b[0])) || (+b[10] || 0) - (+a[10] || 0); });
  var sh = tqnOwnSheet_(TQN.PREFIX + 'QUY_TRINH', TQN_WF_HEADERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, TQN_WF_HEADERS.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, TQN_WF_HEADERS.length).setValues(out);
  var count = {}; out.forEach(function (x) { count[x[0]] = (count[x[0]] || 0) + 1; });
  return { rows: out, count: count, late: out.filter(function (x) { return x[11]; }).length };
}

/* ---------------- 2. Cấp mã ---------------- */
function tqnPatternRegex_(pattern, vals, digits) {
  var esc = function (s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); };
  var parts = String(pattern).split(/(\{[a-z]+\})/i), re = '';
  parts.forEach(function (p) {
    if (!p) return;
    var m = /^\{([a-z]+)\}$/i.exec(p);
    if (!m) { re += esc(p).replace(/ /g, '\\s*').replace(/-/g, '\\s*-\\s*'); return; }
    var k = m[1].toLowerCase();
    if (k === 'yymmdd') re += '\\d{6}';
    else if (k === 'nnn') re += '(\\d{1,' + Math.max(digits, 4) + '})';
    else if (k === 'rr') re += '\\d{1,3}';
    else re += esc(vals[k] || '');
  });
  return new RegExp('^\\s*' + re, 'i');
}
function tqnFill_(pattern, vals, cfg) {
  return String(pattern).replace(/\{([a-z]+)\}/gi, function (_, k) {
    k = k.toLowerCase();
    if (k === 'nnn') return ('000000' + vals.nnn).slice(-cfg.digits);
    if (k === 'rr') return ('0' + vals.rr).slice(-2);
    return vals[k] !== undefined ? vals[k] : '';
  });
}
/** Mọi số hiệu đang có (sổ chi tiết + tên file/ghi chú Sổ trình duyệt + sổ cấp mã). */
function tqnAllCodes_(cfg) {
  var src = tqnReadAll_(cfg), codes = [];
  src.details.forEach(function (d) { codes.push(d.code); });
  src.log.forEach(function (L) { L.tbc.forEach(function (t) { codes.push(t.code); }); });
  var sh = tqnSs_().getSheetByName(TQN.PREFIX + 'CAP_MA');
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) codes.push(String(r[0])); });
  return codes;
}
function tqnNextNumber_(cfg, type, vals, codes) {
  var re = tqnPatternRegex_(cfg.codePattern(type), vals, cfg.digits), max = 0;
  codes.forEach(function (c) { var m = re.exec(tqnCleanCode_(c)); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return max + 1;
}
function tqnYymmdd_(d, tz) { return Utilities.formatDate(d, tz, 'yyyyMMdd').slice(2); }

/** Xem trước mã sẽ cấp (không giữ số). */
function tqnPreviewCode(form) {
  var cfg = tqnCodeCfg_(), type = String(form.type || '').toUpperCase(), pat = cfg.codePattern(type);
  if (!pat) throw new Error('Chưa có mẫu mã cho loại ' + type + ' (thêm dòng code_' + type + ' trong TQN_CONFIG).');
  var vals = { nhom: String(form.nhom || 'COS').toUpperCase(), goi: String(form.goi || 'NP').toUpperCase() };
  vals.nnn = tqnNextNumber_(cfg, type, vals, tqnAllCodes_(cfg));
  vals.rr = cfg.revStart; vals.yymmdd = tqnYymmdd_(form.date ? new Date(form.date) : new Date(), cfg.timezone);
  return tqnFill_(pat, vals, cfg);
}

/** Cấp mã thật: khóa tài liệu → đếm lại → ghi sổ cấp mã → nhả khóa. form: {type, nhom, goi, subject, date, write}. */
function tqnIssueCode(form) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) throw new Error('Đang có người cấp mã, thử lại sau vài giây.');
  try {
    tqnEnsureCodeConfig_();
    var cfg = tqnCodeCfg_(), type = String(form.type || '').toUpperCase(), pat = cfg.codePattern(type);
    if (!pat) throw new Error('Chưa có mẫu mã cho loại ' + type);
    if (!String(form.subject || '').trim()) throw new Error('Nhập nội dung hồ sơ trước khi cấp mã.');
    var vals = { nhom: String(form.nhom || 'COS').toUpperCase(), goi: String(form.goi || 'NP').toUpperCase() };
    vals.nnn = tqnNextNumber_(cfg, type, vals, tqnAllCodes_(cfg));
    vals.rr = cfg.revStart; vals.yymmdd = tqnYymmdd_(form.date ? new Date(form.date) : new Date(), cfg.timezone);
    var code = tqnFill_(pat, vals, cfg), where = '';
    if (form.write) where = tqnWriteToDetail_(cfg, type, code, form.subject, form.owner, ('0' + vals.rr).slice(-2));
    tqnCapMaLog_([new Date(), tqnUser_(), type, code, vals.nnn, ('0' + vals.rr).slice(-2), String(form.subject).trim(), '', where, 'ĐÃ CẤP']);
    tqnAudit_('CAP_MA', code, { loai: type, so: vals.nnn, ghi: where });
    if (SpreadsheetApp.flush) SpreadsheetApp.flush();
    return { code: code, where: where };
  } finally { lock.releaseLock(); }
}

/** Trình lại: giữ số thứ tự + nhóm/gói, đổi ngày, tăng Rev. */
function tqnReviseCode(form) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) throw new Error('Đang có người cấp mã, thử lại sau vài giây.');
  try {
    tqnEnsureCodeConfig_();
    var cfg = tqnCodeCfg_(), old = tqnCleanCode_(form.from || ''), type = '';
    var m = /-(RFI|RFA|RFC|SDS|MSS|MAS|OTS)-/i.exec(old); if (m) type = m[1].toUpperCase();
    if (!type) throw new Error('Không nhận ra loại hồ sơ trong số hiệu: ' + old);
    var pat = cfg.codePattern(type), nn = /-(\d{3,4})(?:[- ]?rev\d+)?$/i.exec(old);
    if (!nn) throw new Error('Không thấy số thứ tự trong: ' + old);
    var nhom = (/-RFA-([A-Z0-9]+)-\d/i.exec(old) || [])[1] || 'COS', goi = (/COA\.([A-Z]+)-/i.exec(old) || [])[1] || 'NP';
    var codes = tqnAllCodes_(cfg), keyOld = tqnDocKey_(old), maxRev = -1;
    codes.forEach(function (c) { if (tqnDocKey_(c) === keyOld) { var r = parseInt(tqnRevFromNo_(c) || '-1', 10); if (r > maxRev) maxRev = r; } });
    var curRev = Math.max(maxRev, parseInt(tqnRevFromNo_(old) || String(cfg.revStart - 1), 10));
    if (/\{rr\}/.test(pat) === false) curRev = Math.max(curRev, parseInt(form.rev || '0', 10));
    var vals = { nhom: nhom.toUpperCase(), goi: goi.toUpperCase(), nnn: parseInt(nn[1], 10), rr: curRev + 1, yymmdd: tqnYymmdd_(new Date(), cfg.timezone) };
    var code = tqnFill_(pat, vals, cfg), rev = ('0' + vals.rr).slice(-2), where = '';
    if (form.write) where = tqnWriteToDetail_(cfg, type, code, form.subject || '', form.owner, rev);
    tqnCapMaLog_([new Date(), tqnUser_(), type, code, vals.nnn, rev, String(form.subject || ''), old, where, 'TRÌNH LẠI']);
    tqnAudit_('TRINH_LAI', old + ' → ' + code, { rev: rev, ghi: where });
    if (SpreadsheetApp.flush) SpreadsheetApp.flush();
    return { code: code, rev: rev, where: where };
  } finally { lock.releaseLock(); }
}

function tqnCapMaLog_(row) { tqnOwnSheet_(TQN.PREFIX + 'CAP_MA', TQN_CAPMA_HEADERS).appendRow(row); }
function tqnUser_() { try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; } }

/* ---------------- 4. Ghi vào sổ chi tiết (chỉ ô trống) ---------------- */
function tqnWriteToDetail_(cfg, type, code, subject, owner, rev) {
  var tab = null;
  Object.keys(cfg.alias).forEach(function (k) { if (!tab && cfg.alias[k] === type && cfg.detailSheets.indexOf(k) >= 0) tab = k; });
  var sh = tab && tqnSs_().getSheetByName(tab);
  if (!sh) throw new Error('Không thấy sổ chi tiết cho loại ' + type);
  var lastC = Math.min(sh.getLastColumn(), 40), hdrN = cfg.hdrTo - cfg.hdrFrom + 1;
  var m = tqnDetailMap_(sh.getRange(cfg.hdrFrom, 1, hdrN, lastC).getValues());
  if (m.code < 0 || m.subject < 0) throw new Error('Không dò được cột mã / nội dung của sổ ' + tab);
  var lastR = Math.max(sh.getLastRow(), cfg.firstRow), vals = sh.getRange(cfg.firstRow, 1, lastR - cfg.firstRow + 1, lastC).getValues();
  var lastData = -1;
  vals.forEach(function (r, i) { if (/\d{6}/.test(tqnCleanCode_(r[m.code]))) lastData = i; });
  var target = -1;
  for (var i = lastData + 1; i < vals.length; i++) {
    var r = vals[i];
    if (!String(r[m.code] || '').trim() && !String(r[m.subject] || '').trim()) { target = i; break; }
  }
  if (target < 0) target = vals.length;   // thêm sau dòng cuối
  var row = cfg.firstRow + target;
  var cells = [[m.code, code], [m.subject, subject]];
  if (m.owner >= 0 && owner) cells.push([m.owner, owner]);
  if (m.rev >= 0 && rev) cells.push([m.rev, rev]);
  cells.forEach(function (c) {
    var cur = sh.getRange(row, c[0] + 1).getValue();
    if (String(cur === null || cur === undefined ? '' : cur).trim()) throw new Error('Ô ' + tab + '!' + tqnColLetter_(c[0] + 1) + row + ' đã có dữ liệu – không ghi đè.');
  });
  cells.forEach(function (c) { sh.getRange(row, c[0] + 1).setValue(c[1]); });
  return tab + '!' + row;
}

/* ---------------- Giao diện ---------------- */
function tqnWorkflowData() {
  var r = tqnRunAll_(), wf = r.wf, cols = {};
  TQN_WF_STEPS.slice(1).forEach(function (s) { cols[s] = []; });
  var now = new Date(), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  wf.rows.forEach(function (x) {
    var d = tqnIsDate_(x[9]) ? Math.round((new Date(x[9].getFullYear(), x[9].getMonth(), x[9].getDate()) - t0) / 86400000) : '';
    var open = !/^5/.test(x[0]);
    // trạng thái hạn: QUA_HAN | SAP_HAN (≤2 ngày) | CON_HAN | KHONG_HAN | XONG
    var sla = !open ? 'XONG' : (x[11] ? 'QUA_HAN' : (d === '' ? 'KHONG_HAN' : (d <= 2 ? 'SAP_HAN' : 'CON_HAN')));
    cols[x[0]].push({ bic: x[1], loai: x[2], so: x[3], rev: x[4], nd: x[5], han: tqnFmtDate_(x[9], r.cfg.timezone), d: d, sla: sla,
      cho: x[10], late: x[11], y: x[12], at: x[13], nguoi: x[6] });
  });
  return { steps: TQN_WF_STEPS.slice(1), cols: cols, late: wf.late, total: wf.rows.length, at: Utilities.formatDate(new Date(), r.cfg.timezone, 'dd/MM/yyyy HH:mm') };
}
function tqnMenuWorkflow() {
  var html = tqnHtml_('Workflow').setWidth(1100).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quy trình hồ sơ – ai đang giữ việc');
}
function tqnMenuNewDoc() {
  tqnEnsureCodeConfig_();
  var html = tqnHtml_('NewDoc').setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tạo hồ sơ mới / Trình lại – cấp mã');
}

/* ---------------- Tô màu tình trạng (chỉ định dạng, không đổi dữ liệu) ---------------- */
var TQN_STATUS_COLORS = [   // cùng bảng màu ô chọn (TQN)
  ['Quá hạn', '#FCE4E4', '#B42318'], ['Không duyệt - trình lại', '#FDE9DD', '#A8430A'], ['Đang chờ', '#FEF3D7', '#8A5A00'],
  ['Chưa trình', '#F1F3F6', '#5F6B7A'], ['Đã phản hồi', '#E5EEFB', '#1B3F7A'], ['Duyệt có ý kiến', '#EAF5E3', '#3F6E12'], ['Đã duyệt', '#E1F4EF', '#0F6E56']
];
/** Thêm quy tắc tô màu nhạt cho cột Tình trạng của Sổ trình duyệt, đặt LÊN TRƯỚC quy tắc cũ (quy tắc cũ vẫn giữ, không xóa).
 *  Chỉ đổi màu hiển thị, không đổi dữ liệu. Bấm lại không thêm trùng. */
function tqnApplyStatusColors() {
  var cfg = tqnConfig_(), sh = tqnSs_().getSheetByName(cfg.log_sheet);
  if (!sh) throw new Error('Không thấy tab ' + cfg.log_sheet);
  var all = sh.getRange(1, 1, Math.min(sh.getLastRow(), 60), Math.min(sh.getLastColumn(), 30)).getValues(), h = -1, col = -1;
  for (var r = 0; r < all.length && h < 0; r++) {
    var t = all[r].map(function (x) { return tqnNoAccent_(tqnClean_(x)); });
    if (t.indexOf('loai ho so') >= 0 && t.indexOf('ma ho so') >= 0) { h = r + 1; col = t.indexOf('tinh trang') + 1; }
  }
  if (h < 0 || col < 1) throw new Error('Không dò được cột Tình trạng');
  var rng = sh.getRange(h + 1, col, Math.max(sh.getMaxRows() - h, 1), 1), rules = sh.getConditionalFormatRules(), mine = {};
  var key = function (v, bg) { return String(v) + '|' + String(bg || '').toUpperCase(); };
  rules.forEach(function (ru) {
    var c = ru.getBooleanCondition && ru.getBooleanCondition(); if (!c) return;
    var bg = c.getBackground ? c.getBackground() : '';
    (c.getCriteriaValues() || []).forEach(function (v) { mine[key(v, bg)] = 1; });
  });
  var add = TQN_STATUS_COLORS.filter(function (x) { return !mine[key(x[0], x[1])]; }).map(function (x) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([rng]).build();
  });
  if (add.length) sh.setConditionalFormatRules(add.concat(rules));
  tqnAudit_('TO_MAU', cfg.log_sheet + ' cột ' + tqnColLetter_(col), { them: add.length });
  return add.length;
}
function tqnMenuStatusColors() { SpreadsheetApp.getUi().alert('Đã thêm ' + tqnApplyStatusColors() + ' quy tắc tô màu cho cột Tình trạng (chỉ đổi màu hiển thị, không đổi dữ liệu).'); }
