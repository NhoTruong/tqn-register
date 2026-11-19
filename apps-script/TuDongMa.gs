/**
 * TuDongMa.gs – Tự cấp mã hồ sơ khi gõ Nội dung (SD = bản vẽ shop SDS, RFI…).
 * BPTC chỉ cần gõ "Nội dung chi tiết" ở dòng mới → ô "Mã hồ sơ" tự điền số tiếp theo.
 * Nguyên tắc: chỉ ghi vào Ô TRỐNG (Mã hồ sơ, STT, Rev); không sửa mã đã có; có khóa chống trùng số;
 * mọi mã cấp ghi vào TQN_CAP_MA (Trạng thái = TỰ ĐỘNG) + TQN_AUDIT.
 * Bật/tắt theo loại: khóa auto_code_types trong TQN_CONFIG (mặc định "SDS RFI").
 */
var TQN_AUTO_CODE_DEFAULT = ['auto_code_types', 'SDS RFI', 'Loại hồ sơ tự cấp mã khi gõ Nội dung (SDS = tab SD, RFI…). Cách nhau dấu cách. Để trống = tắt'];

function tqnEnsureAutoCodeCfg_() {
  tqnEnsureCodeConfig_();
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG), have = false;
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { if (String(r[0]).trim() === TQN_AUTO_CODE_DEFAULT[0]) have = true; });
  if (!have) sh.appendRow(TQN_AUTO_CODE_DEFAULT);
}
function tqnAutoTypes_(cfg) {
  var v = cfg.auto_code_types === undefined ? TQN_AUTO_CODE_DEFAULT[1] : cfg.auto_code_types;
  return String(v).toUpperCase().split(/[\s,;]+/).filter(String);
}
/** Loại hồ sơ của tab nếu tab đó được bật tự cấp mã, ngược lại ''. */
function tqnAutoTypeOfTab_(cfg, name) {
  if (cfg.detailSheets.indexOf(name) < 0) return '';
  var t = String(cfg.alias[name] || '').toUpperCase();
  return (t && tqnAutoTypes_(cfg).indexOf(t) >= 0 && cfg.codePattern(t)) ? t : '';
}

/**
 * Tìm (và nếu ghi=true thì điền) mã cho các dòng có Nội dung mà Mã hồ sơ còn trống.
 * r1..r2: khoảng dòng cần xét (số dòng sheet); bỏ trống = cả sổ. Trả [{row, code, subject, stt, rev}].
 */
function tqnCapMaTrong_(sh, cfg, type, r1, r2, ghi) {
  var lastC = Math.min(sh.getLastColumn(), 40), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
  var m = tqnDetailMap_(hdr), cStt = tqnFindCol_(hdr, ['STT']);
  var lastR = sh.getLastRow(); if (lastR < cfg.firstRow) return [];
  var vals = sh.getRange(cfg.firstRow, 1, lastR - cfg.firstRow + 1, lastC).getValues();
  var pat = cfg.codePattern(type), coRev = /\{rr\}/.test(pat);
  var re = /-(RFI|RFA|RFC|SDS|MSS|MAS|OTS)-/i, ma = function (v) { return tqnCleanCode_(v === null || v === undefined ? '' : v); };
  // Mã đã cấp ở sổ cấp mã (gồm cả số đã giữ nhưng chưa ghi sổ) để không cấp trùng
  var codes = vals.map(function (r) { return ma(r[m.code]); }).filter(function (c) { return re.test(c); });
  var shLog = tqnSs_().getSheetByName(TQN.PREFIX + 'CAP_MA');
  if (shLog && shLog.getLastRow() >= 2) shLog.getRange(2, 4, shLog.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) codes.push(ma(r[0])); });
  var from = Math.max((r1 || cfg.firstRow) - cfg.firstRow, 0), to = Math.min((r2 || lastR) - cfg.firstRow, vals.length - 1);
  var out = [], yymmdd = tqnYymmdd_(new Date(), cfg.timezone), daCap = {};
  for (var i = from; i <= to; i++) {
    var r = vals[i], sub = String(r[m.subject] === null || r[m.subject] === undefined ? '' : r[m.subject]).trim();
    if (!sub || ma(r[m.code])) continue;
    // Kế thừa gói (NP/HTKT) và nhóm (RFA) từ mã gần nhất phía trên cùng loại
    var goi = 'NP', nhom = 'COS';
    for (var k = i - 1; k >= 0; k--) {
      var c = ma(vals[k][m.code]); if (!re.test(c)) continue;
      goi = ((/COA\.([A-Z]+)-/i.exec(c) || [])[1] || goi).toUpperCase();
      nhom = ((/-RFA-([A-Z0-9]+)-\d/i.exec(c) || [])[1] || nhom).toUpperCase();
      break;
    }
    var v = { goi: goi, nhom: nhom, yymmdd: yymmdd, rr: cfg.revStart };
    v.nnn = tqnNextNumber_(cfg, type, v, codes);
    var code = tqnFill_(pat, v, cfg);
    codes.push(code); daCap[i] = 1;
    var row = cfg.firstRow + i, it = { row: row, code: code, subject: sub, nnn: v.nnn, stt: '', rev: '' };
    // STT: chỉ điền khi trống, = STT số của dòng trên + 1
    if (cStt >= 0 && !String(r[cStt] === null || r[cStt] === undefined ? '' : r[cStt]).trim()) {
      it.stt = 1;   // đầu mục (I, II…) phía trên mà chưa có số → bắt đầu 1
      for (var j = i - 1; j >= 0; j--) {
        var s = String(vals[j][cStt] === null || vals[j][cStt] === undefined ? '' : vals[j][cStt]).trim();
        if (/^\d+$/.test(s)) { it.stt = parseInt(s, 10) + 1; break; }
        if (/^[IVX]+$/i.test(s)) break;
      }
      vals[i][cStt] = it.stt;
    }
    if (coRev && m.rev >= 0 && !String(r[m.rev] === null || r[m.rev] === undefined ? '' : r[m.rev]).trim()) it.rev = ('0' + v.rr).slice(-2);
    vals[i][m.code] = code;
    out.push(it);
  }
  if (ghi && out.length) {
    var logRows = [];
    out.forEach(function (it) {
      sh.getRange(it.row, m.code + 1).setValue(it.code).setNote('TQN tự cấp lúc ' + tqnGioNgay_(cfg) + ' (theo nội dung vừa gõ).');
      if (it.stt !== '') sh.getRange(it.row, cStt + 1).setValue(it.stt);
      if (it.rev) sh.getRange(it.row, m.rev + 1).setNumberFormat('@').setValue(it.rev);
      logRows.push([new Date(), tqnUser_() || 'tự động', type, it.code, it.nnn, it.rev || (coRev ? ('0' + cfg.revStart).slice(-2) : ''), it.subject, '', sh.getName() + '!' + it.row, 'TỰ ĐỘNG']);
    });
    var moi = !tqnSs_().getSheetByName(TQN.PREFIX + 'CAP_MA'), lg = tqnOwnSheet_(TQN.PREFIX + 'CAP_MA', TQN_CAPMA_HEADERS);
    if (moi) try { lg.hideSheet(); } catch (e) { }   // sổ cấp mã là tab phụ
    lg.getRange(lg.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
    tqnAudit_('CAP_MA_TU_DONG', out.map(function (x) { return x.code; }).join(', '), { tab: sh.getName(), so: out.length });
    try { tqnCacheClear_(); } catch (e) { }
  }
  return out;
}

/** Trình kích hoạt khi sửa ô (cài bằng tqnBatTuDongMa). Chỉ chạy ở tab được bật, dòng dữ liệu. */
function tqnTuDongMaKhiSua(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet(), cfg = tqnCodeCfg_(), type = tqnAutoTypeOfTab_(cfg, sh.getName());
  if (!type) return;
  var r1 = e.range.getRow(), r2 = e.range.getLastRow();
  if (r2 < cfg.firstRow) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return;
  try {
    var out = tqnCapMaTrong_(sh, cfg, type, Math.max(r1, cfg.firstRow), r2, true);
    var tay = tqnKiemOGoTay_(sh, cfg, type, e.range);
    var db = tqnDongBoMa_(sh, cfg, type, Math.max(r1, cfg.firstRow), r2);
    var msg = out.map(function (x) { return 'Dòng ' + x.row + ': ' + x.code; });
    db.forEach(function (x) { var i = -1; out.forEach(function (o, k) { if (o.row === x.row) i = k; }); if (i >= 0) msg[i] = 'Dòng ' + x.row + ': ' + x.moi; else msg.push('Dòng ' + x.row + ': ' + x.cu + ' → ' + x.moi); });
    if (tay.length) { tqnSs_().toast(tay.join('\n'), 'TQN – kiểm mã gõ tay', 8); return; }
    if (msg.length) tqnSs_().toast(msg.join('\n'), 'TQN – mã ' + type + ' tự cập nhật', 6);
  } finally { lock.releaseLock(); }
}

/** Bật tự cấp mã: cài trình kích hoạt "khi sửa" (không trùng). Trả chuỗi báo cáo. */
function tqnBatTuDongMa() {
  tqnEnsureAutoCodeCfg_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnTuDongMaKhiSua') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tqnTuDongMaKhiSua').forSpreadsheet(tqnSs_()).onEdit().create();
  var cfg = tqnCodeCfg_(), tabs = cfg.detailSheets.filter(function (n) { return tqnAutoTypeOfTab_(cfg, n); });
  try { tqnLamDepCotMa(); } catch (e) { }
  tqnAudit_('BAT_TU_CAP_MA', tabs.join(', '), {});
  return tabs;
}
function tqnTatTuDongMa() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnTuDongMaKhiSua') { ScriptApp.deleteTrigger(t); n++; } });
  return n;
}

/** Danh sách dòng có Nội dung nhưng thiếu mã ở mọi tab được bật (không ghi). */
function tqnDongThieuMa_(ghi) {
  var cfg = tqnCodeCfg_(), all = [];
  cfg.detailSheets.forEach(function (n) {
    var type = tqnAutoTypeOfTab_(cfg, n), sh = type && tqnSs_().getSheetByName(n);
    if (!sh) return;
    tqnCapMaTrong_(sh, cfg, type, 0, 0, ghi).forEach(function (x) { x.tab = n; all.push(x); });
  });
  return all;
}

/* ---------- menu ---------- */
function tqnMenuBatTuDongMa() {
  var ui = SpreadsheetApp.getUi(), tabs = tqnBatTuDongMa();
  var msg = '✔ Đã bật tự cấp mã cho tab: ' + tabs.join(', ') + '\nTừ giờ chỉ cần gõ "Nội dung chi tiết" ở dòng mới → ô Mã hồ sơ tự điền số tiếp theo (chỉ điền ô trống, không sửa mã cũ).';
  var thieu = tqnDongThieuMa_(false);
  if (!thieu.length) { ui.alert('Tự cấp mã', msg, ui.ButtonSet.OK); return; }
  var ds = thieu.slice(0, 15).map(function (x) { return x.tab + ' dòng ' + x.row + ' → ' + x.code; }).join('\n');
  var ok = ui.alert('Tự cấp mã', msg + '\n\nĐang có ' + thieu.length + ' dòng có nội dung nhưng chưa có mã:\n' + ds + (thieu.length > 15 ? '\n…' : '') + '\n\nCấp mã cho các dòng này luôn?', ui.ButtonSet.YES_NO);
  if (ok === ui.Button.YES) tqnMenuCapMaTrong_(true);
}
function tqnMenuCapMaTrong() { tqnMenuCapMaTrong_(false); }
function tqnMenuCapMaTrong_(daHoi) {
  var ui = SpreadsheetApp.getUi();
  if (!daHoi) {
    var xem = tqnDongThieuMa_(false);
    if (!xem.length) { ui.alert('Không có dòng nào thiếu mã ở các tab tự cấp mã.'); return; }
    var ds = xem.slice(0, 15).map(function (x) { return x.tab + ' dòng ' + x.row + ' → ' + x.code; }).join('\n');
    if (ui.alert('Cấp mã cho dòng còn trống', ds + (xem.length > 15 ? '\n…' : '') + '\n\nGhi ' + xem.length + ' mã này?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) { ui.alert('Đang có người cấp mã, thử lại sau vài giây.'); return; }
  try { var out = tqnDongThieuMa_(true); ui.alert('Đã cấp ' + out.length + ' mã (ghi ở TQN_CAP_MA, trạng thái TỰ ĐỘNG).'); }
  finally { lock.releaseLock(); }
}
function tqnMenuTatTuDongMa() { SpreadsheetApp.getUi().alert('Đã tắt tự cấp mã (' + tqnTatTuDongMa() + ' trình kích hoạt).'); }

/* ================= Kiểm & sửa mã nhập tay (19/11) =================
 * Mã cũ BPTC gõ tay hay bị: trùng số (vd SDS-025 dùng 2 lần), dư khoảng trắng ("261116- COA"), STT lộn xộn.
 * Kiểm = chỉ đọc. Sửa = có hỏi trước, ghi TQN_CAP_MA (SỬA MÃ TRÙNG / CHUẨN HÓA) + TQN_AUDIT.
 * Mã trùng: dòng ĐẦU giữ mã; dòng sau nếu CHƯA trình (tình trạng trống / đang thực hiện) mới đổi sang số kế tiếp,
 * đã trình rồi thì chỉ báo để anh xử lý tay (không đổi mã đã gửi TVGS).
 */
function tqnMaKhoa_(c) {   // khóa so trùng: bỏ ngày đầu và rev cuối
  return tqnCleanCode_(c).toUpperCase().replace(/^\d{6}-/, '').replace(/[- ]?REV ?\d+$/, '').replace(/\s+/g, '');
}
function tqnKiemMaTab_(sh, cfg, type) {
  var lastC = Math.min(sh.getLastColumn(), 40), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
  var m = tqnDetailMap_(hdr), cStt = tqnFindCol_(hdr, ['STT']), lastR = sh.getLastRow();
  var kq = { tab: sh.getName(), type: type, trung: [], sai: [], stt: [], cStt: cStt, cMa: m.code, cRev: m.rev };
  if (lastR < cfg.firstRow) return kq;
  var vals = sh.getRange(cfg.firstRow, 1, lastR - cfg.firstRow + 1, lastC).getValues();
  var re = /-(RFI|RFA|RFC|SDS|MSS|MAS|OTS)-/i, s_ = function (v) { return String(v === null || v === undefined ? '' : v).trim(); };
  var daGap = {}, codes = [], dem = 0;
  vals.forEach(function (r) { var c = tqnCleanCode_(r[m.code]); if (re.test(c)) codes.push(c); });
  var shLog = tqnSs_().getSheetByName(TQN.PREFIX + 'CAP_MA');
  if (shLog && shLog.getLastRow() >= 2) shLog.getRange(2, 4, shLog.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) codes.push(tqnCleanCode_(r[0])); });
  var yymmdd = tqnYymmdd_(new Date(), cfg.timezone), pat = cfg.codePattern(type);
  vals.forEach(function (r, i) {
    var row = cfg.firstRow + i, goc = s_(r[m.code]), sub = s_(r[m.subject]);
    var laMuc = !sub && goc && !re.test(goc);          // dòng đầu mục (KẾT CẤU CÔNG VIÊN…)
    if (laMuc || /^[IVX]+$/i.test(s_(r[cStt]))) { dem = 0; if (laMuc || !sub) return; }
    if (!sub && !goc) return;
    // STT chạy 1..n trong từng mục
    dem++;
    if (cStt >= 0 && s_(r[cStt]) !== String(dem)) kq.stt.push({ row: row, cu: s_(r[cStt]), moi: dem });
    if (!re.test(goc)) return;
    var sach = tqnCleanCode_(goc);
    if (sach !== goc) kq.sai.push({ row: row, cu: goc, moi: sach });
    var k = tqnMaKhoa_(sach);
    if (daGap[k]) {
      var st = tqnNoAccent_(s_(r[m.status])).toLowerCase(), chuaTrinh = !st || /dang thuc hien|chua trinh/.test(st);
      var it = { row: row, cu: sach, dongDau: daGap[k], subject: sub, doiDuoc: chuaTrinh, moi: '' };
      if (chuaTrinh && pat) {
        var v = { goi: ((/COA\.([A-Z]+)-/i.exec(sach) || [])[1] || 'NP').toUpperCase(), nhom: ((/-RFA-([A-Z0-9]+)-\d/i.exec(sach) || [])[1] || 'COS').toUpperCase(), yymmdd: yymmdd, rr: parseInt(tqnRevFromNo_(sach) || cfg.revStart, 10) };
        v.nnn = tqnNextNumber_(cfg, type, v, codes); it.moi = tqnFill_(pat, v, cfg); it.nnn = v.nnn; codes.push(it.moi);
      }
      kq.trung.push(it);
    } else daGap[k] = row;
  });
  return kq;
}
function tqnKiemMaTatCa_() {
  var cfg = tqnCodeCfg_(), all = [];
  cfg.detailSheets.forEach(function (n) {
    var type = tqnAutoTypeOfTab_(cfg, n), sh = type && tqnSs_().getSheetByName(n);
    if (sh) all.push(tqnKiemMaTab_(sh, cfg, type));
  });
  return all;
}
function tqnKiemMaText_(all) {
  var L = [];
  all.forEach(function (k) {
    L.push('■ ' + k.tab + ': ' + k.trung.length + ' mã trùng · ' + k.sai.length + ' mã dư khoảng trắng · ' + k.stt.length + ' STT lệch');
    k.trung.forEach(function (t) { L.push('  Trùng dòng ' + t.row + ' (đã có ở dòng ' + t.dongDau + '): ' + t.cu + (t.moi ? ' → ' + t.moi : ' → ĐÃ TRÌNH, anh xử lý tay')); });
    k.sai.forEach(function (t) { L.push('  Dòng ' + t.row + ': "' + t.cu + '" → "' + t.moi + '"'); });
    if (k.stt.length) L.push('  STT: ' + k.stt.slice(0, 12).map(function (t) { return 'dòng ' + t.row + ' ' + (t.cu || 'trống') + '→' + t.moi; }).join(', ') + (k.stt.length > 12 ? '…' : ''));
  });
  return L.join('\n');
}
/** Ghi sửa: mã trùng chưa trình, mã dư khoảng trắng, STT. Trả số ô đã sửa. */
function tqnSuaMa_(all) {
  var cfg = tqnCodeCfg_(), n = 0, log = [];
  all.forEach(function (k) {
    var sh = tqnSs_().getSheetByName(k.tab);
    k.sai.forEach(function (t) { sh.getRange(t.row, k.cMa + 1).setValue(t.moi); n++; log.push([new Date(), tqnUser_(), k.type, t.moi, '', '', '', t.cu, k.tab + '!' + t.row, 'CHUẨN HÓA']); });
    k.trung.forEach(function (t) { if (!t.moi) return; sh.getRange(t.row, k.cMa + 1).setValue(t.moi); n++; log.push([new Date(), tqnUser_(), k.type, t.moi, t.nnn, '', t.subject, t.cu, k.tab + '!' + t.row, 'SỬA MÃ TRÙNG']); });
    k.stt.forEach(function (t) { sh.getRange(t.row, k.cStt + 1).setValue(t.moi); n++; });
  });
  if (log.length) { var lg = tqnOwnSheet_(TQN.PREFIX + 'CAP_MA', TQN_CAPMA_HEADERS); lg.getRange(lg.getLastRow() + 1, 1, log.length, log[0].length).setValues(log); }
  tqnAudit_('SUA_MA_NHAP_TAY', tqnKiemMaText_(all).slice(0, 4000), { o: n });
  try { tqnCacheClear_(); } catch (e) { }
  return n;
}
function tqnMenuKiemMa() {
  var ui = SpreadsheetApp.getUi(), all = tqnKiemMaTatCa_(), txt = tqnKiemMaText_(all);
  var co = all.some(function (k) { return k.trung.length || k.sai.length || k.stt.length; });
  if (!co) { ui.alert('Kiểm mã nhập tay', 'Không thấy lỗi.\n' + txt, ui.ButtonSet.OK); return; }
  if (ui.alert('Kiểm mã nhập tay', txt + '\n\nSửa các mục trên? (chỉ đổi mã trùng CHƯA trình; mã đã trình chỉ báo)', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(20000)) { ui.alert('Đang có người cấp mã, thử lại sau.'); return; }
  try { ui.alert('Đã sửa ' + tqnSuaMa_(tqnKiemMaTatCa_()) + ' ô. Mã cũ lưu ở TQN_CAP_MA (cột "Từ số hiệu").'); } finally { lock.releaseLock(); }
}
/** Chạy trong trình soạn: chỉ kiểm, ghi log (không sửa). */
function tqnKiemMaLog() { var s = tqnKiemMaText_(tqnKiemMaTatCa_()); Logger.log(s); return s; }
/** Chạy trong trình soạn: sửa (anh đã đồng ý). */
function tqnSuaMaNgay() {
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(20000)) throw new Error('Đang bận');
  try { var all = tqnKiemMaTatCa_(), s = tqnKiemMaText_(all), n = tqnSuaMa_(all); Logger.log('Đã sửa ' + n + ' ô\n' + s); return n; } finally { lock.releaseLock(); }
}

/* ================= Đồng bộ mã theo Ngày trình + Rev (19/11) =================
 * RFI: mã = <ngày trình>-COA-TBC-RFI-<số> rev<Rev>. Khi BPTC gõ/sửa "Ngày trình duyệt" hoặc "Rev",
 * phần ngày + rev trong mã tự đổi theo (giữ nguyên số thứ tự). Mã mới cấp mà dòng đã có ngày trình cũng lấy ngày đó.
 * Chỉ áp cho tab có cột "Ngày trình duyệt" (ngày) và mẫu mã có {rr} (rev). SD không có cột này → không đổi.
 */
function tqnDongBoMa_(sh, cfg, type, r1, r2) {
  var pat = cfg.codePattern(type); if (!pat) return [];
  var lastC = Math.min(sh.getLastColumn(), 40), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
  var m = tqnDetailMap_(hdr), cNgay = tqnFindCol_(hdr, ['Ngày trình duyệt']), coRev = /\{rr\}/.test(pat) && m.rev >= 0;
  if (cNgay < 0 && !coRev) return [];
  var from = Math.max(r1, cfg.firstRow), to = Math.min(r2, sh.getLastRow()); if (to < from) return [];
  var vals = sh.getRange(from, 1, to - from + 1, lastC).getValues(), out = [];
  vals.forEach(function (r, i) {
    var cu = tqnCleanCode_(r[m.code]); if (!cu) return;
    var re = tqnPatternRegex_(pat, { goi: ((/COA\.([A-Z]+)-/i.exec(cu) || [])[1] || 'NP'), nhom: ((/-RFA-([A-Z0-9]+)-\d/i.exec(cu) || [])[1] || 'COS') }, cfg.digits);
    var mm = re.exec(cu); if (!mm) return;                        // mã không theo mẫu → không đụng
    var v = { goi: ((/COA\.([A-Z]+)-/i.exec(cu) || [])[1] || 'NP').toUpperCase(), nhom: ((/-RFA-([A-Z0-9]+)-\d/i.exec(cu) || [])[1] || 'COS').toUpperCase(), nnn: parseInt(mm[1], 10) };
    v.yymmdd = cu.slice(0, 6);
    var nd = cNgay >= 0 ? r[cNgay] : '';
    if (tqnIsDate_(nd)) v.yymmdd = tqnYymmdd_(nd, cfg.timezone);
    v.rr = parseInt(tqnRevFromNo_(cu) || cfg.revStart, 10);
    if (coRev) { var rv = parseInt(String(r[m.rev] === null || r[m.rev] === undefined ? '' : r[m.rev]).trim(), 10); if (!isNaN(rv)) v.rr = rv; }
    var moi = tqnFill_(pat, v, cfg);
    if (moi !== cu && moi !== String(r[m.code]).trim()) out.push({ row: from + i, cu: String(r[m.code]).trim(), moi: moi, subject: String(r[m.subject] || ''), nnn: v.nnn, rev: ('0' + v.rr).slice(-2) });
  });
  if (out.length) {
    out.forEach(function (x) { sh.getRange(x.row, m.code + 1).setValue(x.moi).setNote('TQN tự cập nhật lúc ' + tqnGioNgay_(cfg) + ' theo Ngày trình / Rev.\nMã trước: ' + x.cu); });
    var lg = tqnOwnSheet_(TQN.PREFIX + 'CAP_MA', TQN_CAPMA_HEADERS);
    lg.getRange(lg.getLastRow() + 1, 1, out.length, 10).setValues(out.map(function (x) { return [new Date(), tqnUser_() || 'tự động', type, x.moi, x.nnn, x.rev, x.subject, x.cu, sh.getName() + '!' + x.row, 'ĐỒNG BỘ NGÀY/REV']; }));
    tqnAudit_('DONG_BO_MA', out.map(function (x) { return x.cu + ' → ' + x.moi; }).join('; '), { tab: sh.getName() });
  }
  return out;
}

/* ================= Chuyên nghiệp hóa (19/11) =================
 * - Ô mã tự cấp/tự đổi có ghi chú (giờ, mã trước) để truy vết.
 * - Ai gõ tay vào cột Mã: tự bỏ khoảng trắng thừa; trùng số / sai mẫu → tô đỏ ô + ghi chú + thông báo (không tự xóa).
 * - Cột Mã: bảo vệ kiểu CẢNH BÁO (vẫn sửa được, nhưng Google hỏi lại) + ghi chú hướng dẫn ở tiêu đề.
 * - Định dạng có điều kiện: mã trùng số → nền đỏ nhạt.
 */
function tqnGioNgay_(cfg) { return Utilities.formatDate(new Date(), cfg.timezone, 'HH:mm dd/MM/yyyy'); }

/** Kiểm các ô thuộc cột Mã vừa bị gõ tay. Trả danh sách cảnh báo. */
function tqnKiemOGoTay_(sh, cfg, type, rg) {
  var lastC = Math.min(sh.getLastColumn(), 40), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
  var m = tqnDetailMap_(hdr), c = m.code + 1;
  if (rg.getColumn() > c || rg.getLastColumn() < c) return [];
  var r1 = Math.max(rg.getRow(), cfg.firstRow), r2 = rg.getLastRow(); if (r2 < r1) return [];
  var o = sh.getRange(r1, c, r2 - r1 + 1, 1), go = o.getValues(), pat = cfg.codePattern(type), out = [];
  var re = /-(RFI|RFA|RFC|SDS|MSS|MAS|OTS)-/i, lastR = sh.getLastRow();
  var tatCa = sh.getRange(cfg.firstRow, c, lastR - cfg.firstRow + 1, 1).getValues().map(function (x) { return String(x[0] || ''); });
  go.forEach(function (x, i) {
    var row = r1 + i, raw = String(x[0] === null || x[0] === undefined ? '' : x[0]), cell = sh.getRange(row, c);
    if (!raw.trim() || !re.test(raw)) return;                // để trống / tiêu đề mục → bỏ qua
    var sach = tqnCleanCode_(raw), loi = [];
    if (sach !== raw.trim()) { cell.setValue(sach); out.push('Dòng ' + row + ': đã bỏ khoảng trắng thừa → ' + sach); }
    var vals = { goi: ((/COA\.([A-Z]+)-/i.exec(sach) || [])[1] || 'NP'), nhom: ((/-RFA-([A-Z0-9]+)-\d/i.exec(sach) || [])[1] || 'COS') };
    if (pat && !tqnPatternRegex_(pat, vals, cfg.digits).test(sach)) loi.push('không đúng mẫu ' + pat);
    var k = tqnMaKhoa_(sach), trung = [];
    tatCa.forEach(function (t, j) { if (cfg.firstRow + j !== row && t && tqnMaKhoa_(t) === k) trung.push(cfg.firstRow + j); });
    if (trung.length) loi.push('trùng số với dòng ' + trung.join(', '));
    if (loi.length) {
      cell.setBackground('#FDE2E1').setNote('TQN cảnh báo ' + tqnGioNgay_(cfg) + ': mã gõ tay ' + loi.join('; ') + '.\nNên xóa ô này để TQN tự cấp số đúng.');
      out.push('Dòng ' + row + ': ' + sach + ' – ' + loi.join('; '));
    } else if (cell.getBackground() === '#fde2e1') { cell.setBackground(null).setNote(''); }
  });
  return out;
}

/** Cài giao diện chuyên nghiệp cho cột Mã ở các tab tự cấp mã (bấm nhiều lần không trùng). */
function tqnLamDepCotMa() {
  var cfg = tqnCodeCfg_(), kq = [];
  cfg.detailSheets.forEach(function (n) {
    var type = tqnAutoTypeOfTab_(cfg, n), sh = type && tqnSs_().getSheetByName(n);
    if (!sh) return;
    var lastC = Math.min(sh.getLastColumn(), 40), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
    var m = tqnDetailMap_(hdr), c = m.code + 1, L = tqnColLetter_(c), f = cfg.firstRow, maxR = Math.max(sh.getMaxRows(), f);
    // 1. ghi chú hướng dẫn ở ô tiêu đề "Mã hồ sơ"
    for (var r = 0; r < hdr.length; r++) if (/ma ho so/.test(tqnNoAccent_(String(hdr[r][m.code] || '')))) {
      sh.getRange(cfg.hdrFrom + r, c).setNote('TỰ ĐỘNG (TQN): gõ "Nội dung chi tiết" ở dòng mới → mã tự cấp số tiếp theo.' + (/\{rr\}/.test(cfg.codePattern(type)) ? '\nSửa Ngày trình duyệt / Rev → ngày + rev trong mã tự đổi theo.' : '') + '\nKhông cần gõ tay. Gõ tay sai mẫu / trùng số sẽ bị tô đỏ.\nMẫu: ' + cfg.codePattern(type));
      break;
    }
    // 2. bảo vệ kiểu cảnh báo cho cột mã (thay bản cũ của TQN)
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { if (/^TQN /.test(p.getDescription())) p.remove(); });
    sh.getRange(f, c, maxR - f + 1, 1).protect().setWarningOnly(true).setDescription('TQN – cột Mã hồ sơ tự động, không cần gõ tay');
    // 3. tô đỏ nhạt mã trùng số (bỏ ngày đầu + rev cuối khi so)
    var key = function (ref) { return 'REGEXREPLACE(UPPER(TRIM(' + ref + '&"")),"^\\d{6}\\s*-\\s*|\\s*-?\\s*REV\\s*\\d+$","")'; };
    var fm = '=AND(REGEXMATCH(' + L + f + '&"","-(SDS|RFI|RFA|RFC|MSS|MAS|OTS)-"),SUMPRODUCT(--(' + key('$' + L + '$' + f + ':$' + L + '$' + maxR) + '=' + key(L + f) + '))>1)';
    var fmOk = '=REGEXMATCH(' + L + f + '&"","^\\d{6}-COA")';
    var rules = sh.getConditionalFormatRules().filter(function (rl) { var b = rl.getBooleanCondition(); var v = String(b && b.getCriteriaValues()[0] || ''); return !(v.indexOf('-(SDS|RFI|RFA') >= 0 || v === fmOk); });
    var vung = [sh.getRange(f, c, maxR - f + 1, 1)];
    rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fmOk).setBackground('#E8F5F2').setFontColor('#0B5F57').setRanges(vung).build());   // mã hợp lệ: nền xanh ngọc nhạt = ô tự động
    rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fm).setBackground('#FDE2E1').setFontColor('#B42318').setRanges(vung).build());       // trùng số: đỏ (ưu tiên trên)
    sh.setConditionalFormatRules(rules);
    // 4. tiêu đề cột Mã màu xanh ngọc (giống tiêu đề Ghi chú tab thời tiết) để nhìn là biết cột tự động
    for (var r2 = 0; r2 < hdr.length; r2++) if (/ma ho so/.test(tqnNoAccent_(String(hdr[r2][m.code] || '')))) { sh.getRange(cfg.hdrFrom + r2, c).setBackground('#0F6E66').setFontColor('#FFFFFF'); break; }
    // 5. huy hiệu "MÃ HỒ SƠ TỰ ĐỘNG · đang bật" nổi ở dòng 1 (ảnh, không ghi vào ô), bấm → xem trạng thái
    try {
      sh.getImages().forEach(function (im) { try { if (im.getAltTextTitle() === 'TQN_BADGE_MA') im.remove(); } catch (e) { } });
      var col = (m.owner >= 0 ? m.owner : m.subject + 1) + 1;
      var im = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(TQN_BADGE_MA), 'image/png', 'tqn_ma_tu_dong.png'), col, 1, 4, 2);
      im.setWidth(273).setHeight(26); im.setAltTextTitle('TQN_BADGE_MA'); im.assignScript('tqnMenuTrangThaiMa');
    } catch (e) { }
    kq.push(n);
  });
  tqnAudit_('LAM_DEP_COT_MA', kq.join(', '), {});
  return kq;
}

/** Bấm huy hiệu: trạng thái tự cấp mã + 5 mã gần nhất. */
function tqnMenuTrangThaiMa() {
  var on = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'tqnTuDongMaKhiSua'; });
  var cfg = tqnCodeCfg_(), tabs = cfg.detailSheets.filter(function (n) { return tqnAutoTypeOfTab_(cfg, n); });
  var lg = tqnSs_().getSheetByName(TQN.PREFIX + 'CAP_MA'), gan = [];
  if (lg && lg.getLastRow() >= 2) {
    var n = Math.min(5, lg.getLastRow() - 1), v = lg.getRange(lg.getLastRow() - n + 1, 1, n, 10).getValues().reverse();
    gan = v.map(function (r) { return '• ' + (tqnIsDate_(r[0]) ? Utilities.formatDate(r[0], cfg.timezone, 'HH:mm dd/MM') : '') + '  ' + r[3] + '  (' + r[9] + ')'; });
  }
  var msg = (on ? '✔ ĐANG BẬT' : '✖ ĐANG TẮT – menu TQN Register → Tự cấp mã: BẬT') + ' cho tab: ' + tabs.join(', ') +
    '\n\nCách dùng: gõ "Nội dung chi tiết" ở dòng mới → Mã tự cấp số tiếp theo (ô xanh ngọc).' +
    '\nRFI: sửa Ngày trình duyệt / Rev → ngày + rev trong mã tự đổi.\nGõ tay sai mẫu / trùng số → ô đỏ.' +
    (gan.length ? '\n\nGần nhất:\n' + gan.join('\n') : '');
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('TQN – Mã hồ sơ tự động', msg + '\n\nKiểm toàn bộ mã gõ tay bây giờ?', ui.ButtonSet.YES_NO) === ui.Button.YES) tqnMenuKiemMa();
}
var TQN_BADGE_MA = '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';

/** Huy hiệu "Mã hồ sơ tự động" thu nhỏ vừa dòng 1, không đè hàng nút ở dòng 2 (lỗi thấy 19/11 22:21). */
function tqnSuaHuyHieuMa() {
  var cfg = tqnConfig_(), n = 0;
  cfg.detailSheets.forEach(function (name) {
    var sh = tqnSs_().getSheetByName(name); if (!sh) return;
    sh.getImages().forEach(function (im) {
      try { if (im.getAltTextTitle() === 'TQN_BADGE_MA') { im.setWidth(273).setHeight(26); im.setAnchorCell(im.getAnchorCell().offset(1 - im.getAnchorCell().getRow(), 0)).setAnchorCellYOffset(2); n++; } } catch (e) { }
    });
  });
  Logger.log('Đã thu nhỏ ' + n + ' huy hiệu');
  return n;
}
