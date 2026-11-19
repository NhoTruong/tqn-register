/**
 * TienDo.gs – Tab "TQN Tiến độ – hồ sơ chặn": gắn mỗi bản vẽ shop (SD) với hạng mục thi công + ngày dự kiến
 * → biết hồ sơ nào đang CHẶN việc ngoài công trường.
 * - Tự lấy danh sách từ tab SD (mã, nội dung, hạng mục, khối RE4/RE5/Công viên). Anh chỉ điền ô VÀNG: ngày thi công, việc ngoài trời.
 * - Cột Tình trạng là công thức nối thẳng sang SD (luôn mới). Cột "Chặn thi công?" tự tính: ✔ đủ hồ sơ / ⛔ chặn / ⚠ cần duyệt trước.
 * - Bấm làm mới nhiều lần: GIỮ nguyên những gì anh đã điền, chỉ thêm dòng mã mới.
 */
var TQN_TD_TEN = 'TQN Tiến độ – hồ sơ chặn';
var TQN_TD_COT = ['Mã hồ sơ', 'Nội dung bản vẽ', 'Hạng mục thi công', 'Khối', 'Tình trạng hồ sơ (tự lấy từ SD)', 'Ngày thi công dự kiến', 'Việc ngoài trời?', 'Chặn thi công?', 'Ghi chú'];
var TQN_TD_DONG1 = 4;   // dòng tiêu đề cột

function tqnHangMucTuNoiDung_(nd) {
  var s = String(nd || '').replace(/\s+/g, ' ').trim();
  s = s.replace(/^(bản vẽ\s+)?(chi tiết\s+)?/i, '').replace(/^hạng mục\s+(kết cấu\s+)?/i, '').replace(/\s*REHAUS\s*-?\s*RE\s*\d+\s*$/i, '').replace(/\s*\(LD-\d+\)\s*$/i, '');
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function tqnKhoiTuNoiDung_(nd, ma) {
  var m = /RE\s*(\d+)/i.exec(String(nd || ''));
  if (m) return 'RE' + m[1];
  if (/COA\.HTKT/i.test(String(ma || ''))) return 'Công viên';
  return '';
}

/** Tạo / làm mới tab. Trả số dòng mới thêm. */
function tqnTaoSheetTienDo() {
  var ss = tqnSs_(), cfg = tqnCodeCfg_(), sd = ss.getSheetByName('SD');
  if (!sd) throw new Error('Không thấy tab SD');
  var lastC = Math.min(sd.getLastColumn(), 40), m = tqnDetailMap_(sd.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues());
  var ds = [], lastR = sd.getLastRow();
  if (lastR >= cfg.firstRow) sd.getRange(cfg.firstRow, 1, lastR - cfg.firstRow + 1, lastC).getValues().forEach(function (r) {
    var ma = tqnCleanCode_(r[m.code]); if (!/-SDS-/i.test(ma)) return;
    ds.push({ ma: ma, nd: String(r[m.subject] || '').trim() });
  });
  var sh = ss.getSheetByName(TQN_TD_TEN), cu = {}, moi = !sh;
  if (moi) {
    var dangMo = null; try { dangMo = ss.getActiveSheet(); } catch (e) { }
    sh = ss.insertSheet(TQN_TD_TEN, 2);
    try { if (dangMo) ss.setActiveSheet(dangMo); } catch (e) { }
  } else if (sh.getLastRow() > TQN_TD_DONG1) {
    sh.getRange(TQN_TD_DONG1 + 1, 1, sh.getLastRow() - TQN_TD_DONG1, TQN_TD_COT.length).getValues().forEach(function (r) { if (r[0]) cu[tqnMaKhoa_(r[0])] = r; });
  }
  // dữ liệu: giữ ô anh đã điền (C, D, F, G, I), cập nhật mã/nội dung theo SD
  var codeCol = tqnColLetter_(m.code + 1), stCol = tqnColLetter_((m.status >= 0 ? m.status : m.code + 3) + 1), them = 0;
  var rows = ds.map(function (x, i) {
    var r = TQN_TD_DONG1 + 1 + i, c = cu[tqnMaKhoa_(x.ma)];
    if (!c) them++;
    return [x.ma, x.nd, c && c[2] ? c[2] : tqnHangMucTuNoiDung_(x.nd), c && c[3] ? c[3] : tqnKhoiTuNoiDung_(x.nd, x.ma),
      '=IFERROR(INDEX(SD!$' + stCol + ':$' + stCol + ',MATCH(A' + r + ',SD!$' + codeCol + ':$' + codeCol + ',0))&"","")',
      c ? c[5] : '', c && c[6] ? c[6] : '',
      '=IF(F' + r + '="","",IF(REGEXMATCH(LOWER(E' + r + '),"đã duyệt|duyệt có ý kiến|đã hoàn thành"),"✔ Đủ hồ sơ",IF(F' + r + '-TODAY()<0,"⛔ Đã tới ngày – chưa duyệt",IF(F' + r + '-TODAY()<=7,"⛔ Chặn – còn "&(F' + r + '-TODAY())&" ngày","⚠ Cần duyệt trước "&TEXT(F' + r + ',"dd/mm")))))',
      c ? c[8] : ''];
  });
  var n = rows.length, W = TQN_TD_COT.length;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  if (sh.getLastRow() > TQN_TD_DONG1) sh.getRange(TQN_TD_DONG1 + 1, 1, sh.getLastRow() - TQN_TD_DONG1, W).clearContent();
  // tiêu đề
  sh.getRange(1, 1, 1, W).merge().setValue('TIẾN ĐỘ THI CÔNG ↔ HỒ SƠ SHOP – hồ sơ nào đang chặn việc ngoài công trường')
    .setBackground('#1B3F7A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(14).setVerticalAlignment('middle');
  sh.getRange(2, 1, 1, W).merge().setValue('Anh chỉ điền ô VÀNG: Ngày thi công dự kiến (ngày bắt đầu hạng mục cần bản vẽ này) và Việc ngoài trời (Có/Không). ' +
    'Cột Tình trạng tự lấy từ SD · cột "Chặn thi công?" tự tính: ⛔ còn ≤ 7 ngày mà chưa duyệt · ⚠ cần duyệt trước ngày · ✔ đủ hồ sơ. Bấm làm mới không mất dữ liệu đã điền.')
    .setBackground('#EEF3FB').setFontColor('#1F2937').setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  sh.setRowHeight(1, 36); sh.setRowHeight(2, 40); sh.setRowHeight(3, 6);
  sh.getRange(TQN_TD_DONG1, 1, 1, W).setValues([TQN_TD_COT]).setBackground('#1B3F7A').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(TQN_TD_DONG1, 6, 1, 2).setBackground('#B45309');                 // cột cần điền: cam
  sh.getRange(TQN_TD_DONG1, 8).setBackground('#0F6E66');                      // cột tự tính: xanh ngọc
  sh.setRowHeight(TQN_TD_DONG1, 40);
  if (n) {
    var vung = sh.getRange(TQN_TD_DONG1 + 1, 1, n, W);
    vung.setValues(rows).setFontSize(10).setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
    var nen = []; for (var i = 0; i < n; i++) { var z = i % 2 ? '#F8FAFC' : '#FFFFFF'; nen.push([z, z, z, z, z, '#FFF8DB', '#FFF8DB', '#E8F5F2', z]); }
    vung.setBackgrounds(nen);
    sh.getRange(TQN_TD_DONG1 + 1, 6, n, 1).setNumberFormat('dd/MM/yyyy').setHorizontalAlignment('center')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Ngày bắt đầu hạng mục thi công cần bản vẽ này (dd/mm/yyyy)').build());
    sh.getRange(TQN_TD_DONG1 + 1, 7, n, 1).setHorizontalAlignment('center')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Có', 'Không'], true).setAllowInvalid(false).build());
    sh.getRange(TQN_TD_DONG1 + 1, 4, n, 1).setHorizontalAlignment('center');
    sh.getRange(TQN_TD_DONG1 + 1, 1, n, 1).setFontColor('#0B5F57').setFontWeight('bold');
    // màu cột Chặn
    var c8 = sh.getRange(TQN_TD_DONG1 + 1, 8, n, 1), R = sh.getConditionalFormatRules().filter(function (rl) { return rl.getRanges().every(function (g) { return g.getColumn() !== 8; }); });
    R.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⛔').setBackground('#FDE2E1').setFontColor('#B42318').setBold(true).setRanges([c8]).build());
    R.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setBackground('#FEF3C7').setFontColor('#92400E').setRanges([c8]).build());
    R.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setBackground('#DCFCE7').setFontColor('#166534').setRanges([c8]).build());
    sh.setConditionalFormatRules(R);
  }
  [190, 330, 250, 80, 150, 120, 95, 190, 220].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(TQN_TD_DONG1); sh.setFrozenColumns(1); sh.setTabColor('#B45309');
  try { sh.setHiddenGridlines(true); } catch (e) { }
  tqnAudit_('TAO_TIEN_DO', n + ' dòng, thêm mới ' + them, {});
  return { tong: n, them: them };
}

/** Đọc tab tiến độ cho bảng Quản lý / thẻ Hôm nay: danh sách đang chặn. */
function tqnHoSoChan_() {
  var sh = tqnSs_().getSheetByName(TQN_TD_TEN); if (!sh || sh.getLastRow() <= TQN_TD_DONG1) return [];
  return sh.getRange(TQN_TD_DONG1 + 1, 1, sh.getLastRow() - TQN_TD_DONG1, TQN_TD_COT.length).getDisplayValues()
    .filter(function (r) { return /^⛔/.test(r[7]); })
    .map(function (r) { return { ma: r[0], hangMuc: r[2], khoi: r[3], ngay: r[5], ngoaiTroi: r[6], chan: r[7] }; });
}

function tqnMenuTienDo() {
  var k = tqnTaoSheetTienDo();
  SpreadsheetApp.getUi().alert('TQN – Tiến độ ↔ hồ sơ', 'Tab "' + TQN_TD_TEN + '": ' + k.tong + ' bản vẽ SD (thêm mới ' + k.them + ').\n\nAnh điền ô VÀNG: Ngày thi công dự kiến + Việc ngoài trời. Cột "Chặn thi công?" tự báo hồ sơ nào đang chặn việc.', SpreadsheetApp.getUi().ButtonSet.OK);
  try { tqnSs_().setActiveSheet(tqnSs_().getSheetByName(TQN_TD_TEN)); } catch (e) { }
}
