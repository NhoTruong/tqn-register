/**
 * QuanLy.gs – Quản lý chặt theo MẪU PHIẾU của dự án (ai cũng cập nhật được, hệ thống tự kiểm + ghi nhật ký).
 *
 *  - Tab TQN_MAU_PHIEU : mỗi loại hồ sơ có các cấp duyệt, mỗi lựa chọn trên phiếu → kết quả.
 *                        Qua dự án khác chỉ sửa bảng này (hoặc để AI đọc phiếu trắng điền giúp), không sửa code.
 *  - Khối "TQN QUẢN LÝ" thêm vào BÊN PHẢI mỗi sổ (không đụng cột gốc):
 *        [kết quả từng cấp theo phiếu] [Chuyển tới] | Ai đang giữ | Link OneDrive | Kiểm tra
 *  - Tab TQN_NHAT_KY   : mọi lần sửa (ai, lúc nào, ô nào, cũ → mới). Chỉ chủ file sửa được tab này.
 *  - Không chặn cứng: sai luật thì tô đỏ + ghi "Kiểm tra" + ghi nhật ký.
 *
 *  Phần LÕI (tqnQlKhop_, tqnQlDanhGia_, tqnQlDocMau_) là hàm thuần, không gọi SpreadsheetApp,
 *  để sau này đưa nguyên sang web app.
 */
var TQN_QL = {
  SH_MAU: 'TQN_MAU_PHIEU',
  SH_LOG: 'TQN_NHAT_KY',
  NHAN: 'TQN QUẢN LÝ',
  MAU_HDR: ['Loại hồ sơ', 'Thứ tự cấp', 'Tên cấp', 'Mã', 'Nguyên văn trên phiếu', 'Kết quả', 'Ghi chú'],
  LOG_HDR: ['Thời gian', 'Người sửa', 'Tab', 'Ô', 'Mã hồ sơ', 'Cột', 'Giá trị cũ', 'Giá trị mới', 'Hệ thống ghi nhận'],
  // Kết quả hợp lệ trong cột "Kết quả" của mẫu phiếu
  KQ: {
    TIEP: 'Chuyển cấp sau duyệt',
    DONG: 'Đóng – đạt',
    DONG_YK: 'Đóng – đạt có ý kiến',
    TRINH_LAI: 'Nhà thầu sửa & trình lại',
    TU_CHOI: 'Bị từ chối – nhà thầu xử lý',
    BO_SUNG: 'Nhà thầu bổ sung thông tin',
    CHUYEN: 'Chuyển tiếp bên khác'
  },
  LINK_RE: /^https:\/\/([a-z0-9-]+\.)*(sharepoint\.com|1drv\.ms|onedrive\.live\.com|onedrive\.com)\//i
};

/** Mẫu 2 cấp TVGS → CĐT (dùng cho dự án có tư vấn). */
var TQN_MAU_PHIEU_2CAP = (function () {
  var abcd = function (type, note) {
    return [
      [type, 1, 'TVGS', 'A', 'Chấp thuận yêu cầu', 'TIEP', note || 'Phiếu ' + type + ' – khối TVQLDA/TVGS'],
      [type, 1, 'TVGS', 'B', 'Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới', 'TIEP', ''],
      [type, 1, 'TVGS', 'C', 'Vui lòng xem ý kiến phản hồi và đệ trình lại', 'TRINH_LAI', ''],
      [type, 1, 'TVGS', 'D', 'Từ chối chấp thuận yêu cầu', 'TU_CHOI', ''],
      [type, 2, 'CĐT', 'A', 'Chấp thuận yêu cầu', 'DONG', 'Khối CHỦ ĐẦU TƯ/BQLXD – kết quả cuối'],
      [type, 2, 'CĐT', 'B', 'Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới', 'DONG_YK', ''],
      [type, 2, 'CĐT', 'C', 'Vui lòng xem ý kiến phản hồi và đệ trình lại', 'TRINH_LAI', ''],
      [type, 2, 'CĐT', 'D', 'Từ chối chấp thuận yêu cầu', 'TU_CHOI', '']
    ];
  };
  return abcd('RFA').concat(abcd('RFC'), abcd('SDS', 'SD tạm dùng phiếu RFA – đổi nếu SD có phiếu riêng'), [
    ['RFI', 1, 'Trả lời RFI', '1', 'Xem trả lời bên dưới', 'DONG', 'Phiếu RFI – HÀNH ĐỘNG KIẾN NGHỊ'],
    ['RFI', 1, 'Trả lời RFI', '2', 'Vui lòng cung cấp thêm thông tin', 'BO_SUNG', ''],
    ['RFI', 1, 'Trả lời RFI', '3', 'Chuyển tiếp tới', 'CHUYEN', 'Ghi bên nhận ở cột "Chuyển tới"']
  ]);
})();

/** Mẫu MẶC ĐỊNH dự án Coastal Quảng Ngãi: KHÔNG có tư vấn – chỉ CĐT duyệt (A/B/C/D nguyên văn phiếu). */
var TQN_MAU_PHIEU_DEFAULT = (function () {
  var cdt = function (type, note) {
    return [
      [type, 1, 'CĐT', 'A', 'Chấp thuận yêu cầu', 'DONG', note || 'Phiếu ' + type + ' – Chủ đầu tư/BQLXD (dự án không có tư vấn)'],
      [type, 1, 'CĐT', 'B', 'Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới', 'DONG_YK', ''],
      [type, 1, 'CĐT', 'C', 'Vui lòng xem ý kiến phản hồi và đệ trình lại', 'TRINH_LAI', ''],
      [type, 1, 'CĐT', 'D', 'Từ chối chấp thuận yêu cầu', 'TU_CHOI', '']
    ];
  };
  return cdt('RFA').concat(cdt('RFC'), cdt('SDS', 'SD tạm dùng phiếu RFA'), [
    ['RFI', 1, 'CĐT trả lời', '1', 'Xem trả lời bên dưới', 'DONG', 'Phiếu RFI – HÀNH ĐỘNG KIẾN NGHỊ'],
    ['RFI', 1, 'CĐT trả lời', '2', 'Vui lòng cung cấp thêm thông tin', 'BO_SUNG', ''],
    ['RFI', 1, 'CĐT trả lời', '3', 'Chuyển tiếp tới', 'CHUYEN', 'Ghi bên nhận ở cột "Chuyển tới"']
  ]);
})();

/* ===================== LÕI (hàm thuần) ===================== */

/** Đọc bảng mẫu phiếu (mảng dòng, không gồm tiêu đề) → { LOAI: { levels:[{name, order, options:[{code,text,kq,label}]}], forward } } */
function tqnQlDocMau_(rows) {
  var prof = {}, loi = [];
  (rows || []).forEach(function (r, i) {
    var type = String(r[0] || '').trim().toUpperCase(), ord = parseInt(r[1], 10), lv = String(r[2] || '').trim(),
      code = String(r[3] || '').trim(), text = String(r[4] || '').trim(), kq = String(r[5] || '').trim().toUpperCase();
    if (!type && !lv && !text) return;
    if (!type || !lv || (!code && !text) || isNaN(ord)) { loi.push('Dòng ' + (i + 2) + ': thiếu Loại / Thứ tự / Tên cấp / Mã'); return; }
    if (!TQN_QL.KQ[kq]) { loi.push('Dòng ' + (i + 2) + ': Kết quả "' + r[5] + '" không hợp lệ'); return; }
    var p = prof[type] = prof[type] || { levels: [], forward: false };
    var L = null; p.levels.forEach(function (x) { if (x.name === lv) L = x; });
    if (!L) { L = { name: lv, order: ord, options: [] }; p.levels.push(L); }
    L.options.push({ code: code, text: text, kq: kq, label: (code ? code + ' · ' : '') + text });
    if (kq === 'CHUYEN') p.forward = true;
  });
  Object.keys(prof).forEach(function (t) { prof[t].levels.sort(function (a, b) { return a.order - b.order; }); });
  prof._loi = loi;
  return prof;
}

function tqnQlBo_(s) { return tqnNoAccent_(String(s == null ? '' : s)).replace(/\s+/g, ' ').trim(); }

/** Khớp giá trị người nhập với lựa chọn của một cấp: nhận "A", "A · Chấp thuận…", hoặc nguyên văn. null = không khớp. */
function tqnQlKhop_(val, level) {
  var v = tqnQlBo_(val);
  if (!v) return null;
  var head = v.split(' · ')[0].split(' - ')[0].split('.')[0].trim();
  for (var i = 0; i < level.options.length; i++) {
    var o = level.options[i];
    if (v === tqnQlBo_(o.label) || v === tqnQlBo_(o.text)) return o;
  }
  for (var j = 0; j < level.options.length; j++) if (level.options[j].code && head === tqnQlBo_(level.options[j].code)) return level.options[j];
  return null;
}

/**
 * Đánh giá một dòng. rec = { noiDung, daTrinh(bool), tinhTrang, picks:[giá trị từng cấp], chuyenToi, link }
 * Trả { giu: 'chữ hiện ở cột Ai đang giữ', nhom: 'DONG'|'NT'|'CHO'|'SOAN'|'', canhBao: [..] }
 */
function tqnQlDanhGia_(rec, p) {
  if (!String(rec.noiDung || '').trim()) return { giu: '', nhom: '', canhBao: [] };
  var cb = [], st = tqnQlBo_(rec.tinhTrang), giu = null, nhom = null, ketThuc = null;
  var stDong = /da duyet|da hoan thanh|hoan tat|da dong|da tra loi/.test(st);
  var stTra = /khong duyet|tra ve|trinh lai/.test(st);
  var picks = rec.picks || [], coChon = picks.some(function (x) { return String(x == null ? '' : x).trim(); });
  for (var i = 0; i < p.levels.length; i++) {
    var L = p.levels[i], raw = picks[i], o = tqnQlKhop_(raw, L);
    if (String(raw || '').trim() && !o) cb.push('"' + String(raw).trim() + '" không có trong mẫu phiếu cấp ' + L.name);
    if (ketThuc) { if (o) cb.push(L.name + ' có kết quả nhưng hồ sơ đã ' + ketThuc + ' ở cấp trước'); continue; }
    if (!o) {
      // cấp này chưa có kết quả – cấp sau mà có thì sai thứ tự
      for (var k = i + 1; k < p.levels.length; k++) if (String(picks[k] || '').trim()) cb.push(p.levels[k].name + ' có kết quả nhưng ' + L.name + ' còn trống');
      if (coChon && rec.daTrinh) { giu = '⏳ ' + L.name; nhom = 'CHO'; }
      break;
    }
    if (o.kq === 'TIEP') { if (i === p.levels.length - 1) { giu = '✔ Đóng'; nhom = 'DONG'; ketThuc = 'đóng'; } continue; }
    if (o.kq === 'DONG') { giu = '✔ Đóng'; nhom = 'DONG'; ketThuc = 'đóng'; }
    else if (o.kq === 'DONG_YK') { giu = '✔ Đóng (có ý kiến)'; nhom = 'DONG'; ketThuc = 'đóng'; }
    else if (o.kq === 'TRINH_LAI') { giu = '↩ Nhà thầu – sửa & trình lại (' + L.name + ' ' + o.code + ')'; nhom = 'NT'; ketThuc = 'trả về'; }
    else if (o.kq === 'TU_CHOI') { giu = '↩ Nhà thầu – bị từ chối (' + L.name + ' ' + o.code + ')'; nhom = 'NT'; ketThuc = 'trả về'; }
    else if (o.kq === 'BO_SUNG') { giu = '↩ Nhà thầu – bổ sung thông tin'; nhom = 'NT'; ketThuc = 'trả về'; }
    else if (o.kq === 'CHUYEN') {
      var to = String(rec.chuyenToi || '').trim();
      giu = '⏳ ' + (to || 'bên được chuyển (chưa ghi)'); nhom = 'CHO'; ketThuc = 'chuyển tiếp';
      if (!to) cb.push('Chọn "Chuyển tiếp" nhưng chưa ghi Chuyển tới ai');
    }
  }
  if (!giu) {                                   // chưa có kết quả theo phiếu → dùng Tình trạng cũ
    if (!coChon && stDong) { giu = '✔ Đóng (theo Tình trạng)'; nhom = 'DONG'; }
    else if (!coChon && stTra) { giu = '↩ Nhà thầu (theo Tình trạng)'; nhom = 'NT'; }
    else if (rec.daTrinh) { giu = '⏳ ' + (p.levels[0] ? p.levels[0].name : 'Bên duyệt'); nhom = 'CHO'; }
    else { giu = '✎ Nhà thầu – đang soạn'; nhom = 'SOAN'; }
  }
  if (coChon && !rec.daTrinh) cb.push('Có kết quả duyệt nhưng chưa có ngày trình');
  if (coChon && nhom === 'DONG' && stTra) cb.push('Tình trạng "' + rec.tinhTrang + '" trái với kết quả phiếu (đã đóng)');
  if (coChon && nhom === 'DONG' && !stDong) cb.push('Phiếu đã đóng – Tình trạng nên cập nhật thành "Đã duyệt"');
  if (coChon && nhom === 'NT' && stDong) cb.push('Tình trạng "' + rec.tinhTrang + '" trái với kết quả phiếu (bị trả về)');
  var link = String(rec.link || '').trim();
  if (nhom === 'DONG' && !link) cb.push('Thiếu link OneDrive phiếu đã ký');
  if (link && !TQN_QL.LINK_RE.test(link)) cb.push('Link không phải OneDrive/SharePoint');
  return { giu: giu, nhom: nhom, canhBao: cb };
}

/* ===================== GẮN VÀO SHEET ===================== */

function tqnEnsureMauPhieu_() {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN_QL.SH_MAU);
  if (sh) return sh;
  sh = tqnOwnSheet_(TQN_QL.SH_MAU, TQN_QL.MAU_HDR);
  sh.getRange(2, 1, TQN_MAU_PHIEU_DEFAULT.length, 7).setValues(TQN_MAU_PHIEU_DEFAULT);
  var kq = Object.keys(TQN_QL.KQ);
  try {
    sh.getRange(2, 6, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(kq, true).setAllowInvalid(false)
      .setHelpText(kq.map(function (k) { return k + ' = ' + TQN_QL.KQ[k]; }).join('\n')).build());
    sh.getRange(1, 1, 1, 7).setNote('Sửa bảng này khi qua dự án khác / đổi mẫu phiếu.\nKết quả: ' + kq.map(function (k) { return k + ' = ' + TQN_QL.KQ[k]; }).join('; '));
    sh.setColumnWidth(5, 420); sh.setColumnWidth(7, 280);
  } catch (e) { }
  return sh;
}

function tqnQlMau_() {
  var sh = tqnEnsureMauPhieu_(), n = sh.getLastRow() - 1;
  return tqnQlDocMau_(n > 0 ? sh.getRange(2, 1, n, 7).getValues() : []);
}

/** Loại hồ sơ của tab (theo type_alias) nếu có trong mẫu phiếu. */
function tqnQlLoaiTab_(cfg, prof, name) {
  if (cfg.detailSheets.indexOf(name) < 0) return '';
  var t = String(cfg.alias[name] || name).toUpperCase();
  return prof[t] ? t : '';
}

/** Dòng tiêu đề chính (có "Mã hồ sơ") trong vùng tiêu đề. */
function tqnQlDongTieuDe_(hdr, cfg) {
  for (var r = 0; r < hdr.length; r++) for (var c = 0; c < hdr[r].length; c++) if (tqnQlBo_(hdr[r][c]) === 'ma ho so') return cfg.hdrFrom + r;
  return cfg.hdrFrom;
}

/** Tên các cột của khối quản lý cho một loại hồ sơ. */
function tqnQlTenCot_(p) {
  var a = p.levels.map(function (L) { return L.name + ' · theo phiếu'; });
  if (p.forward) a.push('Chuyển tới');
  return a.concat(['Ai đang giữ', 'Link OneDrive', 'Kiểm tra']);
}

/**
 * Tìm khối TQN QUẢN LÝ trong tab; tao=true thì thêm vào bên phải nếu chưa có.
 * Trả { start (cột 1-based), lv:[cột từng cấp], to, giu, link, kt, map (tqnDetailMap_), hdrRow } hoặc null.
 */
function tqnQlKhoi_(sh, cfg, p, tao) {
  var lastC = Math.max(sh.getLastColumn(), 1), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
  var hr = tqnQlDongTieuDe_(hdr, cfg), row = hdr[hr - cfg.hdrFrom], start = -1;
  for (var c = 0; c < row.length; c++) if (String(row[c]).indexOf(TQN_QL.NHAN) === 0) { start = c + 1; break; }
  var ten = tqnQlTenCot_(p);
  if (start < 0) {
    if (!tao) return null;
    start = lastC + 2;                                    // chừa 1 cột trống ngăn cách
    var need = start + ten.length - 1;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    sh.getRange(hr, start).setValue(TQN_QL.NHAN + ' · tự động');
    sh.getRange(hr + 1, start, 1, ten.length).setValues([ten]);
    try {
      sh.getRange(hr, start, 2, ten.length).setBackground('#0F6E66').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
      ten.forEach(function (t, i) { sh.setColumnWidth(start + i, /Kiểm tra/.test(t) ? 300 : (/Ai đang giữ/.test(t) ? 210 : 150)); });
    } catch (e) { }
  }
  var sub = sh.getRange(hr + 1, start, 1, ten.length).getValues()[0], at = function (name) {
    for (var i = 0; i < sub.length; i++) if (tqnQlBo_(sub[i]) === tqnQlBo_(name)) return start + i;
    return -1;
  };
  var k = { start: start, lv: p.levels.map(function (L) { return at(L.name + ' · theo phiếu'); }), to: at('Chuyển tới'),
    giu: at('Ai đang giữ'), link: at('Link OneDrive'), kt: at('Kiểm tra'), hdrRow: hr };
  k.map = tqnDetailMap_(hdr);
  k.han = tqnQlHan_(hdr, k.map, p);
  return k;
}

/** Nền khối quản lý: thân bảng trắng, kẻ viền nhạt, tiêu đề gộp 1 ô (sổ gốc tô nền tối ngoài bảng). */
function tqnQlNenKhoi_(sh, cfg, k, p) {
  var w = tqnQlTenCot_(p).length, top = k.hdrRow + 2, n = Math.max(sh.getMaxRows() - top + 1, 1);
  try {
    var t = sh.getRange(k.hdrRow, k.start, 1, w);
    if (!t.isPartOfMerge()) t.merge();
    t.setValue(TQN_QL.NHAN + ' · tự động');
    t.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setFontSize(10);
    sh.getRange(k.hdrRow + 1, k.start, 1, w).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10).setWrap(true);
    var ten = tqnQlTenCot_(p);   // cột gọn: không bắt kéo ngang
    ten.forEach(function (x, i) {
      sh.setColumnWidth(k.start + i, /Kiểm tra/.test(x) ? 170 : /Ai đang giữ/.test(x) ? 140 : /Link/.test(x) ? 95 : /Chuyển/.test(x) ? 100 : 120);
    });
    sh.getRange(top, k.start, n, w).setFontSize(10).setWrap(true);
    sh.getRange(top, k.start, n, w).setBackground('#FFFFFF').setFontColor('#1F2937').setFontWeight('normal')
      .setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#D0D5DD', SpreadsheetApp.BorderStyle.SOLID);
    if (cfg.firstRow > top) sh.getRange(top, k.start, cfg.firstRow - top, w).setBackground('#E7F3F1');
    if (k.start > 1) sh.getRange(1, k.start - 1, sh.getMaxRows(), 1).setBackground(null);
  } catch (e) { }
}

/** Ghi lại dropdown + màu cho các cột cấp duyệt (gọi khi cài / đổi mẫu phiếu). */
function tqnQlDinhDang_(sh, cfg, p, k) {
  var n = Math.max(sh.getMaxRows() - cfg.firstRow + 1, 1), rules = sh.getConditionalFormatRules().filter(function (r) {
    var rg = r.getRanges()[0]; return !(rg && rg.getColumn() >= k.start);
  });
  tqnQlNenKhoi_(sh, cfg, k, p);
  var mau = { TIEP: '#D9F2E6', DONG: '#B7E4C7', DONG_YK: '#D8F3DC', TRINH_LAI: '#FFE8A3', TU_CHOI: '#F8C9C4', BO_SUNG: '#FFE8A3', CHUYEN: '#DCE6F7' };
  p.levels.forEach(function (L, i) {
    if (k.lv[i] < 0) return;
    var rg = sh.getRange(cfg.firstRow, k.lv[i], n, 1);
    rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(L.options.map(function (o) { return o.label; }), true)
      .setAllowInvalid(true).setHelpText('Chọn đúng ô đã đánh dấu trên phiếu ' + L.name + ' (mẫu ở tab ' + TQN_QL.SH_MAU + ')').build());
    L.options.forEach(function (o) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(o.label.split(' · ')[0] + (o.code ? ' ·' : ''))
        .setBackground(mau[o.kq] || '#ffffff').setRanges([rg]).build());
    });
  });
  if (k.giu > 0) {
    var g = sh.getRange(cfg.firstRow, k.giu, n, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('· trễ').setBackground('#D92D20').setFontColor('#ffffff').setBold(true).setRanges([g]).build());
    [['✔', '#B7E4C7', '#0B5F57'], ['↩', '#F8C9C4', '#8A1C12'], ['⏳', '#FFE8A3', '#6B4E00'], ['✎', '#ECEFF3', '#475467'], ['⤳', '#F2F4F7', '#98A2B3']].forEach(function (x) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([g]).build());
    });
  }
  if (k.kt > 0) {
    var t = sh.getRange(cfg.firstRow, k.kt, n, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setBackground('#FDE2E1').setFontColor('#8A1C12').setRanges([t]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✔').setFontColor('#0B5F57').setRanges([t]).build());
  }
  sh.setConditionalFormatRules(rules);
  // 2 cột tự tính: cảnh báo khi gõ tay (không khóa cứng) – xóa bản cũ để không chồng nhiều lớp
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (x) { if (/^TQN – cột tự tính/.test(x.getDescription())) x.remove(); });
  [k.giu, k.kt].forEach(function (c) {
    if (c < 0) return;
    var pr = sh.getRange(cfg.firstRow, c, n, 1).protect().setDescription('TQN – cột tự tính, không cần gõ');
    pr.setWarningOnly(true);
  });
}

/** Đọc lại các dòng r1..r2 (số dòng sheet) và ghi cột Ai đang giữ + Kiểm tra. Trả số dòng có cảnh báo. */
/** Hạn từng cấp: lấy "(max N)" trên tiêu đề "Số ngày phản hồi" của khối; không có thì han_mac_dinh (4). */
function tqnQlHan_(hdr, m, p) {
  var md = parseInt((tqnConfig_() || {}).han_mac_dinh, 10); if (isNaN(md)) md = 4;
  var khoi = { 'TVGS': m.tvgs, 'CĐT': m.cdt, 'TVTK': m.tvtk }, han = {};
  p.levels.forEach(function (L) {
    var b = khoi[L.name] || m.cdt, n = md;
    if (b && b.days != null && b.days >= 0) {
      for (var r = 0; r < hdr.length; r++) { var x = String(hdr[r][b.days] || '').match(/max\s*(\d+)/i); if (x) { n = +x[1]; break; } }
    }
    han[L.name] = { ngay: n, khoi: b };
  });
  return han;
}
/** Số ngày trễ khi hồ sơ đang chờ 1 cấp (⏳) quá hạn; mốc = ngày gần nhất đã có (trình / cấp trước trả). */
function tqnQlTre_(giu, row, k, p, homNay) {
  if (!giu || String(giu).charAt(0) !== '⏳' || !k.han) return 0;
  var m = k.map, ten = String(giu).replace(/^⏳\s*/, ''), lv = null;
  p.levels.forEach(function (L) { if (!lv && ten.indexOf(L.name) === 0) lv = L; });
  if (!lv || !k.han[lv.name]) return 0;
  var cols = [m.nt && m.nt.actual, m.nt && m.nt.plan, m.tvgs && lv.name !== 'TVGS' ? m.tvgs.actual : null];
  var moc = null;
  cols.forEach(function (c) { if (c != null && c >= 0 && tqnIsDate_(row[c]) && (!moc || row[c] > moc)) moc = row[c]; });
  if (!moc) return 0;
  var d0 = new Date(moc); d0.setHours(0, 0, 0, 0);
  var cho = Math.round((homNay - d0) / 864e5), qua = cho - k.han[lv.name].ngay;
  return qua > 0 ? qua : 0;
}

/** Tách mã: bỏ ngày đầu (yyMMdd-) và đuôi rev → { goc, rev }. "261116-COA-TBC-RFA-COS-001-rev02" → goc "COA-TBC-RFA-COS-001", rev 2 */
function tqnQlTachRev_(ma) {
  var s = String(ma || '').trim(), x = s.match(/[-\s_]*rev\s*(\d+)\s*$/i);
  if (!s || !x) return null;
  return { goc: s.slice(0, x.index).replace(/^\d{6}-/, '').toUpperCase(), rev: +x[1] };
}
function tqnQlRevMax_(mas) {
  var o = {};
  mas.forEach(function (ma) { var r = tqnQlTachRev_(ma); if (r && (o[r.goc] == null || r.rev > o[r.goc])) o[r.goc] = r.rev; });
  return o;
}
function tqnQlTinhDong_(sh, cfg, p, k, r1, r2) {
  r1 = Math.max(r1 || cfg.firstRow, cfg.firstRow); r2 = Math.min(r2 || sh.getLastRow(), sh.getLastRow());
  if (r2 < r1) return 0;
  var n = r2 - r1 + 1, lastC = Math.max(sh.getLastColumn(), k.kt), v = sh.getRange(r1, 1, n, lastC).getValues();
  var lk = k.link > 0 ? sh.getRange(r1, k.link, n, 1).getRichTextValues() : null, m = k.map, out = [], loi = 0;
  var ngay = [m.nt && m.nt.plan, m.nt && m.nt.actual].filter(function (c) { return c != null && c >= 0; });
  var homNay = new Date(); homNay.setHours(0, 0, 0, 0);
  var revMax = {};
  if (m.code != null && m.code >= 0 && sh.getLastRow() >= cfg.firstRow) revMax = tqnQlRevMax_(sh.getRange(cfg.firstRow, m.code + 1, sh.getLastRow() - cfg.firstRow + 1, 1).getValues().map(function (x) { return x[0]; }));
  for (var i = 0; i < n; i++) {
    var row = v[i], url = '';
    if (lk && lk[i][0]) url = lk[i][0].getLinkUrl() || '';
    if (!url && k.link > 0) url = String(row[k.link - 1] || '');
    var st = tqnQlBo_(row[m.status]);
    var rec = {
      noiDung: row[m.subject], tinhTrang: row[m.status],
      daTrinh: ngay.some(function (c) { return tqnIsDate_(row[c]); }) || /da trinh|cho duyet/.test(st),
      picks: k.lv.map(function (c) { return c > 0 ? row[c - 1] : ''; }), chuyenToi: k.to > 0 ? row[k.to - 1] : '', link: url
    };
    var d = tqnQlDanhGia_(rec, p);
    var rv = m.code != null && m.code >= 0 ? tqnQlTachRev_(row[m.code]) : null;
    if (rv && /^[⏳↩✎]/.test(d.giu) && revMax[rv.goc] > rv.rev) { d.giu = '⤳ Đã thay bằng rev' + ('0' + revMax[rv.goc]).slice(-2); d.canhBao = []; }
    var tre = tqnQlTre_(d.giu, row, k, p, homNay);
    if (tre) d.giu += ' · trễ ' + tre + ' ngày';
    if (d.canhBao.length) loi++;
    out.push([d.giu, d.giu ? (d.canhBao.length ? '⚠ ' + d.canhBao.join('; ') : '✔') : '']);
  }
  sh.getRange(r1, k.giu, n, 1).setValues(out.map(function (x) { return [x[0]]; }));
  sh.getRange(r1, k.kt, n, 1).setValues(out.map(function (x) { return [x[1]]; }));
  return loi;
}

/** Ô Link vừa dán: kiểm OneDrive, hiện gọn "📎 OneDrive", gắn link vào ô Mã (nếu ô Mã chưa có link). */
function tqnQlXuLyLink_(sh, k, r, raw) {
  var cell = sh.getRange(r, k.link), rt = cell.getRichTextValue(), url = (rt && rt.getLinkUrl()) || String(raw || '').trim();
  if (!url) { cell.setBackground(null).setNote(''); return ''; }
  if (!TQN_QL.LINK_RE.test(url)) {
    cell.setBackground('#FDE2E1').setNote('⚠ Không phải link OneDrive/SharePoint. Mở file trên OneDrive → Chia sẻ → Sao chép liên kết.');
    return 'LINK_SAI';
  }
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('📎 OneDrive').setLinkUrl(url).build()).setBackground(null)
    .setNote('Gắn lúc ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm dd/MM/yyyy'));
  var mc = k.map.code + 1, ma = sh.getRange(r, mc), mrt = ma.getRichTextValue(), txt = String(ma.getValue() || '');
  if (txt && !(mrt && mrt.getLinkUrl())) ma.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(txt).setLinkUrl(url).build());
  return 'LINK_OK';
}

/** Nhật ký: chỉ ghi thêm. rows = [[...9 cột]] */
function tqnQlGhiNhatKy_(rows) {
  if (!rows.length) return;
  var sh = tqnSs_().getSheetByName(TQN_QL.SH_LOG) || tqnOwnSheet_(TQN_QL.SH_LOG, TQN_QL.LOG_HDR);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, TQN_QL.LOG_HDR.length).setValues(rows);
}

/** Trình kích hoạt "khi sửa" (cài đặt). Ghi nhật ký + tính lại dòng bị sửa. */
function tqnQuanLyKhiSua(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet(), name = sh.getName();
  if (name.indexOf(TQN.PREFIX) === 0) return;
  var cfg = tqnConfig_();
  if (cfg.detailSheets.indexOf(name) < 0) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return;
  try {
    var prof = tqnQlMau_(), type = tqnQlLoaiTab_(cfg, prof, name), p = type && prof[type];
    var k = p ? tqnQlKhoi_(sh, cfg, p, false) : null;
    var r1 = e.range.getRow(), r2 = e.range.getLastRow(), c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
    var user = (e.user && e.user.getEmail && e.user.getEmail()) || '';
    if (!user) try { user = Session.getActiveUser().getEmail(); } catch (x) { }
    if (!user) user = '(không lấy được – xem Lịch sử chỉnh sửa ô)';
    var lastC = Math.max(sh.getLastColumn(), c2), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues();
    var codeCol = (k ? k.map.code : tqnDetailMap_(hdr).code) + 1;
    var tenCot = function (c) { var s = []; hdr.forEach(function (r) { var t = String(r[c - 1] || '').replace(/\s+/g, ' ').trim(); if (t && s.indexOf(t) < 0) s.push(t); }); return s.join(' › ') || ('Cột ' + c); };
    var nr = Math.min(r2 - r1 + 1, 50), nc = Math.min(c2 - c1 + 1, 20);
    var moi = sh.getRange(r1, c1, nr, nc).getDisplayValues(), ma = sh.getRange(r1, codeCol, nr, 1).getDisplayValues();
    var now = new Date(), log = [], mot = (nr === 1 && nc === 1);
    for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) {
      var r = r1 + i, c = c1 + j, ghi = '';
      if (k && c === k.link && r >= cfg.firstRow) ghi = tqnQlXuLyLink_(sh, k, r, moi[i][j]);
      if (k && (c === k.giu || c === k.kt)) ghi = 'Gõ vào cột tự tính – đã tính lại';
      log.push([now, user, name, sh.getRange(r, c).getA1Notation(), ma[i][0], tenCot(c),
        mot ? (e.oldValue === undefined ? '' : e.oldValue) : '(sửa nhiều ô)', moi[i][j], ghi]);
    }
    if (r2 - r1 + 1 > nr || c2 - c1 + 1 > nc) log.push([now, user, name, e.range.getA1Notation(), '', '', '', '(vùng lớn – chỉ ghi 50×20 ô đầu)', '']);
    if (k && r2 >= cfg.firstRow && k.lv.some(function (c) { return c >= c1 && c <= c2; })) {
      var rr1 = Math.max(r1, cfg.firstRow), rr2 = Math.min(r2, sh.getLastRow());
      var vv = sh.getRange(rr1, 1, rr2 - rr1 + 1, Math.max(sh.getLastColumn(), k.kt)).getValues();
      vv.forEach(function (row, i) {
        if (!String(row[k.map.subject] || '').trim()) return;
        var picks = k.lv.map(function (c) { return String(row[c - 1] || ''); }), cu = picks.map(function () { return ''; });
        if (mot && e.oldValue !== undefined) k.lv.forEach(function (c, j) { if (c === c1) cu[j] = String(e.oldValue); });
        var nd = [k.map.nt && k.map.nt.plan, k.map.nt && k.map.nt.actual].filter(function (c) { return c != null && c >= 0; });
        var rec = { noiDung: row[k.map.subject], tinhTrang: row[k.map.status], picks: picks, chuyenToi: k.to > 0 ? row[k.to - 1] : '', link: 'x',
          daTrinh: nd.some(function (c) { return tqnIsDate_(row[c]); }) || /da trinh|cho duyet/.test(tqnQlBo_(row[k.map.status])) };
        tqnQlTuDien_(sh, p, k, rr1 + i, row, picks, cu, user, now, log, tqnQlDanhGia_(rec, p));
      });
    }
    if (k && r2 >= cfg.firstRow) {
      var loi = tqnQlTinhDong_(sh, cfg, p, k, r1, r2);
      var kt = sh.getRange(Math.max(r1, cfg.firstRow), k.kt, Math.min(r2, sh.getLastRow()) - Math.max(r1, cfg.firstRow) + 1, 1).getValues();
      kt.forEach(function (x, i) { if (String(x[0]).indexOf('⚠') === 0) log.push([now, user, name, 'Dòng ' + (Math.max(r1, cfg.firstRow) + i), '', 'Kiểm tra', '', x[0], 'CẢNH BÁO']); });
      if (loi) tqnSs_().toast(kt.map(function (x) { return x[0]; }).filter(function (x) { return String(x).indexOf('⚠') === 0; }).slice(0, 3).join('\n'), 'TQN – cần xem lại', 8);
    }
    tqnQlGhiNhatKy_(log);
  } finally { lock.releaseLock(); }
}

/** Tính lại toàn bộ (gọi mỗi giờ + sau khi sửa mẫu phiếu). Trả [{tab, loi}] */
function tqnQlTinhLaiTatCa() {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), out = [];
  cfg.detailSheets.forEach(function (n) {
    var type = tqnQlLoaiTab_(cfg, prof, n), sh = type && tqnSs_().getSheetByName(n);
    if (!sh) return;
    var k = tqnQlKhoi_(sh, cfg, prof[type], false);
    if (k) out.push({ tab: n, loi: tqnQlTinhDong_(sh, cfg, prof[type], k) });
  });
  return out;
}

/** Cài: mẫu phiếu, nhật ký (chỉ chủ file sửa), khối quản lý ở từng sổ, trình kích hoạt. */
function tqnCaiQuanLy() {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), out = [];
  if (prof._loi.length) out.push('⚠ Mẫu phiếu: ' + prof._loi.join('; '));
  var log = tqnSs_().getSheetByName(TQN_QL.SH_LOG) || tqnOwnSheet_(TQN_QL.SH_LOG, TQN_QL.LOG_HDR);
  try {
    var pr = log.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || log.protect();
    pr.setDescription('TQN nhật ký – chỉ hệ thống ghi');
    var me = Session.getEffectiveUser(); pr.addEditor(me);
    pr.removeEditors(pr.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); }));
    if (pr.canDomainEdit()) pr.setDomainEdit(false);
    out.push('✔ Nhật ký ' + TQN_QL.SH_LOG + ': người khác không sửa/xóa được');
  } catch (e) { out.push('✖ Khóa nhật ký: ' + e.message); }
  cfg.detailSheets.forEach(function (n) {
    var type = tqnQlLoaiTab_(cfg, prof, n), sh = type && tqnSs_().getSheetByName(n);
    if (!sh) return;
    var k = tqnQlKhoi_(sh, cfg, prof[type], true);
    tqnQlDinhDang_(sh, cfg, prof[type], k);
    try { tqnQlToDong_(sh, cfg, k); } catch (e) { out.push('✖ Tô màu dòng ' + n + ': ' + e.message); }
    var loi = tqnQlTinhDong_(sh, cfg, prof[type], k);
    out.push('✔ ' + n + ' (' + type + '): khối quản lý từ cột ' + sh.getRange(1, k.start).getA1Notation().replace(/\d+/, '') + ', ' + loi + ' dòng cần xem lại');
  });
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnQuanLyKhiSua') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tqnQuanLyKhiSua').forSpreadsheet(tqnSs_()).onEdit().create();
  out.push('✔ Ghi nhật ký + tự kiểm khi sửa: BẬT');
  tqnAudit_('CAI_QUAN_LY', out.join(' | '), {});
  return out;
}

function tqnMenuQuanLy() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('TQN – Quản lý theo mẫu phiếu',
    'Sẽ làm:\n• Tạo tab ' + TQN_QL.SH_MAU + ' (A/B/C/D đúng nguyên văn phiếu dự án – sửa được khi qua dự án khác)\n' +
    '• Thêm khối "TQN QUẢN LÝ" ở BÊN PHẢI các sổ SD/RFA/RFC/RFI (không đụng cột cũ)\n' +
    '• Tab ' + TQN_QL.SH_LOG + ': ghi mọi lần sửa, chỉ hệ thống ghi\n\nLàm luôn?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  ui.alert('TQN – Quản lý', tqnCaiQuanLy().join('\n'), ui.ButtonSet.OK);
}
function tqnMenuTinhLaiQuanLy() {
  var r = tqnQlTinhLaiTatCa();
  SpreadsheetApp.getUi().alert('TQN – Tính lại', r.map(function (x) { return x.tab + ': ' + x.loi + ' dòng cần xem lại'; }).join('\n') || 'Chưa cài khối quản lý', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ===================== PHÂN QUYỀN THEO EMAIL ===================== */
/**
 * Tab TQN_PHAN_QUYEN: Email | Vai trò | Sửa kết quả phiếu & link (Có/Không) | Ghi chú
 * - Ai được chia sẻ quyền "Người chỉnh sửa" file đều cập nhật được sổ như thường.
 * - Riêng cột kết quả theo phiếu (TVGS/CĐT/RFI) + Link OneDrive: CHỈ email ghi "Có" sửa được (khóa thật của Google, không chỉ cảnh báo).
 * - Chủ file luôn sửa được.
 */
var TQN_PQ = { SH: 'TQN_PHAN_QUYEN', HDR: ['Email', 'Vai trò', 'Sửa kết quả phiếu & link (Có/Không)', 'Ghi chú', 'Tên trong sổ (cột Người thực hiện)', 'Nhận email (Có/Không)'], MO_TA: 'TQN – cột kết quả phiếu: chỉ người được giao' };

/** Lõi: danh sách email được sửa cột kết quả từ bảng phân quyền. */
function tqnPqDanhSach_(rows) {
  var ok = [], loi = [];
  (rows || []).forEach(function (r, i) {
    var em = String(r[0] || '').trim().toLowerCase(), co = tqnQlBo_(r[2]);
    if (!em) return;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) { loi.push('Dòng ' + (i + 2) + ': email sai "' + r[0] + '"'); return; }
    if (co === 'co' && ok.indexOf(em) < 0) ok.push(em);
  });
  return { ok: ok, loi: loi };
}

function tqnEnsurePhanQuyen_() {
  var sh = tqnSs_().getSheetByName(TQN_PQ.SH);
  if (sh) return sh;
  sh = tqnOwnSheet_(TQN_PQ.SH, TQN_PQ.HDR);
  var me = ''; try { me = Session.getEffectiveUser().getEmail(); } catch (e) { }
  sh.getRange(2, 1, 1, 6).setValues([[me, 'Chủ file / Tổng hợp', 'Có', 'Chủ file luôn sửa được', '', 'Có']]);
  try {
    sh.getRange(2, 3, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Có', 'Không'], true).build());
    sh.getRange(1, 1, 1, 4).setNote('Bước 1: chia sẻ file cho email (Người chỉnh sửa) như bình thường.\nBước 2: ghi email vào đây, "Có" = được chọn kết quả TVGS/CĐT/RFI và dán link phiếu.\nBước 3: menu TQN Register → Áp dụng phân quyền.');
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 180); sh.setColumnWidth(3, 230); sh.setColumnWidth(4, 260);
  } catch (e) { }
  return sh;
}

/** Áp khóa cột kết quả + link theo bảng phân quyền. Trả dòng báo cáo. */
function tqnApDungPhanQuyen() {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), sh = tqnEnsurePhanQuyen_(), n = sh.getLastRow() - 1;
  var ds = tqnPqDanhSach_(n > 0 ? sh.getRange(2, 1, n, 4).getValues() : []), out = [];
  if (ds.loi.length) out.push('⚠ ' + ds.loi.join('; '));
  var me = Session.getEffectiveUser().getEmail();
  cfg.detailSheets.forEach(function (name) {
    var type = tqnQlLoaiTab_(cfg, prof, name), s = type && tqnSs_().getSheetByName(name);
    if (!s) return;
    var k = tqnQlKhoi_(s, cfg, prof[type], false);
    if (!k) return;
    s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { if (p.getDescription() === TQN_PQ.MO_TA) p.remove(); });
    var cols = k.lv.concat([k.to, k.link]).filter(function (c) { return c > 0; }), rows = Math.max(s.getMaxRows() - cfg.firstRow + 1, 1);
    cols.forEach(function (c) {
      var pr = s.getRange(cfg.firstRow, c, rows, 1).protect().setDescription(TQN_PQ.MO_TA);
      pr.addEditor(me);
      pr.removeEditors(pr.getEditors().filter(function (u) { var e = u.getEmail().toLowerCase(); return e !== me.toLowerCase() && ds.ok.indexOf(e) < 0; }));
      var them = ds.ok.filter(function (e) { return e !== me.toLowerCase(); });
      if (them.length) pr.addEditors(them);
      if (pr.canDomainEdit()) pr.setDomainEdit(false);
    });
    out.push('✔ ' + name + ': ' + cols.length + ' cột kết quả/link – ' + (ds.ok.length || 1) + ' người được sửa');
  });
  tqnAudit_('PHAN_QUYEN', out.join(' | '), { duoc_sua: ds.ok });
  return out;
}
function tqnMenuPhanQuyen() {
  tqnEnsurePhanQuyen_();
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('TQN – Phân quyền theo email', 'Đã có tab ' + TQN_PQ.SH + '. Ghi email + "Có" cho người được chọn kết quả phiếu và dán link.\n\nÁp dụng ngay theo bảng hiện tại?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('TQN – Phân quyền', tqnApDungPhanQuyen().join('\n'), ui.ButtonSet.OK);
}

/* ===================== AI ĐỌC MẪU PHIẾU DỰ ÁN MỚI ===================== */
/**
 * Anh đưa phiếu TRẮNG của dự án mới (PDF / ảnh scan / Google Doc trên Drive) → AI đọc ra các cấp duyệt
 * và các ô lựa chọn → ghi tab TQN_MAU_PHIEU_DE_XUAT. AI CHỈ ĐỀ XUẤT: anh xem, sửa, bấm "Dùng đề xuất" mới thay vào mẫu phiếu.
 */
var TQN_QL_DX = 'TQN_MAU_PHIEU_DE_XUAT';

function tqnAiMauPhieuPrompt_(type) {
  return 'Đây là MẪU PHIẾU TRỐNG loại "' + type + '" của một dự án xây dựng ở Việt Nam.\n' +
    'Tìm các KHỐI PHẢN HỒI/PHÊ DUYỆT (ví dụ TVGS, TVQLDA, TVTK, PMC, Chủ đầu tư/BQLXD) theo đúng thứ tự trên phiếu, ' +
    'và trong mỗi khối các Ô LỰA CHỌN (ô tick/checkbox, ví dụ A/B/C/D hoặc "Xem trả lời bên dưới").\n' +
    'Chép NGUYÊN VĂN chữ của từng ô lựa chọn, không dịch, không tóm tắt. Nếu ô không có mã chữ thì đánh số 1,2,3.\n' +
    'Gán kết quả cho mỗi lựa chọn bằng một trong các mã:\n' +
    Object.keys(TQN_QL.KQ).map(function (k) { return '- ' + k + ': ' + TQN_QL.KQ[k]; }).join('\n') + '\n' +
    'Quy tắc: cấp chưa phải cuối mà chấp thuận → TIEP; cấp cuối chấp thuận → DONG (có điều kiện/ý kiến → DONG_YK); ' +
    'yêu cầu sửa trình lại → TRINH_LAI; từ chối → TU_CHOI; yêu cầu thêm thông tin → BO_SUNG; chuyển tiếp bên khác → CHUYEN.\n' +
    'Chỉ trả về JSON, không giải thích: {"cap":[{"ten":"TVGS","thu_tu":1,"lua_chon":[{"ma":"A","nguyen_van":"…","ket_qua":"TIEP"}]}],"ghi_chu":"điều chưa chắc"}';
}

/** Lõi: đổi câu trả lời AI thành các dòng mẫu phiếu + lỗi (dùng lại bộ kiểm của tqnQlDocMau_). */
function tqnAiMauPhieuParse_(text, type) {
  var s = String(text || ''), i = s.indexOf('{'), j = s.lastIndexOf('}');
  if (i < 0 || j <= i) return { rows: [], loi: ['AI không trả JSON'], ghiChu: '' };
  var o; try { o = JSON.parse(s.slice(i, j + 1)); } catch (e) { return { rows: [], loi: ['JSON lỗi: ' + e.message], ghiChu: '' }; }
  var rows = [];
  (o.cap || []).forEach(function (c, ci) {
    (c.lua_chon || []).forEach(function (l, li) {
      rows.push([type, parseInt(c.thu_tu, 10) || ci + 1, String(c.ten || ('Cấp ' + (ci + 1))).trim(), String(l.ma || (li + 1)).trim(),
        String(l.nguyen_van || '').trim(), String(l.ket_qua || '').trim().toUpperCase(), li === 0 ? 'AI đề xuất – cần kiểm' : '']);
    });
  });
  var chk = tqnQlDocMau_(rows);
  if (!rows.length) chk._loi.push('Không tìm thấy ô lựa chọn nào');
  return { rows: rows, loi: chk._loi, ghiChu: String(o.ghi_chu || '') };
}

/** Lấy file từ link/ID Drive → khối nội dung gửi AI (PDF hoặc ảnh). */
function tqnAiKhoiFile_(link) {
  var m = String(link).match(/[-\w]{25,}/);
  if (!m) throw new Error('Không thấy ID file trong link: ' + link);
  var f = DriveApp.getFileById(m[0]), mt = f.getMimeType(), blob;
  if (mt === 'application/pdf' || mt === 'application/vnd.google-apps.document') blob = f.getAs('application/pdf');
  else if (/^image\/(png|jpeg|gif|webp)$/.test(mt)) blob = f.getBlob();
  else throw new Error('Chỉ đọc được PDF, ảnh scan hoặc Google Doc (file này là ' + mt + '). File Word: mở bằng Google Docs rồi gửi link Google Doc.');
  var data = Utilities.base64Encode(blob.getBytes()), isPdf = blob.getContentType() === 'application/pdf';
  return { ten: f.getName(), khoi: isPdf ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: data } }
    : { type: 'image', source: { type: 'base64', media_type: blob.getContentType(), data: data } } };
}

function tqnAiDocMauPhieu(link, type) {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error(why);
  type = String(type || '').trim().toUpperCase();
  if (!type) throw new Error('Thiếu loại hồ sơ (vd RFA, RFI, SDS)');
  var f = tqnAiKhoiFile_(link);
  var res = tqnAiMauPhieuParse_(tqnAiCall_(cfg, [f.khoi, { type: 'text', text: tqnAiMauPhieuPrompt_(type) }], 3000), type);
  var sh = tqnSs_().getSheetByName(TQN_QL_DX) || tqnOwnSheet_(TQN_QL_DX, TQN_QL.MAU_HDR);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  if (res.rows.length) sh.getRange(2, 1, res.rows.length, 7).setValues(res.rows);
  sh.getRange(1, 1).setNote('AI đọc "' + f.ten + '" lúc ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm dd/MM/yyyy') +
    (res.ghiChu ? '\nAI chưa chắc: ' + res.ghiChu : '') + (res.loi.length ? '\nLỗi: ' + res.loi.join('; ') : ''));
  tqnAudit_('AI_MAU_PHIEU', f.ten + ' → ' + res.rows.length + ' dòng', { loai: type, loi: res.loi });
  return res;
}

/** Thay các dòng của loại hồ sơ trong TQN_MAU_PHIEU bằng bảng đề xuất (sau khi anh đã xem). */
function tqnDungDeXuatMauPhieu() {
  var dx = tqnSs_().getSheetByName(TQN_QL_DX);
  if (!dx || dx.getLastRow() < 2) throw new Error('Chưa có đề xuất – chạy "AI đọc mẫu phiếu" trước');
  var moi = dx.getRange(2, 1, dx.getLastRow() - 1, 7).getValues().filter(function (r) { return String(r[0]).trim(); });
  var chk = tqnQlDocMau_(moi);
  if (chk._loi.length) throw new Error('Đề xuất còn lỗi, sửa trước: ' + chk._loi.join('; '));
  var loai = {}; moi.forEach(function (r) { loai[String(r[0]).trim().toUpperCase()] = 1; });
  var sh = tqnEnsureMauPhieu_(), n = sh.getLastRow() - 1, cu = n > 0 ? sh.getRange(2, 1, n, 7).getValues() : [];
  var giu = cu.filter(function (r) { return !loai[String(r[0]).trim().toUpperCase()]; }), all = giu.concat(moi);
  if (n > 0) sh.getRange(2, 1, n, 7).clearContent();
  sh.getRange(2, 1, all.length, 7).setValues(all);
  tqnAudit_('DUNG_MAU_PHIEU', Object.keys(loai).join(', ') + ': ' + moi.length + ' dòng (thay ' + (cu.length - giu.length) + ')', {});
  return Object.keys(loai);
}

function tqnMenuAiMauPhieu() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('AI đọc mẫu phiếu dự án mới', 'Dán link Drive của PHIẾU TRẮNG (PDF / ảnh scan / Google Doc):', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt('Loại hồ sơ', 'Phiếu này cho loại nào? (vd RFA, RFC, RFI, SDS)', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  try {
    var r = tqnAiDocMauPhieu(a.getResponseText(), b.getResponseText());
    var ok = ui.alert('AI đã đọc xong', r.rows.length + ' lựa chọn – xem ở tab ' + TQN_QL_DX + '.' + (r.ghiChu ? '\nAI chưa chắc: ' + r.ghiChu : '') +
      (r.loi.length ? '\n⚠ ' + r.loi.join('; ') : '') + '\n\nĐã kiểm, dùng luôn thay mẫu cũ của loại này?', ui.ButtonSet.YES_NO);
    if (ok === ui.Button.YES) { tqnDungDeXuatMauPhieu(); ui.alert('Đã dùng mẫu mới. Bấm "Quản lý theo mẫu phiếu: cài / cập nhật" để cập nhật dropdown.'); }
  } catch (e) { ui.alert('Lỗi', e.message, ui.ButtonSet.OK); }
}


/* ===================== TỰ ĐIỀN KHI CHỌN KẾT QUẢ (dùng chung sheet + web) ===================== */
/**
 * Chọn kết quả cấp nào (khác giá trị cũ) → ô "Thực tế" của khối cấp đó trống thì điền hôm nay.
 * Tình trạng đổi theo phiếu (Đã duyệt / Không duyệt / Đã trình duyệt). Ghi nhật ký vào mảng log.
 */
function tqnQlTuDien_(sh, p, k, r, row, picks, cu, user, now, log, d) {
  var m = k.map, homNay = new Date(now.getFullYear(), now.getMonth(), now.getDate()), out = { stMoi: '', ngay: [] };
  p.levels.forEach(function (L, i) {
    var v = String(picks[i] || ''); if (!v || v === String(cu[i] || '')) return;
    var n = tqnQlBo_(L.name), b = /tvgs|giam sat|tvqlda/.test(n) ? m.tvgs : (/cdt|chu dau tu/.test(n) ? m.cdt : (/tvtk|thiet ke/.test(n) ? m.tvtk : (i === 0 ? m.tvgs : null)));
    if (b && b.actual >= 0 && !tqnIsDate_(row[b.actual]) && !String(row[b.actual] || '').trim()) {
      sh.getRange(r, b.actual + 1).setValue(homNay); row[b.actual] = homNay; out.ngay.push(L.name);
      log.push([now, user, sh.getName(), sh.getRange(r, b.actual + 1).getA1Notation(), row[m.code], L.name + ' › Thực tế', '', 'hôm nay', 'TỰ ĐIỀN']);
    }
  });
  var st = tqnWebTinhTrang_(d, picks.some(function (x) { return String(x || '').trim(); }));
  if (st && m.status >= 0 && st !== String(row[m.status] || '')) {
    sh.getRange(r, m.status + 1).setValue(st);
    log.push([now, user, sh.getName(), sh.getRange(r, m.status + 1).getA1Notation(), row[m.code], 'Tình trạng', row[m.status], st, 'TỰ ĐỔI']);
    row[m.status] = st; out.stMoi = st;
    if (st === 'Không duyệt') try { tqnEmailTraVe_(sh.getName(), r); } catch (e) { }
  }
  return out;
}

/* ===================== TÔ MÀU DÒNG THEO BÊN ĐANG GIỮ ===================== */
var TQN_QL_TO = [['↩', '#FDECEA'], ['⏳', '#FFF6DB'], ['✔', '#F2F4F7']];
function tqnQlCongThucTo_(k, cfg) {
  var c = tqnQlCot_(k.giu);
  return TQN_QL_TO.map(function (x) { return ['=LEFT($' + c + cfg.firstRow + ';1)="' + x[0] + '"', x[1]]; });
}
function tqnQlCot_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
/** Thêm (hoặc làm lại) 3 quy tắc tô nền cả dòng theo cột Ai đang giữ. Quy tắc cũ của tab được ưu tiên trước. */
function tqnQlToDong_(sh, cfg, k) {
  var ct = tqnQlCongThucTo_(k, cfg), mine = ct.map(function (x) { return x[0].replace(/;/g, ','); }).concat(ct.map(function (x) { return x[0]; }));
  var rules = sh.getConditionalFormatRules().filter(function (rl) {
    var b = rl.getBooleanCondition(), v = String(b && b.getCriteriaValues && b.getCriteriaValues()[0] || '');
    return mine.indexOf(v) < 0;
  });
  var rg = sh.getRange(cfg.firstRow, 1, Math.max(sh.getMaxRows() - cfg.firstRow + 1, 1), Math.max(k.start - 2, 1));
  ct.forEach(function (x) { rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x[0].replace(/;/g, ',')).setBackground(x[1]).setRanges([rg]).build()); });
  sh.setConditionalFormatRules(rules);
  return ct.length;
}

/* ===================== TRÌNH LẠI REV MỚI ===================== */
/** Lõi: mã rev mới từ mã cũ – thay ngày đầu mã bằng hôm nay, tăng rev cuối mã (nếu có). */
function tqnQlMaRevMoi_(ma, yymmdd, revMoi) {
  var s = String(ma || '').trim().replace(/^(\d{6})\s*-\s*/, yymmdd + '-');
  if (/rev\s*\d+$/i.test(s)) s = s.replace(/(rev\s*)(\d+)$/i, function (_, a, n) { return a + ('0' + revMoi).slice(-Math.max(2, n.length)); });
  return s;
}
function tqnQlRevMoi_(rev) { var n = parseInt(String(rev || '').replace(/\D/g, ''), 10); return isNaN(n) ? 1 : n + 1; }

/** Đứng ở dòng hồ sơ bị trả về → tạo dòng rev mới ở dòng trống kế tiếp cuối sổ. */
function tqnTrinhLai(sh, r) {
  var cfg = tqnCodeCfg_(), name = sh.getName();
  if (cfg.detailSheets.indexOf(name) < 0) throw new Error('Tab "' + name + '" không phải sổ hồ sơ');
  if (r < cfg.firstRow) throw new Error('Chọn một dòng hồ sơ (từ dòng ' + cfg.firstRow + ')');
  var lastC = Math.max(sh.getLastColumn(), 1), hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues(), m = tqnDetailMap_(hdr);
  var row = sh.getRange(r, 1, 1, lastC).getValues()[0], ma = String(row[m.code] || '').trim();
  if (!ma || !String(row[m.subject] || '').trim()) throw new Error('Dòng ' + r + ' chưa có mã / nội dung');
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(15000)) throw new Error('Đang có người lưu, thử lại');
  try {
    var last = sh.getLastRow(), all = sh.getRange(1, 1, last, lastC).getValues(), cStt = tqnFindCol_(hdr, ['STT']);
    var dest = -1, maxStt = 0;
    for (var i = cfg.firstRow - 1; i < all.length; i++) {
      if (String(all[i][m.subject] || '').trim()) { dest = -1; var st = parseFloat(all[i][cStt]); if (!isNaN(st)) maxStt = Math.max(maxStt, st); }
      else if (dest < 0 && !/^[IVX]+$/.test(String(all[i][0]).trim()) && !String(all[i][m.code] || '').trim()) dest = i + 1;
      if (String(all[i][m.code] || '').trim() === ma) dest = -1;
    }
    if (dest < 0 || dest <= r) dest = last + 1;
    var rev = tqnQlRevMoi_(m.rev >= 0 ? row[m.rev] : ''), yy = Utilities.formatDate(new Date(), cfg.timezone || 'Asia/Ho_Chi_Minh', 'yyMMdd');
    var maMoi = tqnQlMaRevMoi_(ma, yy, rev);
    if (cStt >= 0 && !String(sh.getRange(dest, cStt + 1).getValue()).trim()) sh.getRange(dest, cStt + 1).setValue(maxStt + 1);
    sh.getRange(dest, m.code + 1).setValue(maMoi);
    sh.getRange(dest, m.subject + 1).setValue(row[m.subject]);
    if (m.owner >= 0) sh.getRange(dest, m.owner + 1).setValue(row[m.owner]);
    if (m.rev >= 0) sh.getRange(dest, m.rev + 1).setValue(('0' + rev).slice(-2));
    if (m.status >= 0) sh.getRange(dest, m.status + 1).setValue('Đang thực hiện');
    sh.getRange(r, m.code + 1).setNote('Đã thay thế bởi ' + maMoi + ' (dòng ' + dest + ') lúc ' + Utilities.formatDate(new Date(), cfg.timezone || 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM/yyyy'));
    var u = ''; try { u = Session.getActiveUser().getEmail(); } catch (e) { }
    tqnQlGhiNhatKy_([[new Date(), u || '(không rõ)', name, 'Dòng ' + dest, maMoi, 'Trình lại', ma, maMoi, 'TRÌNH LẠI']]);
    return { row: dest, ma: maMoi };
  } finally { lock.releaseLock(); }
}
function tqnMenuTrinhLai() {
  var ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet(), r = sh.getActiveCell().getRow();
  try {
    var x = tqnTrinhLai(sh, r);
    sh.getRange(x.row, 2).activate();
    tqnSs_().toast('Dòng ' + x.row + ': ' + x.ma, 'TQN – đã tạo rev mới', 6);
  } catch (e) { ui.alert('Trình lại rev mới', e.message, ui.ButtonSet.OK); }
}

/* ===================== BỘ LỌC SẴN CHO TỪNG NGƯỜI / TÌNH HUỐNG (chế độ lọc của Google) ===================== */
/** Lõi: các yêu cầu tạo chế độ lọc. owners = danh sách tên người thực hiện. */
function tqnQlBoLocReq_(sheetId, k, owners, firstRow, lastRow, lastCol) {
  var range = { sheetId: sheetId, startRowIndex: firstRow - 2, endRowIndex: lastRow, startColumnIndex: 0, endColumnIndex: lastCol };
  var v = function (title, col, type, val) {
    var c = {}; c[col - 1] = { condition: { type: type, values: [{ userEnteredValue: val }] } };
    return { addFilterView: { filter: { title: 'TQN · ' + title, range: range, criteria: c } } };
  };
  var out = [v('Nhà thầu đang giữ', k.giu, 'TEXT_STARTS_WITH', '↩'), v('Đang chờ duyệt', k.giu, 'TEXT_STARTS_WITH', '⏳'),
    v('Cần sửa ⚠', k.kt, 'TEXT_STARTS_WITH', '⚠'), v('Đã đóng', k.giu, 'TEXT_STARTS_WITH', '✔')];
  owners.forEach(function (o) { out.push(v('Việc của ' + o, k.map.owner + 1, 'TEXT_EQ', o)); });
  return out;
}
function tqnQlTaoBoLoc() {
  if (typeof Sheets === 'undefined') throw new Error('Chưa bật dịch vụ "Google Sheets API" trong Apps Script (Dịch vụ → +)');
  var cfg = tqnConfig_(), prof = tqnQlMau_(), ss = tqnSs_(), id = ss.getId(), out = [];
  var cu = (Sheets.Spreadsheets.get(id, { fields: 'sheets(properties(sheetId),filterViews(filterViewId,title))' }).sheets || []);
  var xoa = []; cu.forEach(function (s) { (s.filterViews || []).forEach(function (f) { if (String(f.title).indexOf('TQN · ') === 0) xoa.push({ deleteFilterView: { filterId: f.filterViewId } }); }); });
  var them = [];
  cfg.detailSheets.forEach(function (n) {
    var type = tqnQlLoaiTab_(cfg, prof, n), sh = type && ss.getSheetByName(n);
    if (!sh) return;
    var k = tqnQlKhoi_(sh, cfg, prof[type], false); if (!k) return;
    var owners = [];
    if (k.map.owner >= 0 && sh.getLastRow() >= cfg.firstRow) sh.getRange(cfg.firstRow, k.map.owner + 1, sh.getLastRow() - cfg.firstRow + 1, 1).getValues()
      .forEach(function (x) { var t = String(x[0] || '').trim(); if (t && owners.indexOf(t) < 0) owners.push(t); });
    var req = tqnQlBoLocReq_(sh.getSheetId(), k, owners.slice(0, 15), cfg.firstRow, sh.getMaxRows(), k.kt);
    them = them.concat(req); out.push('✔ ' + n + ': ' + req.length + ' chế độ lọc');
  });
  if (xoa.length || them.length) Sheets.Spreadsheets.batchUpdate({ requests: xoa.concat(them) }, id);
  return out;
}
function tqnMenuBoLoc() {
  var ui = SpreadsheetApp.getUi();
  try { ui.alert('TQN – bộ lọc sẵn', tqnQlTaoBoLoc().join('\n') + '\n\nMở: Dữ liệu → Chế độ lọc → chọn "TQN · …". Mỗi người lọc riêng, không làm xáo bảng của người khác.', ui.ButtonSet.OK); }
  catch (e) { ui.alert('TQN – bộ lọc', e.message, ui.ButtonSet.OK); }
}


/* ===================== EMAIL NHẮC (bị trả về · 7h sáng · thứ Hai cho sếp) ===================== */
/** Lõi: bảng phân quyền → { ten(chuẩn hóa): email, sep: [email], tatCa: [email] } */
function tqnEmailDanhBa_(rows) {
  var o = { ten: {}, sep: [], tatCa: [] };
  (rows || []).forEach(function (r) {
    var em = String(r[0] || '').trim().toLowerCase(), vt = tqnQlBo_(r[1]), ten = tqnQlBo_(r[4]), nhan = tqnQlBo_(r[5]);
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em) || nhan === 'khong') return;
    if (o.tatCa.indexOf(em) < 0) o.tatCa.push(em);
    if (ten) ten.split(/\s*[,;]\s*/).forEach(function (t) { if (t) o.ten[t] = em; });
    if (/sep|giam doc|chi huy|chu nhiem|quan ly/.test(vt)) o.sep.push(em);
  });
  return o;
}
/** Lõi: gom việc theo email người phụ trách. rows = tqnWebDuLieu().rows */
function tqnEmailGom_(rows, db) {
  var g = {};
  rows.forEach(function (r) {
    var n = r.giu.charAt(0), em = db.ten[tqnQlBo_(r.nguoi)];
    if (!em) return;
    if (n === '↩' || r.kt.indexOf('⚠') === 0) (g[em] = g[em] || []).push(r);
  });
  return g;
}
function tqnEmailLink_(r) {
  var ss = tqnSs_(), sh = ss.getSheetByName(r.tab);
  return ss.getUrl() + (sh ? '#gid=' + sh.getSheetId() + '&range=B' + r.row : '');
}
function tqnEmailDong_(r) {
  return '<tr><td><a href="' + tqnEmailLink_(r) + '">' + r.ma + '</a></td><td>' + String(r.nd).slice(0, 80) + '</td><td>' + r.giu + '</td><td style="color:#B42318">' + (r.kt.indexOf('⚠') === 0 ? r.kt : '') + '</td></tr>';
}
function tqnEmailBang_(ds) {
  return '<table border="1" cellpadding="5" style="border-collapse:collapse;font:13px Arial"><tr style="background:#1B3F7A;color:#fff"><th>Mã</th><th>Nội dung</th><th>Đang giữ</th><th>Cần sửa</th></tr>' + ds.map(tqnEmailDong_).join('') + '</table>';
}
function tqnEmailDb_() {
  var sh = tqnEnsurePhanQuyen_(), n = sh.getLastRow() - 1;
  return tqnEmailDanhBa_(n > 0 ? sh.getRange(2, 1, n, 6).getValues() : []);
}
function tqnEmailGui_(to, sub, html) {
  if (!to) return 0;
  MailApp.sendEmail({ to: to, subject: '[TQN Register] ' + sub, htmlBody: html + '<p style="color:#667085;font:12px Arial">Email tự động từ TQN Register – ' + tqnSs_().getName() + '</p>', name: 'TQN Register' });
  return 1;
}

/** Gọi khi một dòng vừa bị trả về (C/D) – gửi ngay cho người phụ trách. */
function tqnEmailTraVe_(tab, rowNum) {
  if (!tqnEmailBat_()) return 0;
  var db = tqnEmailDb_(), d = tqnWebDuLieu(), r = d.rows.filter(function (x) { return x.tab === tab && x.row === rowNum; })[0];
  if (!r) return 0;
  var em = db.ten[tqnQlBo_(r.nguoi)];
  return tqnEmailGui_(em, 'Hồ sơ bị trả về: ' + r.ma, '<p>Hồ sơ <b>' + r.ma + '</b> – ' + r.nd + '</p><p>Kết quả: ' + r.picks.filter(String).join(' / ') + '</p><p>' + r.giu + '</p><p><a href="' + tqnEmailLink_(r) + '">Mở đúng dòng trong Register</a></p>');
}

/** 7h sáng: mỗi người nhận 1 email các hồ sơ mình đang giữ / cần sửa. */
function tqnEmailSang() {
  if (!tqnEmailBat_()) return 0;
  var db = tqnEmailDb_(), g = tqnEmailGom_(tqnWebDuLieu().rows, db), n = 0;
  Object.keys(g).forEach(function (em) { n += tqnEmailGui_(em, 'Việc hôm nay: ' + g[em].length + ' hồ sơ', '<p>Chào bạn, đây là các hồ sơ đang chờ bạn xử lý:</p>' + tqnEmailBang_(g[em])); });
  tqnAudit_('EMAIL_SANG', n + ' email', {});
  return n;
}

/** Thứ Hai: email tóm tắt cho sếp (Vai trò có chữ Sếp / Giám đốc / Chỉ huy / Quản lý). */
function tqnEmailSep() {
  if (!tqnEmailBat_()) return 0;
  var db = tqnEmailDb_(), rows = tqnWebDuLieu().rows, dem = { DONG: 0, NT: 0, CHO: 0 }, can = [];
  rows.forEach(function (r) {
    var c = r.giu.charAt(0); if (c === '✔') dem.DONG++; else if (c === '↩') dem.NT++; else if (c === '⏳') dem.CHO++;
    if (c === '↩' || c === '⏳' || r.kt.indexOf('⚠') === 0) can.push(r);
  });
  var html = '<h3 style="font:bold 16px Arial;color:#1B3F7A">Tình hình hồ sơ – ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') + '</h3>' +
    '<p style="font:14px Arial">Tổng ' + rows.length + ' · ✔ Đã đóng ' + dem.DONG + ' · ⏳ Đang chờ duyệt ' + dem.CHO + ' · ↩ Nhà thầu đang giữ ' + dem.NT + '</p>' +
    (can.length ? '<p><b>Cần chú ý (' + can.length + '):</b></p>' + tqnEmailBang_(can.slice(0, 15)) : '<p>Không có hồ sơ cần chú ý.</p>') +
    '<p><a href="' + tqnSs_().getUrl() + '">Mở Register</a></p>';
  var n = tqnEmailGui_(db.sep.join(','), 'Tóm tắt tuần', html);
  tqnAudit_('EMAIL_SEP', db.sep.join(','), {});
  return n;
}

function tqnEmailBat_() { return PropertiesService.getDocumentProperties().getProperty('tqnEmail') === 'on'; }
function tqnBatEmail() {
  PropertiesService.getDocumentProperties().setProperty('tqnEmail', 'on');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (/^tqnEmail(Sang|Sep)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tqnEmailSang').timeBased().everyDays(1).atHour(7).nearMinute(5).create();
  ScriptApp.newTrigger('tqnEmailSep').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(20).create();
  return ['✔ Email khi hồ sơ bị trả về: ngay lúc chọn C/D', '✔ Email việc hằng ngày: ~07:05', '✔ Email tóm tắt cho sếp: thứ Hai ~07:20'];
}
function tqnTatEmail() {
  PropertiesService.getDocumentProperties().setProperty('tqnEmail', 'off');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (/^tqnEmail(Sang|Sep)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t); });
}
function tqnMenuEmail() {
  tqnEnsurePhanQuyen_();
  var ui = SpreadsheetApp.getUi(), db = tqnEmailDb_();
  var msg = 'Danh bạ ở tab ' + TQN_PQ.SH + ': ' + db.tatCa.length + ' email, ' + Object.keys(db.ten).length + ' tên khớp cột Người thực hiện, ' + db.sep.length + ' người nhận tóm tắt (Vai trò có "Sếp/Chỉ huy/Quản lý").\n\n' +
    (tqnEmailBat_() ? 'Đang BẬT. Tắt email?' : 'Đang TẮT. Bật email nhắc?');
  if (ui.alert('TQN – Email nhắc', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (tqnEmailBat_()) { tqnTatEmail(); ui.alert('Đã tắt email nhắc.'); }
  else ui.alert('TQN – Email nhắc', tqnBatEmail().join('\n'), ui.ButtonSet.OK);
}


/* ===================== ẨN / HIỆN KHỐI TƯ VẤN (dự án không có tư vấn) ===================== */
/** Lõi: các khoảng cột (1-based) của khối Tư vấn giám sát + Tư vấn thiết kế, gồm cả cột trống cuối khối. */
function tqnQlCotTuVan_(m) {
  return [m.tvgs, m.tvtk].filter(function (b) { return b; }).map(function (b) { return { from: b.start + 1, n: b.end - b.start + 1 }; });
}
function tqnAnHienTuVan(an) {
  var cfg = tqnConfig_(), out = [];
  cfg.detailSheets.forEach(function (name) {
    var sh = tqnSs_().getSheetByName(name); if (!sh) return;
    var hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, Math.max(sh.getLastColumn(), 1)).getValues();
    var mp = tqnDetailMap_(hdr), ds = tqnQlCotTuVan_(mp);
    ds.forEach(function (k) { if (an) sh.hideColumns(k.from, k.n); else sh.showColumns(k.from, k.n); });
    // Dự án không tư vấn: khối CĐT là nơi duyệt chính → luôn hiện (sổ gốc đang ẩn sẵn P–T)
    if (an && mp.cdt) sh.showColumns(mp.cdt.start + 1, mp.cdt.end - mp.cdt.start + 1);
    if (ds.length) out.push(name + ': ' + ds.map(function (k) { return k.n + ' cột'; }).join(' + '));
  });
  PropertiesService.getDocumentProperties().setProperty('tqnAnTuVan', an ? 'on' : 'off');
  tqnAudit_(an ? 'AN_TU_VAN' : 'HIEN_TU_VAN', out.join(' | '), {});
  return out;
}
function tqnMenuAnTuVan() { var ui = SpreadsheetApp.getUi(); ui.alert('Đã ẩn khối Tư vấn (dữ liệu vẫn giữ, chỉ ẩn cột)', tqnAnHienTuVan(true).join('\n') || 'Không thấy khối tư vấn', ui.ButtonSet.OK); }
function tqnMenuHienTuVan() { var ui = SpreadsheetApp.getUi(); ui.alert('Đã hiện lại khối Tư vấn', tqnAnHienTuVan(false).join('\n'), ui.ButtonSet.OK); }

/* ===== Sổ gọn: chữ 10, dòng vừa nội dung (giống bản xem trước) – chỉ đổi định dạng, không đổi dữ liệu ===== */
function tqnGonMotSo_(sh, cfg) {
  var last = sh.getLastRow(), lc = sh.getLastColumn();
  sh.getRange(1, 1, 1, lc).setFontSize(16); sh.setRowHeight(1, 30);
  sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lc).setFontSize(10);
  if (last >= cfg.firstRow) {
    var n = last - cfg.firstRow + 1;
    sh.getRange(cfg.firstRow, 1, n, lc).setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sh.setRowHeights(cfg.firstRow, n, 24);
    sh.autoResizeRows(cfg.firstRow, n);       // dòng 2 hàng chữ tự cao lên vừa đủ
  }
  return sh.getName() + ': ' + Math.max(last - cfg.firstRow + 1, 0) + ' dòng';
}
function tqnGonSo(ten) {
  var cfg = tqnConfig_(), ds = ten ? [ten] : cfg.detailSheets, out = [];
  ds.forEach(function (x) { var sh = tqnSs_().getSheetByName(x); if (sh) out.push(tqnGonMotSo_(sh, cfg)); });
  tqnAudit_('GON_SO', out.join(' | '), {});
  return out;
}
function tqnGonSoRFA() { return tqnGonSo('RFA'); }
function tqnGonTatCaSo() { return tqnGonSo(); }

/* ===== Nối hộp Quản lý hồ sơ với khối TQN QUẢN LÝ ===== */
function tqnQlTheoMa_() {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), out = {};
  cfg.detailSheets.forEach(function (n) {
    var type = tqnQlLoaiTab_(cfg, prof, n), sh = type && tqnSs_().getSheetByName(n);
    if (!sh || sh.getLastRow() < cfg.firstRow) return;
    var k = tqnQlKhoi_(sh, cfg, prof[type], false); if (!k || k.giu < 0 || k.map.code == null || k.map.code < 0) return;
    var nr = sh.getLastRow() - cfg.firstRow + 1;
    var ma = sh.getRange(cfg.firstRow, k.map.code + 1, nr, 1).getValues(), gi = sh.getRange(cfg.firstRow, k.giu, nr, 1).getValues();
    for (var i = 0; i < nr; i++) if (ma[i][0] && gi[i][0]) out[tqnKeyOf_(ma[i][0])] = { giu: String(gi[i][0]), tab: n, row: cfg.firstRow + i };
  });
  return out;
}
function tqnQlGanVaoManager_(rows) {
  var ql = tqnQlTheoMa_(), khongTv = PropertiesService.getDocumentProperties().getProperty('tqnAnTuVan') === 'on';
  rows.forEach(function (x) {
    if (khongTv && /^(TVGS|TVTK)$/.test(String(x.bic))) x.bic = 'CĐT';   // dự án không tư vấn: bên duyệt là CĐT
    var q = ql[tqnKeyOf_(x.so)]; if (!q) return;
    var g = q.giu; x.ql = g; x.bic = g.replace(/^[⏳↩✔⤳✎]\s*/, '');
    if (/^⏳/.test(g)) { if (x.buoc === 1 || x.buoc === 5) x.buoc = 2; x.sla = /· trễ/.test(g) ? 'QUA_HAN' : (x.sla === 'QUA_HAN' ? 'CON_HAN' : x.sla); x.next = 'NHAC'; }
    else if (/^↩/.test(g)) { x.buoc = 4; x.sla = 'KHONG_HAN'; x.next = 'TRINH_LAI'; }
    else if (/^[✔⤳]/.test(g)) { x.buoc = 5; x.sla = 'XONG'; x.next = ''; }
  });
}
