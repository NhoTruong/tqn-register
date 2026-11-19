/**
 * WebApp.gs – Trang web cập nhật Register ("phiếu điện tử").
 * Triển khai: Apps Script → Triển khai → Tùy chọn triển khai mới → Ứng dụng web
 *   Thực thi với tư cách: NGƯỜI DÙNG TRUY CẬP (để nhật ký ghi đúng email, tôn trọng phân quyền khóa cột)
 *   Ai có quyền truy cập: Bất kỳ ai có tài khoản Google (người đó phải được chia sẻ quyền sửa file).
 * Web chỉ ghi vào khối TQN QUẢN LÝ (kết quả theo phiếu, Chuyển tới, Link) – không đụng cột gốc.
 * Web chặt hơn sửa ô trực tiếp: sai mẫu phiếu / đóng mà thiếu link / sai thứ tự cấp → không cho Lưu.
 */
function doGet() {
  return tqnHtml_('Web').setTitle('TQN Register – cập nhật hồ sơ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Lõi: tách lỗi CHẶN (không cho lưu) và lỗi chỉ CẢNH BÁO. */
function tqnWebKiemLuu_(rec, p) {
  var d = tqnQlDanhGia_(rec, p), chan = [], canh = [];
  d.canhBao.forEach(function (c) { (/^Tình trạng|^Phiếu đã đóng|chưa có ngày trình/.test(c) ? canh : chan).push(c); });
  return { d: d, chan: chan, canh: canh };
}

function tqnWebNguoi_() {
  var u = ''; try { u = Session.getActiveUser().getEmail(); } catch (e) { }
  return u || '(không rõ email)';
}

/** Đọc mọi sổ nhóm 1 đã có khối quản lý → dữ liệu cho web. */
function tqnWebDuLieu() {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), rows = [], loai = {};
  cfg.detailSheets.forEach(function (name) {
    var type = tqnQlLoaiTab_(cfg, prof, name), sh = type && tqnSs_().getSheetByName(name);
    if (!sh) return;
    var p = prof[type], k = tqnQlKhoi_(sh, cfg, p, false);
    if (!k || sh.getLastRow() < cfg.firstRow) return;
    loai[type] = { levels: p.levels.map(function (L) { return { name: L.name, options: L.options.map(function (o) { return { code: o.code, text: o.text, kq: o.kq, label: o.label }; }) }; }), forward: p.forward };
    var n = sh.getLastRow() - cfg.firstRow + 1, v = sh.getRange(cfg.firstRow, 1, n, Math.max(sh.getLastColumn(), k.kt)).getValues();
    var lk = sh.getRange(cfg.firstRow, k.link, n, 1).getRichTextValues(), m = k.map;
    v.forEach(function (r, i) {
      if (!String(r[m.subject] || '').trim()) return;
      var url = (lk[i][0] && lk[i][0].getLinkUrl()) || '';
      rows.push({ tab: name, row: cfg.firstRow + i, type: type, ma: String(r[m.code] || ''), nd: String(r[m.subject] || ''),
        nguoi: m.owner >= 0 ? String(r[m.owner] || '') : '', tt: String(r[m.status] || ''),
        picks: k.lv.map(function (c) { return c > 0 ? String(r[c - 1] || '') : ''; }), to: k.to > 0 ? String(r[k.to - 1] || '') : '',
        link: url, giu: String(r[k.giu - 1] || ''), kt: String(r[k.kt - 1] || '') });
    });
  });
  return { me: tqnWebNguoi_(), loai: loai, rows: rows, luc: Utilities.formatDate(new Date(), cfg.timezone || 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM') };
}

/** Lõi: Tình trạng nên có theo kết quả phiếu ('' = không đổi). */
function tqnWebTinhTrang_(d, coChon) {
  if (!coChon) return '';
  if (d.nhom === 'DONG') return 'Đã duyệt';
  if (d.nhom === 'NT') return 'Không duyệt';
  if (d.nhom === 'CHO') return 'Đã trình duyệt';
  return '';
}

/** Tạo hồ sơ mới từ web: ghi Nội dung + Người thực hiện vào dòng trống đầu tiên sau dòng cuối, rồi tự cấp mã. */
function tqnWebTaoMoi(o) {
  var cfg = tqnCodeCfg_(), sh = tqnSs_().getSheetByName(o.tab);
  if (!sh || cfg.detailSheets.indexOf(o.tab) < 0) throw new Error('Không có sổ ' + o.tab);
  if (!String(o.nd || '').trim()) throw new Error('Chưa gõ nội dung');
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Đang có người lưu, thử lại sau vài giây');
  try {
    var hdr = sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, Math.max(sh.getLastColumn(), 1)).getValues(), m = tqnDetailMap_(hdr);
    var last = sh.getLastRow(), col = sh.getRange(1, m.subject + 1, Math.max(last, 1), 1).getValues(), r = cfg.firstRow;
    for (var i = col.length - 1; i >= cfg.firstRow - 1; i--) if (String(col[i][0]).trim()) { r = i + 2; break; }
    sh.getRange(r, m.subject + 1).setValue(String(o.nd).trim());
    if (m.owner >= 0 && o.nguoi) sh.getRange(r, m.owner + 1).setValue(String(o.nguoi).trim());
    var type = tqnAutoTypeOfTab_(cfg, o.tab), ma = '';
    if (type) { var c = tqnCapMaTrong_(sh, cfg, type, r, r, true); ma = c.length ? c[0].code : ''; }
    tqnQlGhiNhatKy_([[new Date(), tqnWebNguoi_(), o.tab, 'Dòng ' + r, ma, 'Tạo mới', '', String(o.nd).trim(), 'WEB']]);
    return { row: r, ma: ma };
  } finally { lock.releaseLock(); }
}

/** Lưu phiếu từ web. o = {tab,row,picks:[nhãn],to,link}. Trả dòng mới hoặc ném lỗi tiếng Việt. */
function tqnWebLuu(o) {
  var cfg = tqnConfig_(), prof = tqnQlMau_(), sh = tqnSs_().getSheetByName(o.tab), type = sh && tqnQlLoaiTab_(cfg, prof, o.tab);
  if (!type) throw new Error('Không tìm thấy sổ ' + o.tab);
  var p = prof[type], k = tqnQlKhoi_(sh, cfg, p, false);
  if (!k) throw new Error('Sổ ' + o.tab + ' chưa cài khối quản lý');
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('Đang có người lưu, thử lại sau vài giây');
  try {
    var r = o.row, m = k.map, row = sh.getRange(r, 1, 1, Math.max(sh.getLastColumn(), k.kt)).getValues()[0];
    if (o.ma && String(row[m.code]) !== o.ma) throw new Error('Dòng ' + r + ' đã đổi (mã khác) – tải lại trang');
    var ngay = [m.nt && m.nt.plan, m.nt && m.nt.actual].filter(function (c) { return c != null && c >= 0; });
    var rec = { noiDung: row[m.subject], tinhTrang: row[m.status], picks: o.picks || [], chuyenToi: o.to || '', link: String(o.link || '').trim(),
      daTrinh: ngay.some(function (c) { return tqnIsDate_(row[c]); }) || /da trinh|cho duyet/.test(tqnQlBo_(row[m.status])) };
    var kq = tqnWebKiemLuu_(rec, p);
    if (kq.chan.length) throw new Error('Chưa lưu: ' + kq.chan.join('; '));
    var cu = k.lv.map(function (c) { return String(row[c - 1] || ''); }), user = tqnWebNguoi_(), now = new Date(), log = [];
    k.lv.forEach(function (c, i) {
      var v = String((o.picks || [])[i] || '');
      if (v !== cu[i]) { sh.getRange(r, c).setValue(v); log.push([now, user, o.tab, sh.getRange(r, c).getA1Notation(), row[m.code], p.levels[i].name + ' · theo phiếu', cu[i], v, 'WEB']); }
    });
    if (k.to > 0 && String(o.to || '') !== String(row[k.to - 1] || '')) { sh.getRange(r, k.to).setValue(o.to || ''); log.push([now, user, o.tab, sh.getRange(r, k.to).getA1Notation(), row[m.code], 'Chuyển tới', row[k.to - 1], o.to || '', 'WEB']); }
    var linkCu = (sh.getRange(r, k.link).getRichTextValue() || { getLinkUrl: function () { return ''; } }).getLinkUrl() || '';
    if (rec.link !== linkCu) {
      sh.getRange(r, k.link).setValue(rec.link);
      log.push([now, user, o.tab, sh.getRange(r, k.link).getA1Notation(), row[m.code], 'Link OneDrive', linkCu, rec.link, 'WEB ' + tqnQlXuLyLink_(sh, k, r, rec.link)]);
    }
    var td = tqnQlTuDien_(sh, p, k, r, row, o.picks || [], cu, user, now, log, kq.d);
    if (td.stMoi) kq.canh = kq.canh.filter(function (c) { return !/^Tình trạng|^Phiếu đã đóng/.test(c); });
    tqnQlTinhDong_(sh, cfg, p, k, r, r);
    kq.canh.forEach(function (c) { log.push([now, user, o.tab, 'Dòng ' + r, row[m.code], 'Kiểm tra', '', '⚠ ' + c, 'CẢNH BÁO']); });
    if (log.length) tqnQlGhiNhatKy_(log);
    var g = sh.getRange(r, k.giu).getValue(), t = sh.getRange(r, k.kt).getValue();
    return { giu: String(g), kt: String(t), canh: kq.canh, soThayDoi: log.length };
  } finally { lock.releaseLock(); }
}

function tqnMenuMoWeb() {
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  var ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Trang web chưa triển khai', 'Apps Script → Triển khai → Tùy chọn triển khai mới → Ứng dụng web\n• Thực thi với tư cách: Người dùng truy cập\n• Quyền truy cập: Bất kỳ ai có tài khoản Google\nRồi bấm lại mục này.', ui.ButtonSet.OK); return; }
  ui.showModalDialog(HtmlService.createHtmlOutput('<p style="font:14px Arial">Link trang cập nhật (gửi cho đội):</p><input style="width:100%" value="' + url + '" onclick="this.select()"><p><a href="' + url + '" target="_blank">Mở trang web ↗</a></p>').setWidth(520).setHeight(160), 'TQN Register – trang web');
}
