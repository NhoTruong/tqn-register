/**
 * Manager.gs – BẢNG QUẢN LÝ HỒ SƠ (bố cục gọn: 3 con số · 3 nút chính · việc cần xử lý trước).
 * Dữ liệu lấy từ các hàm đã kiểm (tqnRunAll_, tqnDocsData). Chỉ ghi vào tab TQN_*, trừ "Đọc phiếu → Ghi vào"
 * (chỉ ghi vào Ô TRỐNG của sổ chi tiết, có xác nhận, có nhật ký).
 */
var TQN_PHIEU_HEADERS = ['Thời điểm', 'Người tạo', 'Loại', 'Số hiệu', 'Rev', 'Nội dung', 'File PDF', 'Đã in', 'Phiếu giao nhận'];
var TQN_GN_HEADERS = ['Thời điểm', 'Người lập', 'Số phiếu giao nhận', 'Số hồ sơ', 'Danh sách số hiệu', 'File PDF'];
var TQN_PHIEU_DEFAULTS = [
  ['phieu_ben_gui', 'Công ty TNHH TBCONS', 'Phiếu: Ban hành bởi – Công ty'],
  ['phieu_gui_toi', 'Công ty Cổ Phần Đầu Tư Kinh Doanh BĐS NHSLand', 'Phiếu: Gửi tới – Công ty'],
  ['phieu_nguoi_lap', 'Đặng Tuấn Thịnh', 'Phiếu: Người lập'],
  ['phieu_nguoi_nhan', '', 'Phiếu: Người nhận mặc định'],
  ['phieu_so_hop_dong', '', 'Phiếu: Số hợp đồng'],
  ['phieu_so_ngay_ph', '3', 'Phiếu: số ngày chờ phản hồi ghi trên phiếu'],
  ['phieu_du_an', 'Khu đô thị sinh thái Coastal Quảng Ngãi', 'Phiếu: tên dự án ở đầu phiếu'],
  ['phieu_dia_chi', 'xã Tư Nghĩa, tỉnh Quảng Ngãi', 'Phiếu: địa chỉ dự án'],
  ['phieu_thu_muc', 'TQN_PHIEU', 'Tên thư mục Google Drive lưu PDF phiếu'],
  ['code_TML', '{yymmdd}-COA-TBC-TML-{nnn}', 'Mẫu mã phiếu giao nhận']
];
var TQN_GROUPS = { SDS: 1, MSS: 1, MAS: 1, OTS: 1, SCH: 1, RFA: 1, REP: 1, HSE: 1, COS: 1, 'Thanh toán': 1, 'Mua sắm': 1, RFI: 2, RFC: 2, TML: 3, 'Thông báo': 3, 'Lấy ý kiến': 3, 'Công văn': 3, NCR: 4 };
var TQN_GROUP_NAMES = ['', 'Trình duyệt', 'Hỏi · đổi', 'Giao nhận · văn bản', 'Hiện trường'];

function tqnEnsurePhieuConfig_() {
  tqnEnsureConfig_();
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_PHIEU_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}
function tqnPhieuCfg_() {
  var cfg = tqnConfig_(), o = {};
  TQN_PHIEU_DEFAULTS.forEach(function (r) { o[r[0]] = cfg[r[0]] !== undefined && cfg[r[0]] !== '' ? String(cfg[r[0]]) : r[1]; });
  return o;
}
function tqnReadOwn_(name, headers) {
  var sh = tqnSs_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
}
function tqnKeyOf_(so) { return tqnCleanCode_(so).toUpperCase().replace(/\s+/g, ''); }

/** Mọi dữ liệu cho bảng quản lý. */
/** nhanh = true: dùng bảng TQN_QUY_TRINH đã có (tự cập nhật mỗi giờ / lần "Nhập và kiểm" gần nhất), không đọc lại 32 tab. */
function tqnWfFromSheet_() {
  var sh = tqnSs_().getSheetByName(TQN.PREFIX + 'QUY_TRINH');
  if (!sh || sh.getLastRow() < 2) return null;
  return { rows: sh.getRange(2, 1, sh.getLastRow() - 1, TQN_WF_HEADERS.length).getValues() };
}
function tqnManagerData(nhanh) {
  tqnEnsurePhieuConfig_();
  var wf = nhanh ? tqnWfFromSheet_() : null, tz;
  if (wf) tz = tqnConfig_().timezone; else { var r = tqnRunAll_(); wf = r.wf; tz = r.cfg.timezone; }
  var now = new Date(), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var phieu = {}, gn = {};
  tqnReadOwn_(TQN.PREFIX + 'PHIEU', TQN_PHIEU_HEADERS).forEach(function (p) { phieu[tqnKeyOf_(p[3])] = { pdf: p[6], in: p[7], gn: p[8] }; });
  var rows = wf.rows.map(function (x) {
    var step = parseInt(String(x[0]), 10) || 1;
    var d = tqnIsDate_(x[9]) ? Math.round((new Date(x[9].getFullYear(), x[9].getMonth(), x[9].getDate()) - t0) / 86400000) : '';
    var sla = step === 5 ? 'XONG' : (x[11] ? 'QUA_HAN' : (d === '' ? (step === 1 ? 'CHUA' : 'KHONG_HAN') : (d <= 2 ? 'SAP_HAN' : 'CON_HAN')));
    var loai = String(x[2] || ''), ph = phieu[tqnKeyOf_(x[3])] || null;
    var next = step === 4 ? 'TRINH_LAI' : (step === 2 || step === 3) ? 'NHAC' : (step === 1 ? (ph ? 'GIAO_NHAN' : 'TAO_PHIEU') : '');
    return { buoc: step, buocTen: x[0], bic: x[1], loai: loai, nhom: TQN_GROUPS[loai] || 1, so: String(x[3] || ''), rev: x[4], nd: String(x[5] || ''),
      nguoi: x[6], den: x[7], ngay: tqnFmtDate_(x[8], tz), han: tqnFmtDate_(x[9], tz), d: d, cho: x[10], sla: sla, y: x[12], at: x[13],
      phieu: ph ? (ph.pdf || 'Đã tạo') : '', gn: ph ? ph.gn : '', next: next };
  });
  var docs = tqnDocsData(now).map(function (x) { x.nhom = /NCR|Chỉ thị/.test(x.loai) ? 4 : 3; return x; });
  var k = { quaHan: 0, sapHan: 0, cho: 0, traVe: 0, chuaPhieu: 0, xong: 0, tong: rows.length + docs.length };
  rows.forEach(function (x) {
    if (x.sla === 'QUA_HAN') k.quaHan++; else if (x.sla === 'SAP_HAN') k.sapHan++;
    if (x.buoc === 2 || x.buoc === 3) k.cho++;
    if (x.buoc === 4) k.traVe++;
    if (x.buoc === 1 && !x.phieu) k.chuaPhieu++;
    if (x.buoc === 5) k.xong++;
  });
  docs.forEach(function (x) { if (x.tt !== 'Đã phản hồi') { k.cho++; if (x.cho > 7) k.quaHan++; } });
  var nhom = [0, 0, 0, 0, 0];
  rows.forEach(function (x) { nhom[x.nhom]++; }); docs.forEach(function (x) { nhom[x.nhom]++; });
  var gnList = tqnReadOwn_(TQN.PREFIX + 'GIAO_NHAN', TQN_GN_HEADERS).map(function (g) {
    return { luc: tqnIsDate_(g[0]) ? Utilities.formatDate(g[0], tz, 'dd/MM/yyyy HH:mm') : String(g[0]), nguoi: g[1], so: g[2], n: g[3], ds: g[4], pdf: g[5] };
  }).reverse();
  var log = [];
  var au = tqnSs_().getSheetByName(TQN.SH_AUDIT);
  if (au && au.getLastRow() >= 2) {
    var n = Math.min(40, au.getLastRow() - 1);
    log = au.getRange(au.getLastRow() - n + 1, 1, n, 4).getValues().reverse().map(function (a) {
      return [tqnIsDate_(a[0]) ? Utilities.formatDate(a[0], tz, 'dd/MM HH:mm') : String(a[0]), String(a[1]), String(a[2]), String(a[3]).slice(0, 160)];
    });
  }
  var ai = tqnAiCfg_();
  return { at: Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm'), kpi: k, nhom: nhom, nhomTen: TQN_GROUP_NAMES, rows: rows, docs: docs, gn: gnList, log: log,
    cfg: tqnPhieuCfg_(), aiOn: !tqnAiBlocked_(ai), aiWhy: tqnAiBlocked_(ai) };
}

/* ---------------- Phiếu (RFA / RFI / RFC) và phiếu giao nhận ---------------- */
var TQN_LOGO = { coa: 'data:image/jpeg;base64,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Lxu24RSp8sXfeN+dwl1W7uml4xrnxz+OnxB8c+JfA/7NfgfwvPbeE9Ui07VfEvjDV5o7ae98tZZrW3htUkdzGkkQaR2QB2KhX2k16Z8ONX8e614WjvPiX4UsNG1hZ5o7ix0zVDeQbVkZUkSVo4yVdArgFAVD4PINePN4T+PP7NXjrxbqfwc+Fun+N/DPjTxG+utpsWupp+oaZqEyItzk3AMU0DtGJQQyOrO67WG012/7MPg74k+EPA2qTfFWSBdW1vxVqWsGyttSe7Swiup2ljthKyJu8tWC/KoXjjing54r6y1V5m/evdWilf3baa3XZt99RZhRwSwSlQdNRXJa0r1JNx9/mXM2rPe8UtuXTV7Px2+KcXwZ+FWsfEFdNa/u7OBI9L0yNWLX19NIsNtbDaCQZJ5I484435PAqp+z58W7r4xeAH1rXNOtbDXNK1W60jxLpdrc+atlqFtIY5UB67WwsqZ5MUsbfxVyXx2+EPj345fGDwpo9v4hvPD/AIX8LpJrbazpy2ctxcawD5VrEsdzFMoWKNp5SzRkFni24ZNwi+F3wR+JPwb/AGidW8S/8J7q3i3w9400hJ9evNcNolxp+q2oWGF0S3SCPy5rY+W2yEvutYyzEHhSr49ZkmoS9l8L0Vtr82/NvaO1t2/Ip4fKnkzTqR9u7zWsr6O3s9uTWN5781+WNrm9+2N8avEv7O/7N3ib4x+D9PsLrUtFgge1g1RHaBi9xFEd4RlYjDk8MOcV6YhYjLY69q8j/bu+G/jP4u/speLfh18PdH+36zqUFstjZ+fHF5pS6ikYbpGVR8qMeSM4wOSBXrcLBowPTg12QliHj5p/Byxtppe8+bX0UTzqiwqyuk429pz1L668qjS5dO13Oz669tHUUUV2HAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAU1oo3+8uadRQAioqfdGM0MocbWHFLRQAAADAowM5oooACAetFFFACbRjGKWiigBCoOcjrSgBRgCiigAooooACAeDR0oooAQqD1FL0oooAAAOlFFFACMoYYYUtFFFgCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD/9k=', tbc: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAANMAAACoCAMAAABNPCvpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAJYUExURQAAABpgehtgehdeexpfeRlfeRtgeBtfehROdhldegBmZhlgeBBQcBxfewAAABphexlgehhgehheeRhgeBpfehtgehteeRdeeBxfehpfehtdehpfehthexxgexteeRpfeRhfehxVVRhfeBpgehpfeRxfexpfeRleeABVVRpfeRteeBtheRtfeh5aeBpfeRlgehpfeQ5jcRteeRhfexpgeBdddBpgehpfdxpgeBldehtgexNeexpfdxpfeBlfehpfeRlfexhfehRYdhtdeRlgeRpfeRldeRlfeRpgeh1geBpgextgeRZeehpeeRljextdehpeeRRidiRtbRRmcBtfeBhfeRteeh1fex5ieBtfeBpfeQAzZhpfeQ9aeBlfeBpfeRpeehpfeRpeeBpgeh1ddBRccBtgehddfBpfeRtheRphexxfexlfeDNmZhdddBtheRxheRRcehdeeRleeBpeexpfeRVgdQBgYBldextfehpfeRpeexFedxdedwAAgBpgeRpfeBtdeBtgehtgeBpeeBpgeQCAgBpfeRpgehpgeRxeexpeeRxeexlgehleeRtfeBtgehlheBxgfBZeeRlddhhgeQBVahpfeRpgexhfehtfehpeehhgdxteeRtgehpfegBAYB1idhteehpgdxhfeBlfehpeeRpheBdaeBlfdxpgehleeBpcehhgehhfeRVgdRphehlddxteehpfeRlgeRFmdxtgehNecRtfehpfeRlgeABNZhheehpeeRteeRhdeBddfBdddBteeBlfeBtheBpheRpfeRldfOpv3x4AAADIdFJOUwCI/0/77lWjDawFkBCLATqY0VQg69dMRNXWrKrOuVTL2wmygPJT/tgD8eB80xHkkKYSOYnHC1rNikdoG77aq9JmcxqDz36R+X8156guyB/KpQ0HGdPlQT4i4bkF3BFb7ev84N8sGcAh819sZjMFIWk/GUxRTzsYCDRLTjweTwIoRkLEn8e4AvliUDbzG/bi8/yhQDk07Az9h+tepddfqowIGrxt4Y69WSKrsNoyvLEl1oOt9VIPrBtWdOQKQfUTVUIW/ei2kbop+yC7hQAAAAlwSFlzAAAh1QAAIdUBBJy0nQAABRxJREFUeF7tnOd7G0UQh5ejBzCEQOgEAgRDQg+9l4QSCD10QjXNlBAIJUDoNaG3hN57773Dv8WzZ5Wb9+Tb02lOa/uZ99vO7s38XlmWLFm2c4ZhGIZhGIZhGIZhGIZRldWSAni4I6vzqmLWWHMtdtCm706etddhE1WiOCXJuuyiSSSnkr2rEc0pmcRGasRzKte9ChGd1mMnLSI6JeuzlRIxncr1756oThuwlw5RncoN6Bpz6kQvTgNspkJcpw3ZTIW4ThuxmQpxnSazmQo1O208ZROWsmzKZirU7DTVuc1YyxDBafMQvkPIyW3BYpsITkF8h6DTliy2GbdOW7HYZmvGUSGu0zaMo0Jcp20ZR4W4ThF+NpoWwncYZ0483JFenNhLh6hO27GXDlGdtmcvHfrgNJ3FJjuwlRJ9cGKtyY7spEXtTjvNYK1JbW/EBp12ZjGLP1DoVMAgo6gRzamep6aUSE67MIcmcZx2ZQxVYjjNZAhl+uw0azcGqIE+OyVJMmP3PRhCmf47efbcizk0ieOUJHsziCI1O+0zm5Um+zKJHjU7TXVuP9Ya1Pd4Xr+Tm8xig/2ZRYs+OI36+ukAhlEiplNyINPoENUpvV6fuE4HMY4KcZ3q+UJFdjqYeTSI7HQI82gQ2amWO585EX/AnAjzaGBOxB8wJ8I8GpgT8QfMiTCPBuZE/AFzIsyjgTkRf8CcCPNoEHQ6lMUs/kDQ6TAWszCPBkEn1gT+QNDpcBazMI8GIacjWBP4DiGnI1kTMI8GhU5HHX0MSxLfodBpztxjWZIwjwaFTkF8h0KnIMyjwUR0Oo5DusJ3GHtOx3NIN5zgO4w9p0EO6YZ5vsPYcyp+/glwom/Qk9NJjKMCp3RD2qAnp/mMo8LJHFOeU9IGPTkxjRIcU55T0+t7cTqNYZQ4nYPKcsbI9b04MYsaZ3JSSRqX9+C0AEkU4ahyND9UWN3pLOTQ5GwOK8M5zasrO82WKbThuDDntq6t6pT+QViNnHc+Jwbwr/UaVHS6IDu/Hi7kzCIWZj9zXMnpoj4o+Y8FXczBo3CJzFPB6dLLRIc6uXyIwztwBa/q1unKev54uoD5V109KtcMX8vjzrnreKyA62/g1YZhGIYxOouGb1yc5aabl4j9W269TezfvvSO5tadYuMut0islzZODSy7W9QX3zOl3b0GBu/lU33Kfa0D93Mr5YEH082HRPFh94hYD6VnHhW1Jo8tb03QZgVntZiT7rPa5nG/LZ1WdHCaKSqCkZtFHY7J8oRz7kkWswzQKck7PSUK4GnG0eAZThE49yxLgudyTs/TafAFUSDMo8GLYsLKaWKZOLdKrGe9JJb+AJwW0GlYrHMwjwbydzUv477o3Cti/apbLtZ5p9cCTu51rGugW6c3xDrv9GbICd/AzKOBtlNiTuPD6S2xmhhOb4vVxHCSdHB6R8A8GuScgL5T/eScxLprp3fFqoMTYR4NlJ3eEyvv9P4HQB5gHg2Unfhr/CE3aYlgHD5GIHH+vmdO1dB2wj8QnBBOH4rlWHD66ONPxJpOn372uVjnnb4Qy7zTl1/JNfNoIF8TEr4mzJFz+los806EeTTgg6/EuW9YEnybc/pOLP37Ef1/7V74Hsr3zrkfWMzwY+59I367DDn3k6xIavovRz9zTpNffk33f2O9xe9+O+jk/vhT1tr8xSxqDPzNWSn/tA78y62UeSO3cdjJuf9kscHCkdvMMAzDMAzDMAzDMAzDMIw8/wMqdGCr/355RQAAAABJRU5ErkJggg==' };
var TQN_PAPER_CSS = '<style>.paper{background:#fff;padding:22px 28px;margin:0 auto 18px;max-width:760px;font-family:Arial,sans-serif;font-size:12px;color:#000;page-break-after:always}'
  + '.paper table{width:100%;border-collapse:collapse}.paper td{border:1px solid #000;padding:4px 6px;vertical-align:middle}.paper .ttl{text-align:center;font-weight:bold}'
  + '.paper .pj{font-family:"Times New Roman",serif;font-size:13px;display:flex;justify-content:space-between;padding:4px 8px}.paper .vv{font-size:15px;margin:10px 0 6px}'
  + '.paper .bh{font-weight:bold;margin:10px 0 6px}.paper .gray{background:#d9d9d9;text-align:center;font-weight:bold}.paper .sm td{font-size:11px}</style>';
function tqnEsc_(s) { return String(s === null || s === undefined ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function tqnPaperHead_(c) {
  return '<table><tr><td style="width:30%;height:70px;text-align:center"><img src="' + TQN_LOGO.coa + '" style="height:46px"></td><td style="width:34%"></td>'
    + '<td style="text-align:center"><img src="' + TQN_LOGO.tbc + '" style="height:46px"></td></tr><tr><td colspan="3" style="padding:0"><div class="pj"><span>Dự án: '
    + tqnEsc_(c.phieu_du_an) + '</span><span>Địa chỉ: ' + tqnEsc_(c.phieu_dia_chi) + '</span></div></td></tr></table><div style="height:10px"></div>';
}
function tqnPaperInfo_(c, title, color, so, ngay, han, labHan) {
  var dots = '……………………';
  return '<table><tr><td colspan="4" class="ttl" style="background:' + color + '">' + title + '</td></tr>'
    + '<tr><td style="width:21%">Ban hành bởi</td><td>Công ty: ' + tqnEsc_(c.phieu_ben_gui || dots) + '</td><td style="width:21%">Gửi tới</td><td>Công ty: ' + tqnEsc_(c.phieu_gui_toi || dots) + '</td></tr>'
    + '<tr><td>Người lập</td><td>Ông/ Bà : ' + tqnEsc_(c.phieu_nguoi_lap || dots) + '</td><td>Người nhận</td><td>Ông/ Bà : ' + tqnEsc_(c.phieu_nguoi_nhan || dots) + '</td></tr>'
    + '<tr><td style="height:38px">Ký xác nhận</td><td></td><td></td><td></td></tr><tr><td>Ngày ban hành</td><td>' + ngay + '</td><td>C/C</td><td></td></tr>'
    + '<tr><td>Số tham chiếu</td><td>' + tqnEsc_(so) + '</td><td>' + labHan + '</td><td>' + han + '</td></tr></table>';
}
function tqnPhieuHtml_(x, c, ngay, han) {
  var m = x.loai === 'RFI' ? 'RFI' : (x.loai === 'RFC' ? 'RFC' : 'RFA');
  var T = { RFA: ['YÊU CẦU PHÊ DUYỆT', '#FBD4B4', 'BÊN BAN HÀNH', 'MÔ TẢ THÔNG TIN ĐƯỢC PHÊ DUYỆT', 'Nhà thầu TBCONS đệ trình '],
    RFI: ['YÊU CẦU THÔNG TIN', '#D6E3BC', 'BÊN GỬI', 'MÔ TẢ THÔNG TIN HOẶC SỰ XÁC NHẬN CẦN CÓ', 'Nhà thầu TBCONS đề nghị làm rõ thông tin: '],
    RFC: ['YÊU CẦU THAY ĐỔI', '#D6E3BC', 'BÊN BAN HÀNH', 'CÁC THAY ĐỔI ĐƯỢC YÊU CẦU', 'Nhà thầu TBCONS đề xuất thay đổi: '] }[m];
  var resp = function (who) {
    return '<table class="sm" style="margin-top:8px"><tr><td colspan="4" style="background:#7f7f7f;color:#fff;font-weight:bold">' + who + '</td></tr><tr><td colspan="4" class="ttl">PHẢN HỒI</td></tr>'
      + '<tr><td style="width:5%" class="ttl">A</td><td>☐ Chấp thuận yêu cầu</td><td style="width:5%" class="ttl">C</td><td>☐ Vui lòng xem ý kiến phản hồi và đệ trình lại</td></tr>'
      + '<tr><td class="ttl">B</td><td>☐ Chấp thuận với điều kiện hoặc ý kiến phản hồi bên dưới</td><td class="ttl">D</td><td>☐ Từ chối chấp thuận yêu cầu</td></tr>'
      + '<tr><td colspan="4" style="height:44px;vertical-align:top"><u>Ý kiến:</u></td></tr></table>';
  };
  return '<div class="paper">' + tqnPaperHead_(c) + tqnPaperInfo_(c, T[0], T[1], x.so, ngay, han, 'Hy vọng phản hồi vào')
    + '<div class="vv">Về việc: ' + tqnEsc_(x.nd) + '</div><div class="bh">' + T[2] + '</div>'
    + '<table><tr><td class="gray">' + T[3] + '</td></tr><tr><td style="height:70px;vertical-align:top">• Căn cứ Hợp đồng số ' + tqnEsc_(c.phieu_so_hop_dong || '…')
    + '<br>• Căn cứ Hồ sơ thiết kế bản vẽ thi công được phê duyệt.<br>&nbsp;&nbsp;&nbsp;' + T[4] + tqnEsc_(x.nd) + '</td></tr>'
    + (m === 'RFC' ? '<tr><td>Ảnh hưởng dẫn tới từ những thay đổi: Tiến độ : … &nbsp; Thiết kế : … &nbsp; Chi phí : …</td></tr>' : '')
    + '<tr><td style="height:34px;vertical-align:top">Đính kèm (nếu có):</td></tr></table>'
    + '<div class="bh">PHẢN HỒI TỪ BÊN NHẬN</div>' + resp('TVQLDA/TVGS') + resp('CHỦ ĐẦU TƯ/BQLXD') + '</div>';
}
function tqnTrnHtml_(so, list, c, ngay) {
  return '<div class="paper">' + tqnPaperHead_(c) + tqnPaperInfo_(c, 'PHIẾU GIAO NHẬN HỒ SƠ', '#FBD4B4', so, ngay, String(list.length), 'Số lượng hồ sơ')
    + '<div class="vv">Về việc: Giao nhận ' + list.length + ' hồ sơ</div><div class="bh">DANH MỤC HỒ SƠ GIAO</div><table class="sm">'
    + '<tr style="background:#FCE4D6;font-weight:bold;text-align:center"><td style="width:6%">STT</td><td style="width:27%">Số tham chiếu</td><td>Nội dung hồ sơ</td><td style="width:6%">Rev</td><td style="width:8%">Bản cứng</td><td style="width:11%">Ghi chú</td></tr>'
    + list.map(function (x, i) { return '<tr><td style="text-align:center">' + (i + 1) + '</td><td>' + tqnEsc_(x.so) + '</td><td>' + tqnEsc_(x.nd) + '</td><td style="text-align:center">' + tqnEsc_(x.rev) + '</td><td style="text-align:center">1</td><td></td></tr>'; }).join('')
    + '</table><div style="margin:6px 0 10px">Đính kèm:</div><table class="sm"><tr><td class="gray">BÊN GIAO</td><td class="gray">BÊN NHẬN</td></tr>'
    + '<tr><td style="height:70px;text-align:center;vertical-align:top">Ký, ghi rõ họ tên</td><td style="text-align:center;vertical-align:top">Ký, ghi rõ họ tên</td></tr>'
    + '<tr><td style="text-align:center">Ngày: …/…/20…</td><td style="text-align:center">Ngày: …/…/20…</td></tr></table></div>';
}
/** Lưu PDF vào thư mục Drive (nếu được cấp quyền Drive). Lỗi thì bỏ qua, phiếu vẫn xem/in được. */
function tqnSavePdf_(html, name, c) {
  try {
    if (typeof DriveApp === 'undefined') return '';
    var it = DriveApp.getFoldersByName(c.phieu_thu_muc), f = it.hasNext() ? it.next() : DriveApp.createFolder(c.phieu_thu_muc);
    var blob = Utilities.newBlob('<html><head><meta charset="utf-8">' + TQN_PAPER_CSS + '</head><body>' + html + '</body></html>', 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
    return f.createFile(blob).getUrl();
  } catch (e) { return ''; }
}
function tqnFindRows_(sos) {
  var d = tqnManagerData(true), want = {};
  sos.forEach(function (s) { want[tqnKeyOf_(s)] = 1; });
  return { d: d, list: d.rows.filter(function (x) { return want[tqnKeyOf_(x.so)]; }) };
}
/** Tạo phiếu cho các hồ sơ được chọn → HTML để xem/in + PDF trong Drive + ghi TQN_PHIEU. */
function tqnMakeForms(sos) {
  if (!sos || !sos.length) throw new Error('Chưa chọn hồ sơ nào.');
  var f = tqnFindRows_(sos), c = f.d.cfg, tz = tqnConfig_().timezone, now = new Date();
  var ngay = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
  var han = Utilities.formatDate(new Date(now.getTime() + (parseInt(c.phieu_so_ngay_ph, 10) || 3) * 86400000), tz, 'dd/MM/yyyy');
  var sh = tqnOwnSheet_(TQN.PREFIX + 'PHIEU', TQN_PHIEU_HEADERS), out = [], pdf = 0;
  f.list.forEach(function (x) {
    var h = tqnPhieuHtml_(x, c, ngay, han), url = tqnSavePdf_(h, String(x.so).replace(/[\\\/:*?"<>|]/g, '_'), c);
    if (url) pdf++;
    sh.appendRow([now, Session.getActiveUser().getEmail(), x.loai, x.so, x.rev, x.nd, url, '', '']);
    out.push(h);
  });
  tqnAudit_('TAO_PHIEU', f.list.map(function (x) { return x.so; }).join(', '), { phieu: f.list.length });
  return { n: f.list.length, pdf: pdf, html: TQN_PAPER_CSS + out.join('') };
}
/** Lập phiếu giao nhận: cấp số TML (khóa chống trùng như cấp mã), ghi TQN_GIAO_NHAN, đánh dấu phiếu. */
function tqnMakeTransmittal(sos) {
  if (!sos || !sos.length) throw new Error('Chưa chọn hồ sơ nào.');
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(20000)) throw new Error('Có người khác đang cấp số, thử lại sau.');
  try {
    var f = tqnFindRows_(sos), c = f.d.cfg, cfg = tqnConfig_(), tz = cfg.timezone, now = new Date();
    var used = tqnReadOwn_(TQN.PREFIX + 'GIAO_NHAN', TQN_GN_HEADERS).map(function (g) { return String(g[2]); });
    var re = /-(\d{3,4})\s*$/, max = 0;
    used.forEach(function (u) { var m = re.exec(u); if (m) max = Math.max(max, parseInt(m[1], 10)); });
    var so = String(c.code_TML).replace('{yymmdd}', Utilities.formatDate(now, tz, 'yyyyMMdd').slice(2)).replace('{nnn}', ('000' + (max + 1)).slice(-3));
    var h = tqnTrnHtml_(so, f.list, c, Utilities.formatDate(now, tz, 'dd/MM/yyyy'));
    var url = tqnSavePdf_(h, so, c);
    tqnOwnSheet_(TQN.PREFIX + 'GIAO_NHAN', TQN_GN_HEADERS).appendRow([now, Session.getActiveUser().getEmail(), so, f.list.length, f.list.map(function (x) { return x.so; }).join('\n'), url]);
    var ph = tqnSs_().getSheetByName(TQN.PREFIX + 'PHIEU');
    if (ph && ph.getLastRow() >= 2) {
      var v = ph.getRange(2, 4, ph.getLastRow() - 1, 1).getValues(), want = {};
      f.list.forEach(function (x) { want[tqnKeyOf_(x.so)] = 1; });
      v.forEach(function (r, i) { if (want[tqnKeyOf_(r[0])]) ph.getRange(i + 2, 9).setValue(so); });
    }
    tqnAudit_('GIAO_NHAN', so, { ho_so: f.list.length });
    return { so: so, n: f.list.length, pdf: url ? 1 : 0, html: TQN_PAPER_CSS + h };
  } finally { lock.releaseLock(); }
}
/** Nhãn gáy bìa hồ sơ. */
function tqnSpineLabels(sos) {
  var f = tqnFindRows_(sos), c = f.d.cfg;
  return '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;font-family:Arial">' + f.list.map(function (x) {
    return '<div style="border:2px solid #1B3F7A;border-radius:6px;padding:8px 10px"><div style="font-size:11px;color:#1B3F7A">' + tqnEsc_(c.phieu_du_an).toUpperCase()
      + '</div><div style="font-family:monospace;font-size:14px;font-weight:bold">' + tqnEsc_(x.so) + '</div><div style="font-size:12px">' + tqnEsc_(x.nd) + '</div></div>';
  }).join('') + '</div>';
}
/** Nháp văn bản nhắc phản hồi (không cần AI, không tự gửi). */
function tqnReminderText(so) {
  var f = tqnFindRows_([so]), x = f.list[0], c = f.d.cfg;
  if (!x) throw new Error('Không thấy hồ sơ ' + so);
  var ai = /CĐT|chủ đầu tư/i.test(x.bic) ? 'Chủ đầu tư' : (/TVTK/i.test(x.bic) ? 'Tư vấn thiết kế' : 'Tư vấn giám sát');
  return 'Kính gửi ' + ai + ',\n\n' + c.phieu_ben_gui + ' đã trình hồ sơ ' + x.so + ' (' + x.nd + ')' + (x.ngay ? ' ngày ' + x.ngay : '')
    + '. ' + (x.han ? 'Hạn phản hồi theo quy định là ' + x.han + (x.sla === 'QUA_HAN' ? ', hiện đã quá hạn' : '') + '. ' : '')
    + 'Kính đề nghị Quý ' + ai + ' xem xét và phản hồi sớm để Nhà thầu kịp triển khai thi công theo tiến độ.\n\nTrân trọng,\n' + c.phieu_nguoi_lap;
}
function tqnFolderUrl() {
  var c = tqnPhieuCfg_();
  try { var it = DriveApp.getFoldersByName(c.phieu_thu_muc); return it.hasNext() ? it.next().getUrl() : ''; } catch (e) { return ''; }
}
function tqnMenuManager() {
  var html = tqnHtml_('Manager').setWidth(1180).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quản lý hồ sơ trình duyệt');
}

/* ---------------- Đọc phiếu bằng AI (ảnh/PDF phiếu có ý kiến TVGS/CĐT) ---------------- */
var TQN_FORM_PROMPT = 'Đây là ảnh/bản scan một phiếu hồ sơ trình duyệt (RFA/RFI/RFC/biện pháp/bản vẽ) của công trường xây dựng tại Việt Nam, có thể có ý kiến phản hồi viết tay hoặc đánh máy và con dấu. '
  + 'Đọc phiếu và trả lời ĐÚNG MỘT đối tượng JSON, không thêm chữ nào khác:\n'
  + '{"so_hieu":"số tham chiếu in trên phiếu, giữ nguyên","rev":"số rev nếu có","ben_phan_hoi":"TVGS|CĐT|TVTK|không rõ",'
  + '"ket_qua":"Đã duyệt|Duyệt có ý kiến|Không duyệt|Trả về|chưa có","ngay_phan_hoi":"dd/mm/yyyy hoặc rỗng","nguoi_ky":"tên hoặc rỗng",'
  + '"y_kien":"chép ý kiến phản hồi, gọn","do_tin_cay":"CAO|TRUNG_BINH|THAP","ghi_chu":"chỗ nào đọc không rõ"}\n'
  + 'Quy tắc: ô đánh dấu A = Đã duyệt, B = Duyệt có ý kiến, C = Trả về, D = Không duyệt. Không đoán chữ không đọc được: ghi vào ghi_chu và hạ do_tin_cay.';

function tqnFormParse_(text) {
  var s = String(text || ''), a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a < 0 || b <= a) throw new Error('AI không trả về JSON');
  var o = JSON.parse(s.slice(a, b + 1));
  if (!o.so_hieu) throw new Error('AI không đọc được số tham chiếu trên phiếu');
  ['rev', 'ben_phan_hoi', 'ket_qua', 'ngay_phan_hoi', 'nguoi_ky', 'y_kien', 'do_tin_cay', 'ghi_chu'].forEach(function (k) { o[k] = String(o[k] || ''); });
  if (o.ngay_phan_hoi && !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(o.ngay_phan_hoi)) { o.ghi_chu += ' (ngày "' + o.ngay_phan_hoi + '" không đúng dạng, bỏ)'; o.ngay_phan_hoi = ''; }
  return o;
}
/** So khớp kết quả đọc phiếu với sổ chi tiết; trả đề xuất ghi (chỉ ô trống). */
function tqnFormMatch_(o) {
  var cfg = tqnConfig_(), src = tqnReadAll_(cfg), key = tqnDocKey_(tqnCleanCode_(o.so_hieu));
  var cand = src.details.filter(function (d) { return d.key === key; });
  if (!cand.length) return { found: false, msg: 'Không thấy số hiệu ' + o.so_hieu + ' trong các sổ chi tiết.' };
  cand.sort(function (a, b) { return (parseInt(b.rev, 10) || 0) - (parseInt(a.rev, 10) || 0); });
  var d = cand.filter(function (x) { return !o.rev || parseInt(x.rev, 10) === parseInt(o.rev, 10); })[0] || cand[0];
  var m = src.maps[d.sheet], blk = /CĐT|chu dau tu/i.test(tqnNoAccent_(o.ben_phan_hoi)) || /CĐT/.test(o.ben_phan_hoi) ? m.cdt : (/TVTK/.test(o.ben_phan_hoi) ? m.tvtk : m.tvgs);
  var blkTen = blk === m.cdt ? 'CĐT' : (blk === m.tvtk ? 'TVTK' : 'TVGS');
  var sh = tqnSs_().getSheetByName(d.sheet), rowV = sh.getRange(d.row, 1, 1, Math.min(sh.getLastColumn(), 40)).getValues()[0];
  var p = [];
  var add = function (ten, col, val) {
    if (col === undefined || col < 0 || !val) return;
    var cur = rowV[col], empty = cur === '' || cur === null;
    p.push({ ten: ten, o: tqnColLetter_(col + 1) + d.row, col: col + 1, cur: tqnIsDate_(cur) ? Utilities.formatDate(cur, cfg.timezone, 'dd/MM/yyyy') : String(cur || ''), moi: val, ghi: empty });
  };
  add('Tình trạng hồ sơ', m.status, o.ket_qua && o.ket_qua !== 'chưa có' ? o.ket_qua : '');
  if (blk) { add(blkTen + ' – ngày thực tế', blk.actual, o.ngay_phan_hoi); add(blkTen + ' – ý kiến / vướng mắc', blk.note, o.y_kien); }
  return { found: true, at: d.sheet + '!' + d.row, sheet: d.sheet, row: d.row, so: d.code, nd: d.subject, de_xuat: p };
}
/** Gọi AI đọc 1 ảnh/PDF (base64). Kết quả CHỈ là đề xuất, chưa ghi gì. */
function tqnAiReadForm(b64, mime) {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error('Chưa bật AI: ' + why);
  if (String(b64 || '').length > 7000000) throw new Error('File quá lớn (tối đa ~5 MB). Chụp lại ảnh nhỏ hơn.');
  var part = /pdf/i.test(mime) ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } }
    : { type: 'image', source: { type: 'base64', media_type: mime || 'image/jpeg', data: b64 } };
  var o = tqnFormParse_(tqnAiCall_(cfg, [part, { type: 'text', text: TQN_FORM_PROMPT }]));
  var r = tqnFormMatch_(o); r.ai = o;
  tqnAudit_('AI_DOC_PHIEU', o.so_hieu, { tin_cay: o.do_tin_cay });
  return r;
}
/** Ghi các đề xuất đã chọn – CHỈ vào ô đang trống. Ô đã có chữ thì bỏ qua (không đè). */
function tqnAiApplyForm(sheet, row, items) {
  var cfg = tqnConfig_(), dl = (cfg.detail_sheets || '').split(',').map(function (s) { return s.trim(); });
  if (dl.indexOf(sheet) < 0) throw new Error('Chỉ ghi vào sổ chi tiết đã khai báo.');
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(20000)) throw new Error('Sổ đang bận, thử lại.');
  try {
    var sh = tqnSs_().getSheetByName(sheet), ok = [], skip = [];
    (items || []).forEach(function (it) {
      var rg = sh.getRange(row, it.col), cur = rg.getValue();
      if (cur !== '' && cur !== null) { skip.push(it.ten); return; }
      var v = it.moi, m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v));
      rg.setValue(m ? new Date(+m[3], +m[2] - 1, +m[1]) : v); ok.push(it.ten);
    });
    tqnAudit_('AI_GHI_PHIEU', sheet + '!' + row + ': ' + ok.join(', ') + (skip.length ? ' | bỏ (ô đã có): ' + skip.join(', ') : ''), { ghi: ok.length, bo: skip.length });
    return { ghi: ok, bo: skip };
  } finally { lock.releaseLock(); }
}
