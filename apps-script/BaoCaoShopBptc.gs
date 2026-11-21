
/* ===== TQN BÁO CÁO SHOP DRAWING + BPTC cho sếp (bản 2 – 21/11) – tab "BC SHOP-BPTC"
 * Đầu trang theo mẫu phiếu RFI. Chỉ đọc tab SD, M0S (không sửa sổ gốc).
 * Phân loại theo NGÀY THỰC TẾ trong sổ, không đoán theo chữ:
 *   ĐÃ TRÌNH = có ngày "Thực tế" của Nhà thầu đệ trình.  Đã duyệt = Tình trạng ghi rõ đã duyệt/chấp thuận.
 *   Trả về = ghi trả về/không duyệt/trình lại.  Đã phản hồi = có ngày thực tế TVGS/CĐT nhưng chưa ghi kết quả.  Còn lại = đang chờ duyệt.
 *   CHƯA TRÌNH: "Đã hoàn thành" = NT làm xong, chờ trình (KHÔNG phải đã duyệt) · "Đang thực hiện" · trống = kế hoạch. */
var TQN_BC = { TAB: 'BC SHOP-BPTC', NGUON: [['SD', 'Shop drawing'], ['M0S', 'Biện pháp thi công']], NGAY_CHO: 7,
  IMG: { COA: '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', TBC: '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', XUAT: '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', MOI: 'iVBORw0KGgoAAAANSUhEUgAAAJkAAAAuCAIAAAC08SwgAAAFmUlEQVR42u2ae0xTVxzHfy198yhPC5SCYApSoGDBjYFTyeaTZehghgwY8zH9k2XT7D/dP0ucyQjZH/sH5bFpMoFEYVnMNpuhwXQ+KM4yHgN5tRAEa+FCgduH3R/Xe3eGtVKKCDfnG/749XDur+eezz3n/H6/Xk7cuXOAxQpx8RRglliYJRZmiYVZYpZYmCUWZomFWWIh4r2uLw4UCuuLijTR0QRJ7q2rGyMI1sxpXVHRzvh4q81W3tR0b3SU/etyhiTLGhvvmkxBQmFVfj6Xw2EHyHKNZmd8/AxJljQ0rCZIAPAL3rXrdd223en8pbc3Uy5/U6Egnc67JtN6B6kMC/u+oMBqs5U2Nv41Pr7K387BtXUc+2B5JyGPV5KRkSmXr4PY5+HJk370mffOhQsPnzx56SUbQ0Jajx1DW3bX1v7z+LFn5wBQ2db2nU63qM/ZPXuK1Wrm4/DU1I7qaq+G7YvnbIXi/eTkLLlcFhDgLxAQJDlKEHeMxssGA3VHpMPxWU6OfmzsBH2ILmPG1u66LExJWdTywXMtblWSkcHj/m/kUpHogErl+5CW4TnC3//SoUM/FRd/lJ6eGB4uFYl4XG6oWJwmkx3Nyvr18OHm0lJ/gQAA5u32ELGYhXssxx25gyqV3xICWllAwN7ERLSlWK0W8VZgj/HWc6hEcqWkJDcu7lk2otfvqK5OrKzcX1//e38/dZvpUVF/V1TUFhbGBgc/tlpZyDI7NlYeFETZNqeTmcrcjRuXcvknGs1/98DhlGVkrNTAvPL87b59MVIpZde0t3+l1Q5PTdmczq6JiRNXr94aHmZ65iUkLDgc9R0dLGRZhCzKb27edNv+vBYcjlmbDQCy5HLVhg3PDptNm6gJnfThqV+GZ3VkZF5CAmXP2+2VbW3of5+6XGdv3EBb3q2puW00so2lhM9ntrIxgqi5d2/IYnkW/iiVgUKhh5S0wWBYtIAOZ2ZSxqX7931Jdr31vFupZGyd0Ug9CqgMjx6h9aw0mYyFOcm+xEQqHACAlp4eF8DPPT3URxGPl5+U5OHaHzo6XAAAUJCcHCIWK8PCcmJjqZVxmYaxPHnrOT0ykrF7Jyfd+uxFwnI10p89LAtTUxm7ubsbAFq6uz3Et6iGLJbWgQEqYytWq5k1dKWriyBJX0blredQiYSxpxYW3Pq0zM8zdjjSnyUsowIDsxUKyu43m7snJgCgz2zuoR/trTExccHBHjzU6fWU8fGWLQdp8PV0oy/yyvNSKsgcJCx3sa/uU5iSwlTSm5HliNqeE82bg4ODFgv1WEj4fADQjYz0uisyeCuvPJvn5hg7WCRy2wdtR/uzhSWywX6xbdvQqVPU35fbt6O8PTz1rufWSu1KLEpvPaPV86SICLd9ksLDGfvBK662rzZLTXR0fEjIS7vFSKVv0PuwWzV2dlrpuHGUILT9/Ss1wqV7/q2vj7HfUigC6GiOUapMFk3n0FabDU03gQW/RRchizLv/PlBOhVhELYdP84kmh6yMavN1tTZWa7RAMCPHR1O14odRkv3/GB8/I+BASrFFPP5FTk5X7e2onUGdKep1et9DM3WFkuBn997mzdT9qPZ2UUgAcA0PW2anqbS8/1JSae12nm7/UXezmi1Z7TaVzHOpXs+ee1aS1kZVcD6dOtWPy63tr19fHZ2U2jo57m5b9M1rNtGY9WtW7Ae3xHRHj3qdoLMVmsQXQfQjYy4vVY3MvJhWhoA+AsEe5XKK11dsIZlnps7cPFiVX4+VZI9kpl5hC4vMAdwo8Fw+vp1x9OnrHrfB416/nzB/qkzGimW1Ia8xlkCwKTVWtLQkK1QFKhU1G9eEj5/hiRNBHHHaGwwGFYkwAb8XgHg9wqwMEsszBILs8TCLDFLLMwSC7PEwiwxSyzMEguzxMIssTBL9utfIshd8GG+B20AAAAASUVORK5CYII=' } };
var TQN_BC_NHAN = { DUYET: 'Đã duyệt', CHO: 'Chờ duyệt', PHANHOI: 'Đã phản hồi', TRAVE: 'Trả về', XONG: 'Hoàn thành – chờ trình', LAM: 'Đang thực hiện', KH: 'Kế hoạch' };
var TQN_BC_MAU = { DUYET: ['#E6F4EA', '#1E7B3A'], CHO: ['#FEF3D6', '#8A6100'], PHANHOI: ['#E8F0FE', '#1A56B8'], TRAVE: ['#FCE8E6', '#C5221F'], XONG: ['#F1EDFC', '#5B3FD1'], LAM: ['#F1F3F6', '#4D5A6B'], KH: ['#FFFFFF', '#8792A2'] };
function tqnBcBo_(s) { return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim(); }
function tqnBcNgay_(v) { if (v instanceof Date && !isNaN(v)) return v; var m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(String(v || '').trim()); if (!m) return null; var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]); }
function tqnBcLoai_(tt, ngayTrinh, coPhanHoi) {
  var t = tqnBcBo_(tt);
  if (/khong duyet|tu choi|tra ve|trinh lai|sua lai|bo sung/.test(t)) return 'TRAVE';
  if (/da duyet|da phe duyet|chap thuan|duyet co y kien/.test(t)) return 'DUYET';
  if (ngayTrinh || /da trinh|cho duyet|cho phe duyet/.test(t)) return coPhanHoi ? 'PHANHOI' : 'CHO';
  if (/hoan thanh/.test(t)) return 'XONG';
  if (/dang thuc hien|dang lam/.test(t)) return 'LAM';
  return 'KH';
}
function tqnBcDoc_(ss, tab) {
  var sh = ss.getSheetByName(tab); if (!sh) return null;
  var v = sh.getDataRange().getValues(), H = v.slice(0, Math.min(6, v.length)), col = {}, nhom = {};
  H.forEach(function (r) { r.forEach(function (x, c) { var t = tqnBcBo_(x);
    if (t === 'ma ho so' && col.ma == null) col.ma = c; else if (t.indexOf('noi dung') === 0 && col.nd == null) col.nd = c;
    else if (t.indexOf('nguoi thuc hien') === 0 && col.nguoi == null) col.nguoi = c; else if (t.indexOf('tinh trang') === 0 && col.tt == null) col.tt = c;
    else if (t === 'rev' && col.rev == null) col.rev = c;
    else if (t.indexOf('nha thau de trinh') === 0) nhom.nt = c; else if (t.indexOf('tu van giam sat') === 0) nhom.tv = c; else if (t.indexOf('chu dau tu') === 0) nhom.cdt = c; }); });
  var thucTe = function (g) { var kq = null; if (g == null) return null; H.forEach(function (r) { for (var c = g; c < g + 5 && c < r.length; c++) if (kq == null && /^thuc te$/.test(tqnBcBo_(r[c]))) kq = c; }); return kq; };
  var vuong = function (g) { var kq = null; if (g == null) return null; H.forEach(function (r) { for (var c = g; c < g + 5 && c < r.length; c++) if (kq == null && /^vuong mac/.test(tqnBcBo_(r[c]))) kq = c; }); return kq; };
  col.ngay = thucTe(nhom.nt); col.tv = thucTe(nhom.tv); col.cdt = thucTe(nhom.cdt);
  col.vm = [vuong(nhom.nt), vuong(nhom.tv), vuong(nhom.cdt)].filter(function (x) { return x != null; });
  if (col.ma == null || col.nd == null) return { loi: 'Không dò được cột Mã/Nội dung ở tab ' + tab, ds: [] };
  var dau = 0; H.forEach(function (r, i) { if (r.some(function (x) { return /^stt$/i.test(String(x).trim()); })) dau = i + 1; });
  var ds = [], homNay = new Date();
  for (var i = dau; i < v.length; i++) {
    var r = v[i], ma = String(r[col.ma] || '').replace(/\s+/g, ' ').trim(), nd = String(r[col.nd] || '').replace(/\s+/g, ' ').trim();
    if (!ma || !/\d/.test(ma)) continue;                               // bỏ dòng mục (KẾT CẤU, I, NĂM 2026…)
    var tt = col.tt != null ? String(r[col.tt] || '').trim() : '', ng = col.ngay != null ? tqnBcNgay_(r[col.ngay]) : null;
    var ph = (col.tv != null && tqnBcNgay_(r[col.tv])) || (col.cdt != null && tqnBcNgay_(r[col.cdt])) || null;
    if (ng && ng > homNay) ng = null;                                    // ngày tương lai = kế hoạch, chưa phải đã trình
    var loai = tqnBcLoai_(tt, ng, ph), cho = (loai === 'CHO' && ng) ? Math.floor((homNay - ng) / 864e5) : '';
    var vm = col.vm.map(function (c) { return String(r[c] || '').trim(); }).filter(String).join(' · ');
    ds.push({ ma: ma, nd: nd, nguoi: col.nguoi != null ? String(r[col.nguoi] || '').trim() : '', rev: col.rev != null ? String(r[col.rev] === '' ? '' : r[col.rev]) : '',
      tt: tt, loai: loai, ngay: ng, ph: ph, cho: cho, vm: vm });
  }
  return { ds: tqnBcGopBo_(ds) };
}
/** Gộp shop dầm + sàn cùng tầng, cùng công trình thành 1 bộ (anh Nho yêu cầu 21/11).
 *  Trạng thái bộ = thành phần chậm nhất (bộ chỉ xong khi đủ dầm + sàn); có thành phần bị trả về thì bộ = trả về. */
function tqnBcGopBo_(ds) {
  var HANG = { KH: 0, LAM: 1, XONG: 2, CHO: 3, PHANHOI: 4, DUYET: 5 }, nhom = {}, ra = [];
  ds.forEach(function (o) {
    var m = /^(.*?th[eé]p\s+)(d[aầ]m|s[aà]n)\s+(.+)$/i.exec(o.nd), loai = m && tqnBcBo_(m[2]);
    if (!m || /mong/.test(tqnBcBo_(m[3])) && loai === 'dam' && !/tang/.test(tqnBcBo_(m[3]))) { ra.push(o); return; }
    var k = tqnBcBo_(m[1] + '|' + m[3]); (nhom[k] = nhom[k] || { dau: m[1], sau: m[3], ds: [] }).ds.push(o);
  });
  Object.keys(nhom).forEach(function (k) {
    var g = nhom[k]; if (g.ds.length < 2) { ra.push(g.ds[0]); return; }
    var tra = g.ds.filter(function (o) { return o.loai === 'TRAVE'; }).length, cham = g.ds.slice().sort(function (a, b) { return HANG[a.loai] - HANG[b.loai]; })[0];
    var ng = g.ds.map(function (o) { return o.ngay; }).filter(Boolean).sort(function (a, b) { return b - a; })[0] || null;
    var loai = tra ? 'TRAVE' : cham.loai, cho = (loai === 'CHO' && ng) ? Math.floor((new Date() - ng) / 864e5) : '';
    ra.push({ ma: g.ds.map(function (o) { return o.ma; }).join('\n'), nd: g.dau.replace(/\s+$/, '') + ' dầm + sàn ' + g.sau + '  (bộ ' + g.ds.length + ' bản vẽ)',
      nguoi: g.ds.map(function (o) { return o.nguoi; }).filter(function (x, i, a) { return x && a.indexOf(x) === i; }).join(', '),
      rev: g.ds.map(function (o) { return o.rev; }).filter(function (x, i, a) { return x !== '' && a.indexOf(x) === i; }).join('/'),
      tt: '', loai: loai, ngay: ng, cho: cho, vm: g.ds.map(function (o) { return o.vm; }).filter(String).join(' · '), bo: g.ds.length });
  });
  return ra;
}
function tqnBcAnh_(sh, b64, mime, col, row, ox, oy) { return sh.insertImage(Utilities.newBlob(Utilities.base64Decode(b64), mime, 'h'), col, row, ox || 0, oy || 0); }
/** Tạo / làm mới tờ báo cáo */
function tqnTaoBaoCaoShopBptc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(TQN_BC.TAB) || ss.insertSheet(TQN_BC.TAB), tz = ss.getSpreadsheetTimeZone(), homNay = new Date();
  sh.getImages().forEach(function (im) { im.remove(); });
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) { }
  sh.clear(); if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  var N = 9, R = function (r, c, nr, nc) { return sh.getRange(r, c, nr || 1, nc || 1); }, F = 'Arial', GRID = '#C9D1DB', NAVY = '#1B3F7A';
  [42, 205, 360, 46, 150, 86, 72, 120, 190].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); }); sh.setColumnWidth(10, 24); sh.setColumnWidth(11, 230);
  R(1, 1, sh.getMaxRows(), N).setFontFamily(F).setFontSize(10).setVerticalAlignment('middle').setFontColor('#18212D');
  // --- Đầu trang theo mẫu phiếu RFI ---
  sh.setRowHeights(1, 3, 24); R(1, 1, 3, 3).merge(); R(1, 4, 3, 3).merge(); R(1, 7, 3, 3).merge();
  R(1, 1, 3, N).setBorder(true, true, true, true, true, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  R(4, 1, 1, 5).merge().setValue('Dự án: Khu đô thị sinh thái Coastal Quảng Ngãi'); R(4, 6, 1, 4).merge().setValue('Địa chỉ: xã Tư Nghĩa, tỉnh Quảng Ngãi');
  R(4, 1, 1, N).setFontFamily('Times New Roman').setFontSize(12).setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID); R(4, 1).setHorizontalAlignment('left'); sh.setRowHeight(4, 28);
  tqnBcAnh_(sh, TQN_BC.IMG.COA, 'image/jpeg', 2, 1, 55, 4).setWidth(150).setHeight(70);
  tqnBcAnh_(sh, TQN_BC.IMG.TBC, 'image/png', 8, 1, 86, 4).setWidth(97).setHeight(66);
  sh.setRowHeight(5, 10);
  R(6, 1, 1, N).merge().setValue('BÁO CÁO TÌNH HÌNH TRÌNH DUYỆT SHOP DRAWING & BIỆN PHÁP THI CÔNG').setBackground('#C6E0B4').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center').setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(6, 34);
  R(7, 1, 1, N).merge().setValue('Tính đến: ' + Utilities.formatDate(homNay, tz, 'HH:mm, dd/MM/yyyy') + '      Đơn vị lập: TBCONS').setFontStyle('italic').setFontColor('#4D5A6B').setHorizontalAlignment('center'); sh.setRowHeight(7, 22);
  // --- Tổng hợp ---
  var data = TQN_BC.NGUON.map(function (n) { return { tab: n[0], ten: n[1], d: tqnBcDoc_(ss, n[0]) }; }), row = 9;
  R(row, 1, 1, N).merge().setValue('1. TỔNG HỢP').setFontWeight('bold').setFontSize(11).setFontColor(NAVY); sh.setRowHeight(row, 24); row++;
  var h1 = row; R(h1, 1, 2, 1).merge().setValue('STT'); R(h1, 2, 2, 1).merge().setValue('Hạng mục');
  R(h1, 3, 1, 5).merge().setValue('HỒ SƠ ĐÃ TRÌNH'); R(h1, 8, 1, 2).merge().setValue('CHƯA TRÌNH');
  R(h1 + 1, 3, 1, 7).setValues([['Tổng đã trình', 'Đã duyệt', 'Chờ duyệt', 'Trả về', 'Chờ > ' + TQN_BC.NGAY_CHO + ' ngày', 'Đang chuẩn bị', 'Kế hoạch']]);
  R(h1, 1, 2, 7).setBackground(NAVY).setFontColor('#FFFFFF'); R(h1, 8, 2, 2).setBackground('#5F6B7A').setFontColor('#FFFFFF');
  R(h1, 1, 2, N).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true); sh.setRowHeight(h1, 24); sh.setRowHeight(h1 + 1, 28);
  row = h1 + 2; var tongPh = 0;
  data.forEach(function (x, k) {
    var ds = (x.d && x.d.ds) || [], d = function (l) { return ds.filter(function (o) { return o.loai === l; }).length; };
    var trinh = ds.filter(function (o) { return /DUYET|CHO|PHANHOI|TRAVE/.test(o.loai); }).length, tre = ds.filter(function (o) { return o.cho !== '' && o.cho > TQN_BC.NGAY_CHO; }).length;
    tongPh += d('PHANHOI');
    R(row, 1, 1, N).setValues([[k + 1, x.ten + (x.d ? '' : ' (không có tab ' + x.tab + ')'), trinh, d('DUYET'), d('CHO') + d('PHANHOI'), d('TRAVE'), tre, d('XONG') + d('LAM'), d('KH')]]);
    R(row, 1, 1, N).setHorizontalAlignment('center').setFontSize(12); R(row, 2).setHorizontalAlignment('left').setFontWeight('bold').setFontSize(11);
    R(row, 3).setFontWeight('bold'); R(row, 4).setFontColor('#1E7B3A').setFontWeight('bold'); R(row, 5).setFontColor('#8A6100').setFontWeight('bold');
    R(row, 6).setFontColor(d('TRAVE') ? '#C5221F' : '#18212D'); R(row, 7).setFontColor(tre ? '#C5221F' : '#18212D').setFontWeight(tre ? 'bold' : 'normal');
    R(row, 8, 1, 2).setFontColor('#5F6B7A'); sh.setRowHeight(row, 30); row++;
  });
  R(h1, 1, row - h1, N).setBorder(true, true, true, true, true, true, GRID, SpreadsheetApp.BorderStyle.SOLID);
  R(h1, 8, row - h1, 1).setBorder(null, true, null, null, null, null, '#5F6B7A', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  R(row, 1, 1, N).merge().setValue('Ghi chú: Đã trình = có ngày thực tế nhà thầu đệ trình.  Đang chuẩn bị = Hoàn thành (nhà thầu làm xong, CHƯA trình, CHƯA duyệt) + Đang thực hiện.  Shop dầm + sàn cùng tầng, cùng công trình tính là 1 bộ.' + (tongPh ? '  Có ' + tongPh + ' hồ sơ TVGS/CĐT đã phản hồi nhưng sổ chưa ghi kết quả (tính vào Chờ duyệt).' : '')).setFontSize(9).setFontStyle('italic').setFontColor('#8792A2').setWrap(true); sh.setRowHeight(row, 30);
  row += 2;
  // --- Chi tiết hồ sơ đã trình ---
  var thuTu = { TRAVE: 1, CHO: 2, PHANHOI: 3, DUYET: 4 }, so = 2;
  data.forEach(function (x, k) {
    var ds = ((x.d && x.d.ds) || []).filter(function (o) { return thuTu[o.loai]; }).sort(function (a, b) { return thuTu[a.loai] - thuTu[b.loai] || (b.cho || 0) - (a.cho || 0) || a.ma.localeCompare(b.ma); });
    R(row, 1, 1, N).merge().setValue((so++) + '. ' + x.ten.toUpperCase() + ' – HỒ SƠ ĐÃ TRÌNH (' + ds.length + ')').setFontWeight('bold').setFontSize(11).setFontColor(NAVY); sh.setRowHeight(row, 24); row++;
    R(row, 1, 1, N).setValues([['STT', 'Mã hồ sơ', 'Nội dung', 'Rev', 'Tình trạng', 'Ngày trình', 'Số ngày chờ', 'Người thực hiện', 'Vướng mắc / ý kiến']]).setFontWeight('bold').setBackground('#EEF2F7').setHorizontalAlignment('center').setWrap(true); sh.setRowHeight(row, 26);
    var hd = row; row++;
    if (x.d && x.d.loi) { R(row, 1, 1, N).merge().setValue(x.d.loi).setFontColor('#C5221F'); row += 2; return; }
    if (!ds.length) { R(row, 1, 1, N).merge().setValue('Chưa có hồ sơ đã trình').setFontColor('#8792A2').setHorizontalAlignment('center'); R(hd, 1, 2, N).setBorder(true, true, true, true, true, true, GRID, SpreadsheetApp.BorderStyle.SOLID); row += 2; return; }
    var vals = ds.map(function (o, i) { return [i + 1, o.ma, o.nd, o.rev, TQN_BC_NHAN[o.loai], o.ngay ? Utilities.formatDate(o.ngay, tz, 'dd/MM/yyyy') : '', o.cho === '' ? '' : o.cho, o.nguoi, o.vm]; });
    var rg = R(row, 1, vals.length, N); rg.setValues(vals).setVerticalAlignment('middle');
    R(row, 3, vals.length, 1).setWrap(true); R(row, 9, vals.length, 1).setWrap(true).setFontSize(9);
    R(row, 2, vals.length, 1).setFontSize(9).setWrap(true);
    R(row, 1, vals.length, 1).setHorizontalAlignment('center'); R(row, 4, vals.length, 5).setHorizontalAlignment('center');
    ds.forEach(function (o, i) { var m = TQN_BC_MAU[o.loai]; R(row + i, 5).setBackground(m[0]).setFontColor(m[1]).setFontWeight('bold');
      if (i % 2) { R(row + i, 1, 1, 4).setBackground('#FAFBFC'); R(row + i, 6, 1, 4).setBackground('#FAFBFC'); }
      if (o.cho !== '' && o.cho > TQN_BC.NGAY_CHO) R(row + i, 7).setFontColor('#C5221F').setFontWeight('bold'); });
    R(hd, 1, vals.length + 1, N).setBorder(true, true, true, true, true, true, GRID, SpreadsheetApp.BorderStyle.SOLID);
    row += vals.length + 1;
  });
  // --- Hồ sơ chưa trình nhưng đã làm (hoàn thành / đang thực hiện) – gọn ---
  var cho = []; data.forEach(function (x) { ((x.d && x.d.ds) || []).forEach(function (o) { if (o.loai === 'XONG' || o.loai === 'LAM') cho.push({ x: x, o: o }); }); });
  var kh = 0; data.forEach(function (x) { kh += ((x.d && x.d.ds) || []).filter(function (o) { return o.loai === 'KH'; }).length; });
  R(row, 1, 1, N).merge().setValue((so++) + '. HỒ SƠ ĐANG CHUẨN BỊ, CHƯA TRÌNH (' + cho.length + ')').setFontWeight('bold').setFontSize(11).setFontColor(NAVY); sh.setRowHeight(row, 24); row++;
  if (cho.length) {
    R(row, 1, 1, N).setValues([['STT', 'Mã hồ sơ', 'Nội dung', '', 'Tình trạng', 'Hạng mục', '', 'Người thực hiện', 'Ghi chú']]).setFontWeight('bold').setBackground('#EEF2F7').setHorizontalAlignment('center');
    R(row, 6, 1, 2).merge(); var hd2 = row; row++;
    cho.sort(function (a, b) { return (a.o.loai === 'XONG' ? 0 : 1) - (b.o.loai === 'XONG' ? 0 : 1); });
    cho.forEach(function (c, i) { R(row, 1, 1, N).setValues([[i + 1, c.o.ma, c.o.nd, '', TQN_BC_NHAN[c.o.loai], c.x.ten, '', c.o.nguoi, c.o.vm]]); R(row, 6, 1, 2).merge();
      var m = TQN_BC_MAU[c.o.loai]; R(row, 5).setBackground(m[0]).setFontColor(m[1]).setFontWeight('bold'); R(row, 2).setFontSize(9); R(row, 3).setWrap(true); R(row, 1).setHorizontalAlignment('center'); R(row, 5, 1, 3).setHorizontalAlignment('center'); row++; });
    R(hd2, 1, cho.length + 1, N).setBorder(true, true, true, true, true, true, GRID, SpreadsheetApp.BorderStyle.SOLID);
  }
  R(row, 1, 1, N).merge().setValue('Ngoài ra còn ' + kh + ' hồ sơ trong kế hoạch chưa bắt đầu (xem tab SD, M0S).').setFontStyle('italic').setFontColor('#8792A2'); row += 2;
  // --- Ký tên ---
  R(row, 2, 1, 2).merge().setValue('NGƯỜI LẬP'); R(row, 6, 1, 4).merge().setValue('CHỈ HUY TRƯỞNG / BAN CHỈ HUY');
  R(row, 1, 1, N).setFontWeight('bold').setHorizontalAlignment('center'); R(row + 1, 2, 1, 2).merge().setValue('(Ký, ghi rõ họ tên)'); R(row + 1, 6, 1, 4).merge().setValue('(Ký, ghi rõ họ tên)');
  R(row + 1, 1, 1, N).setFontStyle('italic').setFontColor('#8792A2').setHorizontalAlignment('center'); sh.setRowHeight(row + 2, 60); row += 3;
  // --- Nút (ngoài vùng in) ---
  tqnBcAnh_(sh, TQN_BC.IMG.XUAT, 'image/png', 11, 6, 6, 0).assignScript('tqnXuatBaoCaoShopBptc');
  tqnBcAnh_(sh, TQN_BC.IMG.MOI, 'image/png', 11, 9, 6, 0).assignScript('tqnTaoBaoCaoShopBptc');
  if (sh.getMaxRows() > row + 5) sh.deleteRows(row + 5, sh.getMaxRows() - row - 4);
  sh.setHiddenGridlines(true); ss.setActiveSheet(sh); sh.setActiveSelection('A1'); SpreadsheetApp.flush();
  return { tab: TQN_BC.TAB, dongCuoi: row };
}
/** Nút Xuất: máy chủ tạo PDF (A4 ngang, vừa trang, cột A:I) + Excel, LƯU VÀO DRIVE thư mục "TQN – Báo cáo Shop BPTC",
 *  rồi hiện hộp có nút Mở PDF / Tải PDF / Tải Excel. Không phụ thuộc trình duyệt chặn link. */
function tqnXuatBaoCaoShopBptc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(TQN_BC.TAB); if (!sh) { tqnTaoBaoCaoShopBptc(); sh = ss.getSheetByName(TQN_BC.TAB); }
  SpreadsheetApp.flush();
  var tz = ss.getSpreadsheetTimeZone(), ten = 'BC_Shop_BPTC_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
  var base = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?gid=' + sh.getSheetId(), last = sh.getLastRow();
  var hd = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  var pdfUrl = base + '&format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenum=CENTER&horizontal_alignment=CENTER&top_margin=0.35&bottom_margin=0.35&left_margin=0.3&right_margin=0.3&range=A1:I' + last;
  var r = UrlFetchApp.fetch(pdfUrl, hd);
  if (r.getResponseCode() !== 200) throw new Error('Google không xuất được PDF (mã ' + r.getResponseCode() + '). Thử lại sau ít phút.');
  var dp = PropertiesService.getDocumentProperties(), fid = dp.getProperty('TQN_BC_FOLDER'), folder;
  try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = DriveApp.createFolder('TQN – Báo cáo Shop BPTC'); dp.setProperty('TQN_BC_FOLDER', folder.getId()); }
  var pdf = folder.createFile(r.getBlob().setName(ten + '.pdf'));
  var x = UrlFetchApp.fetch(base + '&format=xlsx', hd), xlsx = x.getResponseCode() === 200 ? folder.createFile(x.getBlob().setName(ten + '.xlsx')) : null;
  var dl = function (f) { return 'https://drive.google.com/uc?export=download&id=' + f.getId(); };
  var b = 'display:inline-block;color:#fff;padding:10px 16px;border-radius:7px;text-decoration:none;font-weight:bold;margin:0 8px 8px 0;cursor:pointer;border:0;font:bold 13px Arial';
  var h = '<div style="font:14px Arial;padding:4px 6px"><p style="margin:2px 0 12px">✓ Đã lưu vào Drive, thư mục <b>' + folder.getName() + '</b>:<br><b>' + ten + '.pdf</b>' + (xlsx ? ' và .xlsx' : '') + '</p>'
    + '<button style="' + b + ';background:#1B3F7A" onclick="window.open(\'' + pdf.getUrl() + '\',\'_blank\')">Mở PDF</button>'
    + '<button style="' + b + ';background:#2F6FDB" onclick="window.open(\'' + dl(pdf) + '\',\'_blank\')">⤓ Tải PDF</button>'
    + (xlsx ? '<button style="' + b + ';background:#1F8A8A" onclick="window.open(\'' + dl(xlsx) + '\',\'_blank\')">⤓ Tải Excel</button>' : '')
    + '<p style="margin:6px 0 0"><a href="' + folder.getUrl() + '" target="_blank">Mở thư mục báo cáo trên Drive</a></p>'
    + '<p style="color:#6b7280;font-size:12px;margin-top:10px">Cần số liệu mới nhất: bấm LÀM MỚI trước khi xuất.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(470).setHeight(230), 'Xuất báo cáo Shop drawing / BPTC');
}
