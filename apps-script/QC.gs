/**
 * QC.gs – kiểm lỗi dữ liệu. Chỉ BÁO, không sửa nguồn. Kết quả ghi TQN_QC.
 */
var TQN_QC_HEADERS = ['Mức', 'Mã luật', 'Vị trí', 'Số hiệu', 'Chi tiết', 'Kiểm lúc'];

function tqnQC_(cfg, imp) {
  var src = imp.src, out = [], now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  var add = function (lv, rule, at, code, msg) { out.push([lv, rule, at, code, msg, now]); };
  var isSubmitted = function (s) { var t = tqnLower_(s); return cfg.statusSubmitted.some(function (k) { return t.indexOf(k) >= 0; }); };

  // Q01 trùng số hiệu + rev nhưng khác nội dung
  var g = {};
  src.details.forEach(function (d) { (g[d.type + '|' + d.key + '|' + d.rev] = g[d.type + '|' + d.key + '|' + d.rev] || []).push(d); });
  Object.keys(g).forEach(function (k) {
    var a = g[k]; if (a.length < 2) return;
    var subj = {}; a.forEach(function (d) { subj[tqnNoAccent_(d.subject)] = 1; });
    if (Object.keys(subj).length > 1)
      add('LỖI', 'Q01_TRUNG_MA', a.map(function (d) { return d.sheet + '!' + d.row; }).join(', '), a[0].code,
        'Một số hiệu dùng cho ' + a.length + ' hồ sơ khác nội dung: ' + a.map(function (d) { return '"' + d.subject + '"'; }).join(' / '));
  });

  src.details.forEach(function (d) {
    var at = d.sheet + '!' + d.row;
    // Q02 khoảng trắng / tab trong mã
    if (d.rawCode && d.rawCode !== d.code && d.rawCode.trim() !== d.code) add('CẢNH BÁO', 'Q02_KHOANG_TRANG', at, d.code, 'Số hiệu có khoảng trắng hoặc tab thừa: "' + d.rawCode + '"');
    else if (d.rawCode && d.rawCode !== d.rawCode.trim()) add('CẢNH BÁO', 'Q02_KHOANG_TRANG', at, d.code, 'Số hiệu có khoảng trắng đầu/cuối');
    // Q03 cấu trúc yymmdd-...-nnn
    var m = /^(\d{6})-(.+?)-(\d{3})/.exec(d.code);
    if (!m) add('CẢNH BÁO', 'Q03_CAU_TRUC', at, d.code, 'Không theo cấu trúc yymmdd-…-nnn (QĐ ghi mã COA mục II.2)');
    else {
      var y = 2000 + +m[1].slice(0, 2), mo = +m[1].slice(2, 4), da = +m[1].slice(4, 6), dt = new Date(y, mo - 1, da);
      if (dt.getMonth() !== mo - 1 || dt.getDate() !== da) add('CẢNH BÁO', 'Q03_CAU_TRUC', at, d.code, '6 số đầu không phải ngày hợp lệ');
      var nh = m[2].toUpperCase(), nd = d.subject.toUpperCase();
      // Q04 gói NP / HTKT không khớp nội dung
      if (nh.indexOf('HTKT') >= 0 && (nd.indexOf('REHAUS') >= 0 || /\bRE[45]\b/.test(nd))) add('CẢNH BÁO', 'Q04_GOI_MA', at, d.code, 'Nội dung nhà phố (Rehaus) nhưng mã thuộc gói HTKT');
      if (nh.indexOf('.NP') >= 0 && /\(LD-/.test(nd)) add('CẢNH BÁO', 'Q04_GOI_MA', at, d.code, 'Nội dung công viên (LD-) nhưng mã thuộc gói NP');
    }
    // Q11 ngày "Thực tế" trình nằm ở tương lai → thường là ngày kế hoạch nhập nhầm cột
    var future = tqnIsDate_(d.submitted) && d.submitted > today;
    if (future) add('CẢNH BÁO', 'Q11_NGAY_TUONG_LAI', at, d.code, 'Cột "Thực tế" (ngày trình) ghi ' + tqnFmtDate_(d.submitted, cfg.timezone) + ', sau hôm nay: có thể là ngày kế hoạch nhập nhầm cột Thực tế');
    // Q05 đã trình ở sổ chi tiết nhưng chưa có trong Sổ trình duyệt
    if (!d.log && (isSubmitted(d.status) || /da duyet|khong duyet/.test(tqnNoAccent_(d.status)) || (tqnIsDate_(d.submitted) && !future)))
      add('CẢNH BÁO', 'Q05_CHUA_VAO_SO_TONG', at, d.code, 'Sổ chi tiết ghi "' + (d.status || 'có ngày trình') + '" nhưng không tìm thấy dòng tương ứng trong ' + cfg.log_sheet);
    // Q07 lệch tình trạng giữa 2 sổ
    if (d.log) {
      var a = tqnNoAccent_(d.status), b = tqnNoAccent_(d.log.status || d.log.result);
      var done = /da duyet|khong duyet|da phan hoi/;
      if (b && done.test(b) && !done.test(a))
        add('CẢNH BÁO', 'Q07_LECH_TINH_TRANG', at + ' ↔ ' + d.log.sheet + '!' + d.log.row, d.code,
          'Sổ tổng: "' + (d.log.status || d.log.result) + '" nhưng sổ chi tiết vẫn: "' + (d.status || '(trống)') + '"');
    }
  });

  // Q06 dòng Sổ trình duyệt loại RFI/RFA/RFC/SD/MS/MAT/OTS/SCH nhưng không ghép được sổ chi tiết
  var typed = /^(RFI|RFA|RFC|SD|SDS|MS|MSS|MAT|MAS|OTS|SCH)$/i;
  src.logUnmatched.forEach(function (L) {
    if (typed.test(L.type)) add('THÔNG TIN', 'Q06_SO_TONG_CHUA_GHEP', L.sheet + '!' + L.row, L.wf,
      'Không ghép được với sổ chi tiết' + (L.tbc.length ? ' (mã trong file: ' + L.tbc.map(function (t) { return t.code; }).join(', ') + ')' : ' (cột Link hồ sơ / Ghi chú chưa có mã TBC)'));
  });

  // Q08 cùng một ý kiến dài lặp ở nhiều hồ sơ khác nhau (dấu hiệu chép nhầm)
  var notes = {};
  src.details.forEach(function (d) {
    [['TVGS', d.tvgsNote], ['CĐT', d.cdtNote], ['TVTK', d.tvtkNote]].forEach(function (p) {
      var t = String(p[1] || ''); if (t.length < 30) return;
      var k = tqnNoAccent_(t); (notes[k] = notes[k] || []).push({ d: d, who: p[0], t: t });
    });
  });
  Object.keys(notes).forEach(function (k) {
    var a = notes[k], keys = {}; a.forEach(function (x) { keys[x.d.key] = 1; });
    if (Object.keys(keys).length > 1)
      add('CẢNH BÁO', 'Q08_Y_KIEN_LAP', a.map(function (x) { return x.d.sheet + '!' + x.d.row; }).join(', '), a.map(function (x) { return x.d.code; }).join(' / '),
        'Cùng một ý kiến ' + a[0].who + ' ở ' + a.length + ' hồ sơ khác nhau, kiểm có chép nhầm: "' + a[0].t.slice(0, 120) + '"');
  });

  // Q09 công thức số ngày CĐT trừ nhầm cột "Số ngày" của TVGS
  tqnQCFormulas_(cfg, imp, add);

  // Q10 hồ sơ quá hạn theo Sổ trình duyệt (để nhắc)
  src.log.forEach(function (L) {
    if (tqnNoAccent_(L.status) === 'qua han') add('THÔNG TIN', 'Q10_QUA_HAN', L.sheet + '!' + L.row, L.wf, 'Quá hạn phản hồi từ ' + tqnFmtDate_(L.due, cfg.timezone) + ': ' + L.subject.slice(0, 90));
  });

  var sh = tqnOwnSheet_(TQN.SH_QC, TQN_QC_HEADERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, TQN_QC_HEADERS.length).clearContent();
  var order = { 'LỖI': 0, 'CẢNH BÁO': 1, 'THÔNG TIN': 2 };
  out.sort(function (a, b) { return order[a[0]] - order[b[0]] || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0); });
  if (out.length) sh.getRange(2, 1, out.length, TQN_QC_HEADERS.length).setValues(out);
  var c = {}; out.forEach(function (r) { c[r[1]] = (c[r[1]] || 0) + 1; });
  return { rows: out, byRule: c };
}

function tqnQCFormulas_(cfg, imp, add) {
  var ss = tqnSs_();
  Object.keys(imp.src.maps).forEach(function (n) {
    var m = imp.src.maps[n]; if (!m || !m.cdt || !m.tvgs || m.cdt.days < 0 || m.tvgs.days < 0) return;
    var sh = ss.getSheetByName(n), r0 = cfg.firstRow;
    if (!sh || sh.getLastRow() < r0) return;
    var f = sh.getRange(r0, m.cdt.days + 1, Math.min(5, sh.getLastRow() - r0 + 1), 1).getFormulas();
    var tvDaysLetter = tqnColLetter_(m.tvgs.days + 1);
    for (var i = 0; i < f.length; i++) {
      var s = String(f[i][0] || '');
      if (new RegExp('-\\s*' + tvDaysLetter + '\\d+\\b').test(s)) {
        add('LỖI', 'Q09_CONG_THUC', n + '!' + tqnColLetter_(m.cdt.days + 1) + (r0 + i) + ' (cả cột)', '',
          'Số ngày CĐT phản hồi = ngày CĐT trừ cột ' + tvDaysLetter + ' là "Số ngày" TVGS (một con số), kết quả sai. Công thức: ' + s);
        break;
      }
    }
  });
}
function tqnColLetter_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
