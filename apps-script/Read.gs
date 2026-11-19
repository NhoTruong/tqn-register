/**
 * Read.gs – đọc Register CHỈ ĐỌC. Dò cột theo tiêu đề, không cố định chữ cái cột.
 */

/** Tìm cột theo từ khóa (đã bỏ dấu) trong vùng tiêu đề. Trả về chỉ số cột 0-based hoặc -1. */
function tqnFindCol_(hdr, keys) {
  for (var k = 0; k < keys.length; k++) {
    var key = tqnNoAccent_(keys[k]);
    for (var r = 0; r < hdr.length; r++)
      for (var c = 0; c < hdr[r].length; c++) {
        var t = tqnNoAccent_(tqnClean_(hdr[r][c]));
        if (t && (t === key || t.indexOf(key) === 0)) return c;
      }
  }
  return -1;
}

/**
 * Khối nhiều cột (vd "Tư vấn giám sát phê duyệt"): bắt đầu ở cột có nhãn, kết thúc trước nhãn kế tiếp
 * cùng dòng. Trong khối tìm cột con "Thực tế", "Vướng mắc", "Kế hoạch"/"Ngày trình duyệt", "Số ngày".
 */
function tqnFindBlock_(hdr, keys) {
  for (var r = 0; r < hdr.length; r++) {
    var row = hdr[r];
    for (var c = 0; c < row.length; c++) {
      var t = tqnNoAccent_(tqnClean_(row[c]));
      var hit = keys.some(function (k) { return t && t.indexOf(tqnNoAccent_(k)) === 0; });
      if (!hit) continue;
      var end = c + 1;
      while (end < row.length && !tqnClean_(row[end])) end++;
      var sub = function (labels) {
        for (var rr = r; rr < hdr.length; rr++)
          for (var cc = c; cc < end; cc++) {
            var s = tqnNoAccent_(tqnClean_(hdr[rr][cc]));
            if (s && labels.some(function (l) { return s.indexOf(tqnNoAccent_(l)) === 0; }) && !(rr === r && cc === c && rr === r)) return cc;
          }
        return -1;
      };
      return {
        start: c, end: end - 1,
        actual: sub(['Thực tế']),          // "Thực tế" đầu tiên trong khối (không lấy "Thực tế email")
        note: sub(['Vướng mắc']),
        plan: sub(['Kế hoạch', 'Ngày trình duyệt']),
        days: sub(['Số ngày'])
      };
    }
  }
  return null;
}

/** Bản đồ cột của một sổ chi tiết. */
function tqnDetailMap_(hdr) {
  var m = {
    code: tqnFindCol_(hdr, ['Mã hồ sơ']),
    subject: tqnFindCol_(hdr, ['Nội dung chi tiết', 'Nội dung']),
    owner: tqnFindCol_(hdr, ['Người thực hiện']),
    status: tqnFindCol_(hdr, ['Tình trạng hồ sơ', 'Tình trạng']),
    rev: tqnFindCol_(hdr, ['Rev']),
    ret: tqnFindCol_(hdr, ['Ngày giao nhận']),
    nt: tqnFindBlock_(hdr, ['Nhà thầu đệ trình']),
    tvgs: tqnFindBlock_(hdr, ['Tư vấn giám sát']),
    cdt: tqnFindBlock_(hdr, ['Chủ Đầu Tư']),
    tvtk: tqnFindBlock_(hdr, ['Tư vấn thiết kế'])
  };
  if (m.code < 0) m.code = 1;            // mặc định cột B
  if (m.subject < 0) m.subject = 2;
  return m;
}

/** Ghép giá trị cột con: trả ngày nếu là ngày, ngược lại chữ. */
function tqnCell_(row, c) { return (c === undefined || c < 0 || c >= row.length) ? '' : row[c]; }
function tqnBlockVal_(row, b) {
  if (!b) return { date: '', note: '' };
  var a = tqnCell_(row, b.actual), n = tqnClean_(tqnCell_(row, b.note));
  var date = tqnIsDate_(a) ? a : '';
  var txt = (!date && tqnClean_(a)) ? String(tqnClean_(a)) : '';
  // Khối có ô đầu là chữ (vd M0S "Ghi trên hồ sơ") → đưa vào ghi chú, không bỏ
  var first = tqnCell_(row, b.start);
  if (b.start !== b.actual && first && !tqnIsDate_(first) && typeof first !== 'number') txt = [String(tqnClean_(first)), txt].filter(String).join(' | ');
  return { date: date, note: [txt, n].filter(String).join(' | ') };
}

/** Đọc một sổ chi tiết → danh sách bản ghi. Bỏ qua dòng mã "..." / dòng mục lớn / mã không có số. */
function tqnReadDetail_(sh, cfg) {
  var name = sh.getName(), lastR = sh.getLastRow(), lastC = Math.min(sh.getLastColumn(), 40);
  if (lastR < cfg.hdrFrom || lastC < 2) return { map: null, rows: [] };
  var hdrN = Math.min(cfg.hdrTo, lastR) - cfg.hdrFrom + 1;
  var hdr = sh.getRange(cfg.hdrFrom, 1, hdrN, lastC).getValues();
  var m = tqnDetailMap_(hdr);
  var first = cfg.firstRow;
  if (name === 'TRN-GNHS') first = Math.max(2, cfg.hdrFrom + 1);
  if (lastR < first) return { map: m, rows: [] };
  var vals = sh.getRange(first, 1, lastR - first + 1, lastC).getValues();
  var out = [];
  // Cột "Ngày trình duyệt" (RFI/RFA/RFC) là ngày trình THẬT, khác cột "Kế hoạch" (SD/M0S). Dò ở MỌI dòng tiêu đề.
  var ntp = (m.nt || {}).plan, planIsSubmit = ntp >= 0 && hdr.some(function (h) { return /trinh duyet/.test(tqnNoAccent_(tqnClean_(h[ntp]))); });
  vals.forEach(function (row, i) {
    var raw = row[m.code];
    var code = tqnCleanCode_(raw);
    if (!code || !/\d{6}/.test(code) || /^\.{2,}/.test(code)) return;
    var subj = tqnClean_(row[m.subject]);
    var nt = m.nt || {};
    var plan = tqnCell_(row, nt.plan), act = tqnCell_(row, nt.actual);
    var tv = tqnBlockVal_(row, m.tvgs), cd = tqnBlockVal_(row, m.cdt), tk = tqnBlockVal_(row, m.tvtk);
    out.push({
      sheet: name, row: first + i,
      rawCode: tqnIsDate_(raw) ? '' : String(raw === null || raw === undefined ? '' : raw),
      code: code, key: tqnDocKey_(code), rev: tqnRev_(tqnCell_(row, m.rev), code),
      subject: String(subj), owner: String(tqnClean_(tqnCell_(row, m.owner))),
      status: String(tqnClean_(tqnCell_(row, m.status))),
      plan: tqnIsDate_(plan) ? plan : '', submitted: tqnIsDate_(act) ? act : (tqnIsDate_(plan) && planIsSubmit ? plan : ''),
      tvgsDate: tv.date, tvgsNote: tv.note, cdtDate: cd.date, cdtNote: cd.note,
      tvtkDate: tk.date, tvtkNote: tk.note,
      ret: tqnIsDate_(tqnCell_(row, m.ret)) ? tqnCell_(row, m.ret) : ''
    });
  });
  return { map: m, rows: out, hdr: hdr };
}

/** Đọc Sổ trình duyệt (sổ tổng). Dòng tiêu đề = dòng có cả "Loại hồ sơ" và "Mã hồ sơ". */
function tqnReadLog_(sh) {
  var lastR = sh.getLastRow(), lastC = Math.min(sh.getLastColumn(), 30);
  var all = sh.getRange(1, 1, lastR, lastC).getValues();
  var h = -1;
  for (var r = 0; r < Math.min(all.length, 60); r++) {
    var t = all[r].map(function (x) { return tqnNoAccent_(tqnClean_(x)); });
    if (t.indexOf('loai ho so') >= 0 && t.indexOf('ma ho so') >= 0) { h = r; break; }
  }
  if (h < 0) throw new Error('Không thấy dòng tiêu đề trong ' + sh.getName());
  var H = all[h].map(function (x) { return tqnNoAccent_(tqnClean_(x)); });
  var col = function (k) { var key = tqnNoAccent_(k); for (var i = 0; i < H.length; i++) if (H[i] === key || H[i].indexOf(key) === 0) return i; return -1; };
  var c = {
    type: col('Loại hồ sơ'), wf: col('Mã hồ sơ'), rev: col('Rev'), subject: col('Nội dung chi tiết'), item: col('Hạng mục'),
    owner: col('Người thực hiện'), to: col('Trình đến'), date: col('Ngày trình'), days: col('Thời hạn phản hồi'),
    due: col('Hạn phản hồi'), reply: col('Ngày phản hồi'), result: col('Kết quả'), status: col('Tình trạng'),
    link: col('Link hồ sơ'), note: col('Ghi chú')
  };
  var out = [];
  for (var i = h + 1; i < all.length; i++) {
    var row = all[i], wf = tqnClean_(tqnCell_(row, c.wf));
    if (!wf) continue;
    var link = String(tqnClean_(tqnCell_(row, c.link)));
    var note = String(tqnClean_(tqnCell_(row, c.note)));
    out.push({
      sheet: sh.getName(), row: i + 1, type: String(tqnClean_(tqnCell_(row, c.type))), wf: String(wf),
      rev: tqnRev_(tqnCell_(row, c.rev), ''), subject: String(tqnClean_(tqnCell_(row, c.subject))),
      item: String(tqnClean_(tqnCell_(row, c.item))), owner: String(tqnClean_(tqnCell_(row, c.owner))),
      to: String(tqnClean_(tqnCell_(row, c.to))),
      date: tqnIsDate_(tqnCell_(row, c.date)) ? row[c.date] : '', due: tqnIsDate_(tqnCell_(row, c.due)) ? row[c.due] : '',
      reply: tqnIsDate_(tqnCell_(row, c.reply)) ? row[c.reply] : '',
      result: String(tqnClean_(tqnCell_(row, c.result))), status: String(tqnClean_(tqnCell_(row, c.status))),
      link: link, note: note, tbc: tqnExtractCodes_(link + ' ' + note)
    });
  }
  return out;
}

/** Lấy các mã kiểu TBC (yymmdd-COA...) trong chuỗi (tên file, ghi chú). */
function tqnExtractCodes_(s) {
  var re = /\d{6}\s*-\s*COA[\w.\-]*(?:[ _-]?rev ?\d{1,3})?/gi, m, out = [];
  var seen = {};
  while ((m = re.exec(String(s || ''))) !== null) {
    var t = { code: tqnCleanCode_(m[0]), key: tqnDocKey_(m[0]), rev: tqnRevFromNo_(m[0]) };
    if (!seen[t.key + '|' + t.rev]) { seen[t.key + '|' + t.rev] = 1; out.push(t); }   // 1 mã chỉ tính 1 lần / dòng
  }
  return out;
}
