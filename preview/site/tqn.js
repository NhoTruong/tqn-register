/**
 * TQN Register Universal – Apps Script S1
 * Config.gs: cấu hình theo dự án. Mọi tham số nằm ở tab TQN_CONFIG (cột A khóa, B giá trị),
 * không sửa code khi tạo dự án mới.
 */
var TQN = {
  VERSION: 'S1-0.1.0',
  SH_CONFIG: 'TQN_CONFIG',
  SH_DATA: 'TQN_DATA',
  SH_QC: 'TQN_QC',
  SH_REMIND: 'TQN_NHAC_HAN',
  SH_AUDIT: 'TQN_AUDIT',
  PREFIX: 'TQN_'           // chỉ được ghi vào tab có tiền tố này
};

/** Giá trị mặc định cho profile COASTAL_2026_V1. Ghi vào TQN_CONFIG lần đầu cài. */
var TQN_DEFAULT_CONFIG = [
  ['project_id', 'COASTAL_QNG', 'Mã dự án bất biến, không trùng giữa các dự án'],
  ['project_name', 'Khu đô thị sinh thái Coastal Quảng Ngãi', ''],
  ['code_profile', 'COASTAL_2026_V1', 'Bộ quy tắc mã hồ sơ'],
  ['timezone', 'Asia/Ho_Chi_Minh', 'Múi giờ IANA của công trường'],
  ['log_sheet', 'Sổ trình duyệt', 'Sổ tổng có Kết quả / Hạn / Tình trạng'],
  ['detail_sheets', 'SD,M0S,MAT,OTS,SCH,RFA,RFI,RFC,CRP- Báo cáo,Hồ sơ an toàn,TRN-GNHS', 'Sổ chi tiết kiểu trình duyệt (cột B = mã)'],
  ['type_alias', 'SD=SDS;M0S=MSS;MS=MSS;MAT=MAS;OTS=OTS;SCH=SCH;RFA=RFA;RFI=RFI;RFC=RFC;TRN-GNHS=TML;CRP- Báo cáo=REP;Hồ sơ an toàn=HSE', 'Tên tab → loại chuẩn. Tab gốc không đổi tên'],
  ['status_submitted', 'đã trình duyệt;đang chờ;gửi email', 'Chữ trạng thái (sổ chi tiết) coi như đã trình'],
  ['status_closed', 'đã duyệt;duyệt có ý kiến;không duyệt;đã phản hồi;hủy', 'Chữ trạng thái coi như đã có phản hồi'],
  ['remind_days_ahead', '2', 'Nhắc trước hạn bao nhiêu ngày'],
  ['remind_hour', '7', 'Giờ chạy nhắc hạn hằng ngày (giờ dự án). Apps Script có thể lệch ±15 phút'],
  ['remind_email_enabled', 'Không', 'Có = gửi mail; Không = chỉ ghi tab TQN_NHAC_HAN'],
  ['remind_recipients', '', 'Email nhận nhắc hạn, cách nhau dấu phẩy'],
  ['detail_first_row', '6', 'Dòng dữ liệu đầu tiên của sổ chi tiết'],
  ['header_rows', '1-5', 'Vùng dòng tiêu đề để dò cột']
];

function tqnSs_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Đọc cấu hình thành object; thiếu khóa thì dùng mặc định. */
function tqnConfig_() {
  var cfg = {};
  TQN_DEFAULT_CONFIG.forEach(function (r) { cfg[r[0]] = String(r[1]); });
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k) cfg[k] = String(r[1] === null || r[1] === undefined ? '' : r[1]).trim();
    });
  }
  cfg.detailSheets = tqnSplit_(cfg.detail_sheets, ',');
  cfg.alias = {};
  tqnSplit_(cfg.type_alias, ';').forEach(function (p) {
    var i = p.indexOf('=');
    if (i > 0) cfg.alias[p.slice(0, i).trim()] = p.slice(i + 1).trim();
  });
  cfg.statusSubmitted = tqnSplit_(cfg.status_submitted, ';').map(tqnLower_);
  cfg.statusClosed = tqnSplit_(cfg.status_closed, ';').map(tqnLower_);
  cfg.remindDays = parseInt(cfg.remind_days_ahead, 10) || 2;
  cfg.remindHour = parseInt(cfg.remind_hour, 10) || 7;
  cfg.firstRow = parseInt(cfg.detail_first_row, 10) || 6;
  var hr = String(cfg.header_rows).split('-');
  cfg.hdrFrom = parseInt(hr[0], 10) || 1; cfg.hdrTo = parseInt(hr[1], 10) || 5;
  cfg.emailOn = /^c|^y|^1/i.test(cfg.remind_email_enabled);
  return cfg;
}

/** Tạo tab TQN_CONFIG nếu chưa có (không ghi đè giá trị anh đã sửa). */
function tqnEnsureConfig_() {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN.SH_CONFIG);
  if (sh) return sh;
  sh = ss.insertSheet(TQN.SH_CONFIG);
  sh.getRange(1, 1, 1, 3).setValues([['Khóa', 'Giá trị', 'Ghi chú']]).setFontWeight('bold').setBackground('#1B3F7A').setFontColor('#ffffff');
  sh.getRange(2, 1, TQN_DEFAULT_CONFIG.length, 3).setValues(TQN_DEFAULT_CONFIG);
  sh.getRange(2, 2, TQN_DEFAULT_CONFIG.length, 1).setBackground('#FFF2CC');
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 380); sh.setColumnWidth(3, 460);
  sh.setFrozenRows(1);
  return sh;
}

/* ------------ tiện ích chung ------------ */
function tqnSplit_(s, sep) { return String(s || '').split(sep).map(function (x) { return x.trim(); }).filter(String); }
function tqnLower_(s) { return String(s || '').toLowerCase().trim(); }

/** Làm sạch chữ: tab, khoảng trắng cứng, xuống dòng, khoảng trắng kép, khoảng trắng quanh '-'. */
function tqnClean_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  return String(v).replace(/[\t \r\n]/g, ' ').replace(/ {2,}/g, ' ').trim();
}
function tqnCleanCode_(v) { return String(tqnClean_(v)).replace(/\s*-\s*/g, '-'); }

/** Khóa so khớp hồ sơ: bỏ tiền tố ngày yymmdd- và đuôi rev (giữ nguyên số hiệu gốc ở cột khác). */
function tqnDocKey_(no) {
  var s = tqnCleanCode_(no).replace(/\.pdf$/i, '');
  if (/^\d{6}-/.test(s)) s = s.slice(7);
  s = s.replace(/[-_ ]*rev ?\d{1,3}\b.*$/i, '');
  return s.toUpperCase().trim();
}
function tqnRevFromNo_(no) { var m = /rev ?(\d{1,3})\b/i.exec(String(no || '')); return m ? ('0' + parseInt(m[1], 10)).slice(-2) : ''; }
function tqnRev_(v, no) {
  var s = String(tqnClean_(v)).toUpperCase().replace(/^REV/, '').trim();
  if (!s) return tqnRevFromNo_(no);
  if (/^\d{1,3}(\.0+)?$/.test(s)) return ('0' + parseInt(s, 10)).slice(-2);
  return s;
}
function tqnIsDate_(v) { return v instanceof Date && !isNaN(v.getTime()); }
function tqnFmtDate_(d, tz) { return tqnIsDate_(d) ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : ''; }

/** Chuẩn hóa chữ để so (bỏ dấu tiếng Việt). */
function tqnNoAccent_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
}

/** Ghi nhật ký TQN_AUDIT (append-only). */
function tqnAudit_(action, detail, counts) {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN.SH_AUDIT);
  if (!sh) {
    sh = ss.insertSheet(TQN.SH_AUDIT);
    sh.appendRow(['Thời điểm', 'Người chạy', 'Việc', 'Chi tiết', 'Số liệu', 'Phiên bản']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { user = ''; }
  sh.appendRow([new Date(), user, action, detail, JSON.stringify(counts || {}), TQN.VERSION]);
}

/** Chặn mọi lệnh ghi vào tab gốc: chỉ trả về tab có tiền tố TQN_. */
function tqnOwnSheet_(name, headers) {
  if (name.indexOf(TQN.PREFIX) !== 0) throw new Error('TQN không được ghi vào tab gốc: ' + name);
  var ss = tqnSs_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1B3F7A').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

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
      plan: tqnIsDate_(plan) ? plan : '', submitted: tqnIsDate_(act) ? act : (tqnIsDate_(plan) && nt.plan >= 0 && /trinh duyet/.test(tqnNoAccent_(tqnClean_(hdr[hdr.length > 1 ? 1 : 0][nt.plan]))) ? plan : ''),
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

/**
 * Import.gs – nhập không trùng (idempotent) từ Sổ trình duyệt + sổ chi tiết vào TQN_DATA.
 * Không bao giờ ghi vào tab gốc. Chạy lại bao nhiêu lần cũng ra cùng kết quả.
 */
var TQN_DATA_HEADERS = [
  'Khóa bản ghi', 'Dự án', 'Nguồn', 'Loại', 'Số hiệu (nguyên văn)', 'Khóa so khớp', 'Rev', 'Nội dung',
  'Người thực hiện', 'Tình trạng (sổ chi tiết)', 'Kế hoạch trình', 'Ngày trình',
  'TVGS ngày', 'TVGS ý kiến', 'TVTK ngày', 'TVTK ý kiến', 'CĐT ngày', 'CĐT ý kiến', 'Ngày giao nhận trả về',
  'Mã WF (Sổ trình duyệt)', 'Trình đến', 'Hạn phản hồi', 'Ngày phản hồi', 'Kết quả', 'Tình trạng (Sổ trình duyệt)',
  'Link hồ sơ', 'Ghép WF↔TBC', 'Nhập lần đầu', 'Cập nhật', 'Còn trong nguồn'
];
var TQN_COL = {}; TQN_DATA_HEADERS.forEach(function (h, i) { TQN_COL[h] = i; });

/** Băm ngắn để phân biệt 2 hồ sơ trùng số hiệu nhưng khác nội dung. */
function tqnHash8_(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, tqnNoAccent_(s), Utilities.Charset.UTF_8);
  return b.slice(0, 4).map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}

/** Đọc toàn bộ nguồn (chỉ đọc). Trả về {log, details, maps, missingSheets}. */
function tqnReadAll_(cfg) {
  var ss = tqnSs_(), res = { log: [], details: [], maps: {}, missingSheets: [], hdrs: {} };
  var logSh = ss.getSheetByName(cfg.log_sheet);
  if (logSh) res.log = tqnReadLog_(logSh); else res.missingSheets.push(cfg.log_sheet);
  cfg.detailSheets.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) { res.missingSheets.push(n); return; }
    var r = tqnReadDetail_(sh, cfg);
    res.maps[n] = r.map; res.hdrs[n] = r.hdr;
    r.rows.forEach(function (x) { x.type = cfg.alias[n] || n; res.details.push(x); });
  });
  return res;
}

/** Ghép dòng chi tiết với dòng Sổ trình duyệt qua mã TBC nằm trong tên file / ghi chú. */
function tqnMatch_(src) {
  var byKeyRev = {}, byKey = {};
  src.log.forEach(function (L) {
    L.tbc.forEach(function (t) {
      (byKeyRev[t.key + '|' + t.rev] = byKeyRev[t.key + '|' + t.rev] || []).push(L);
      (byKey[t.key] = byKey[t.key] || []).push(L);
    });
  });
  var used = {};
  src.details.forEach(function (d) {
    var hit = byKeyRev[d.key + '|' + d.rev] || [];
    var how = 'KEY+REV';
    if (!hit.length && byKey[d.key]) { hit = byKey[d.key]; how = 'KEY'; }
    if (hit.length) {
      // nhiều dòng WF cho cùng mã: lấy dòng trình gần nhất
      hit = hit.slice().sort(function (a, b) { return (b.date || 0) - (a.date || 0); });
      d.log = hit[0]; d.matchHow = how + (hit.length > 1 ? ' (' + hit.length + ' dòng WF)' : '');
      used[hit[0].row] = true;
    }
  });
  src.logUnmatched = src.log.filter(function (L) { return !used[L.row]; });
  return src;
}

/** Chạy nhập. Trả về số liệu. */
function tqnImport_(cfg) {
  var src = tqnMatch_(tqnReadAll_(cfg));
  var sh = tqnOwnSheet_(TQN.SH_DATA, TQN_DATA_HEADERS);
  var now = new Date(), old = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TQN_DATA_HEADERS.length).getValues().forEach(function (r) { old[r[0]] = r; });
  }
  var rows = [], seen = {};
  var put = function (key, r) {
    if (seen[key]) { seen[key]++; key = key + '#' + seen[key]; } else seen[key] = 1;   // không bao giờ bỏ im lặng
    r[0] = key; r[1] = cfg.project_id;
    var prev = old[key];
    r[TQN_COL['Nhập lần đầu']] = prev && prev[TQN_COL['Nhập lần đầu']] ? prev[TQN_COL['Nhập lần đầu']] : now;
    r[TQN_COL['Cập nhật']] = now; r[TQN_COL['Còn trong nguồn']] = 'Có';
    rows.push(r); delete old[key];
  };
  var blank = function () { return TQN_DATA_HEADERS.map(function () { return ''; }); };
  src.details.forEach(function (d) {
    var r = blank(), L = d.log || {};
    r[2] = d.sheet + '!' + d.row; r[3] = d.type; r[4] = d.rawCode.trim() ? String(d.rawCode).trim() : d.code; r[5] = d.key; r[6] = d.rev;
    r[7] = d.subject; r[8] = d.owner; r[9] = d.status; r[10] = d.plan; r[11] = d.submitted;
    r[12] = d.tvgsDate; r[13] = d.tvgsNote; r[14] = d.tvtkDate; r[15] = d.tvtkNote; r[16] = d.cdtDate; r[17] = d.cdtNote; r[18] = d.ret;
    r[19] = L.wf || ''; r[20] = L.to || ''; r[21] = L.due || ''; r[22] = L.reply || ''; r[23] = L.result || ''; r[24] = L.status || '';
    r[25] = L.link || ''; r[26] = d.matchHow || '';
    put(cfg.project_id + '|DET|' + d.sheet + '|' + d.code.toUpperCase() + '|' + d.rev + '|' + tqnHash8_(d.subject), r);
  });
  src.logUnmatched.forEach(function (L) {
    var r = blank();
    r[2] = L.sheet + '!' + L.row; r[3] = L.type; r[4] = L.wf; r[5] = L.tbc.length ? L.tbc[0].key : ''; r[6] = L.rev; r[7] = L.subject;
    r[8] = L.owner; r[11] = L.date; r[19] = L.wf; r[20] = L.to; r[21] = L.due; r[22] = L.reply; r[23] = L.result; r[24] = L.status;
    r[25] = L.link; r[26] = 'Chỉ có trong Sổ trình duyệt';
    put(cfg.project_id + '|LOG|' + L.wf.toUpperCase() + '|' + L.rev, r);
  });
  // bản ghi cũ không còn trong nguồn: giữ lại, đánh dấu (không xóa)
  Object.keys(old).forEach(function (k) {
    var r = old[k]; r[TQN_COL['Còn trong nguồn']] = 'Không'; rows.push(r);
  });
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, TQN_DATA_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, TQN_DATA_HEADERS.length).setValues(rows);
  tqnFormatData_(sh, rows.length);
  return {
    src: src, rows: rows,
    counts: {
      chi_tiet: src.details.length, so_tong: src.log.length,
      ghep_duoc: src.details.filter(function (d) { return d.log; }).length,
      chi_co_so_tong: src.logUnmatched.length, khong_con_nguon: Object.keys(old).length,
      tong_dong: rows.length, tab_thieu: src.missingSheets.join(',')
    }
  };
}

function tqnFormatData_(sh, n) {
  if (!n || !sh.setFrozenColumns) return;
  sh.setFrozenColumns(6);
  [10, 11, 12, 14, 16, 18, 21, 22, 27, 28].forEach(function (c) { sh.getRange(2, c + 1, n, 1).setNumberFormat('dd/MM/yyyy'); });
  sh.getRange(2, 7, n, 1).setNumberFormat('@');
}

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

/**
 * Remind.gs – nhắc hạn hằng ngày. Nguồn hạn: Sổ trình duyệt (cột Hạn phản hồi / Tình trạng).
 * Mặc định chỉ ghi tab TQN_NHAC_HAN; gửi mail khi TQN_CONFIG remind_email_enabled = Có.
 */
var TQN_REMIND_HEADERS = ['Nhóm', 'Hạn phản hồi', 'Còn / quá (ngày)', 'Mã WF', 'Số hiệu TBC', 'Loại', 'Nội dung', 'Trình đến', 'Người thực hiện', 'Vị trí'];

/** Danh sách cần nhắc: quá hạn + đến hạn trong N ngày, chưa có Kết quả. */
function tqnRemindList_(cfg, imp, today) {
  today = today || new Date();
  var t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var lim = new Date(t0.getTime() + cfg.remindDays * 86400000);
  var tbcByWf = {};
  imp.src.details.forEach(function (d) { if (d.log) tbcByWf[d.log.row] = d.code; });
  var items = [];
  imp.src.log.forEach(function (L) {
    if (L.result) return;                                  // đã có kết quả thì thôi
    var st = tqnNoAccent_(L.status);
    if (!tqnIsDate_(L.due) && st !== 'qua han') return;    // không có hạn thì không gắn quá hạn
    var due = tqnIsDate_(L.due) ? new Date(L.due.getFullYear(), L.due.getMonth(), L.due.getDate()) : null;
    if (due && due > lim) return;
    var diff = due ? Math.round((due - t0) / 86400000) : '';
    var grp = (due && diff < 0) || st === 'qua han' ? 'QUÁ HẠN' : (diff === 0 ? 'HẠN HÔM NAY' : 'SẮP ĐẾN HẠN');
    items.push([grp, due || '', diff, L.wf, tbcByWf[L.row] || (L.tbc[0] ? L.tbc[0].code : ''), L.type, L.subject, L.to, L.owner, L.sheet + '!' + L.row]);
  });
  var ord = { 'QUÁ HẠN': 0, 'HẠN HÔM NAY': 1, 'SẮP ĐẾN HẠN': 2 };
  items.sort(function (a, b) { return ord[a[0]] - ord[b[0]] || ((a[1] || 0) - (b[1] || 0)); });
  return items;
}

function tqnWriteRemind_(cfg, items) {
  var sh = tqnOwnSheet_(TQN.SH_REMIND, TQN_REMIND_HEADERS);
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, TQN_REMIND_HEADERS.length).clearContent();
  if (items.length) {
    sh.getRange(2, 1, items.length, TQN_REMIND_HEADERS.length).setValues(items);
    sh.getRange(2, 2, items.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  return sh;
}

function tqnRemindHtml_(cfg, items, today) {
  var esc = function (s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var color = { 'QUÁ HẠN': '#A32D2D', 'HẠN HÔM NAY': '#854F0B', 'SẮP ĐẾN HẠN': '#185FA5' };
  var rows = items.map(function (r) {
    return '<tr><td style="color:' + color[r[0]] + ';font-weight:600">' + esc(r[0]) + '</td><td>' + esc(tqnFmtDate_(r[1], cfg.timezone)) + '</td><td>' + esc(r[2]) +
      '</td><td>' + esc(r[4] || r[3]) + '</td><td>' + esc(r[6]) + '</td><td>' + esc(r[7]) + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;font-size:13px"><h3 style="color:#1B3F7A;margin:0 0 6px">TQN Register · ' + esc(cfg.project_name) + '</h3>' +
    '<p>Nhắc hạn ngày ' + esc(tqnFmtDate_(today, cfg.timezone)) + ': ' + items.length + ' hồ sơ.</p>' +
    '<table cellpadding="5" style="border-collapse:collapse" border="1"><tr style="background:#EEF3FB"><th>Nhóm</th><th>Hạn</th><th>Ngày</th><th>Số hiệu</th><th>Nội dung</th><th>Trình đến</th></tr>' +
    rows + '</table><p style="color:#888">Nguồn: ' + esc(cfg.log_sheet) + '. Tự động bởi TQN Register ' + TQN.VERSION + '.</p></div>';
}

/** Hàm cho trigger hằng ngày. Chống chạy trùng trong cùng ngày bằng Script Properties. */
function tqnDailyJob() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) { tqnAudit_('NHAC_HAN', 'Bỏ qua: đang có lần chạy khác giữ khóa', {}); return; }
  try {
    var cfg = tqnConfig_(), today = new Date();
    var dayKey = cfg.project_id + ':' + Utilities.formatDate(today, cfg.timezone, 'yyyy-MM-dd') + ':remind';
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('lastRemind') === dayKey) { tqnAudit_('NHAC_HAN', 'Bỏ qua: hôm nay đã chạy ' + dayKey, {}); return; }
    var imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), items = tqnRemindList_(cfg, imp, today);
    tqnWriteRemind_(cfg, items);
    var sent = 0;
    if (cfg.emailOn && cfg.remind_recipients && items.length) {
      MailApp.sendEmail({
        to: cfg.remind_recipients,
        subject: '[TQN] Nhắc hạn ' + cfg.project_id + ' ' + Utilities.formatDate(today, cfg.timezone, 'dd/MM') + ': ' + items.length + ' hồ sơ',
        htmlBody: tqnRemindHtml_(cfg, items, today)
      });
      sent = 1;
    }
    props.setProperty('lastRemind', dayKey);
    tqnAudit_('NHAC_HAN', dayKey, { nhac: items.length, gui_mail: sent, qc: qc.rows.length });
  } finally { lock.releaseLock(); }
}

/** Bật trigger hằng ngày (xóa trigger cũ cùng tên trước, nên bấm nhiều lần không sinh trùng). */
function tqnInstallDaily() {
  var cfg = tqnConfig_();
  tqnRemoveDaily(true);
  ScriptApp.newTrigger('tqnDailyJob').timeBased().atHour(cfg.remindHour).everyDays(1).inTimezone(cfg.timezone).create();
  tqnAudit_('BAT_NHAC_HAN', 'Hằng ngày khoảng ' + cfg.remindHour + 'h (' + cfg.timezone + ')', {});
  return cfg.remindHour;
}
function tqnRemoveDaily(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnDailyJob') { ScriptApp.deleteTrigger(t); n++; } });
  if (!silent) tqnAudit_('TAT_NHAC_HAN', 'Đã xóa ' + n + ' trigger', {});
  return n;
}

/**
 * Weather.gs – 06:00 mỗi ngày ghi 1 dòng dự báo vào tab "Theo dõi thời tiết" của Register.
 * - Chỉ chạy khi: weather_enabled = Có, có GPS đã xác nhận, đã xác nhận điều khoản nguồn dữ liệu.
 * - Không ghi đè dòng người nhập tay; chỉ cập nhật dòng do TQN ghi (Ghi chú bắt đầu "[TQN]").
 * - Lỗi API → không ghi số giả, ghi nhật ký FAILED.
 */
var TQN_WEATHER_DEFAULTS = [
  ['weather_enabled', 'Không', 'Có = bật cập nhật thời tiết 06:00'],
  ['weather_lat', '15.099038', 'Vĩ độ công trường (vd 15.0xxxx). Phải là tọa độ thật đã xác nhận'],
  ['weather_lon', '108.894552', 'Kinh độ công trường (vd 108.8xxxx)'],
  ['weather_gps_confirmed', 'Có', 'Có = người quản lý đã xác nhận tọa độ đúng công trường (anh Nho gửi 11:12 19/11/2026)'],
  ['weather_sheet', 'Theo dõi thời tiết', 'Tab ghi thời tiết (giữ đúng format cột của Register)'],
  ['weather_hour', '6', 'Giờ chạy (giờ dự án). Apps Script có thể lệch ±15 phút'],
  ['weather_provider', 'OPEN_METEO', 'Nguồn dữ liệu'],
  ['weather_license_ok', 'Không', 'Có = đã kiểm điều khoản nguồn (Open-Meteo free chỉ cho phi thương mại; dùng cho công ty cần gói trả phí)'],
  ['weather_api_key', '', 'Khóa gói trả phí nếu có (để trống với gói miễn phí). Nên đặt ở Script Properties thay vì ô này']
];

var TQN_WMO_VI = {
  0: 'Trời quang', 1: 'Ít mây', 2: 'Có mây', 3: 'Nhiều mây', 45: 'Sương mù', 48: 'Sương mù đóng băng',
  51: 'Mưa phùn nhẹ', 53: 'Mưa phùn', 55: 'Mưa phùn dày', 56: 'Mưa phùn lạnh', 57: 'Mưa phùn lạnh dày',
  61: 'Mưa nhẹ', 63: 'Mưa vừa', 65: 'Mưa to', 66: 'Mưa lạnh nhẹ', 67: 'Mưa lạnh to',
  71: 'Tuyết nhẹ', 73: 'Tuyết vừa', 75: 'Tuyết dày', 77: 'Hạt tuyết',
  80: 'Mưa rào nhẹ', 81: 'Mưa rào', 82: 'Mưa rào rất to', 85: 'Mưa tuyết nhẹ', 86: 'Mưa tuyết to',
  95: 'Dông', 96: 'Dông kèm mưa đá nhẹ', 99: 'Dông kèm mưa đá'
};
var TQN_THU = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];

function tqnWeatherCfg_() {
  var cfg = tqnConfig_();
  TQN_WEATHER_DEFAULTS.forEach(function (r) { if (cfg[r[0]] === undefined) cfg[r[0]] = r[1]; });
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG);
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (String(r[0]).indexOf('weather_') === 0) cfg[String(r[0]).trim()] = String(r[1]).trim(); });
  var yes = function (v) { return /^c|^y|^1/i.test(String(v || '')); };
  cfg.wOn = yes(cfg.weather_enabled); cfg.wGps = yes(cfg.weather_gps_confirmed); cfg.wLic = yes(cfg.weather_license_ok);
  cfg.lat = parseFloat(String(cfg.weather_lat).replace(',', '.')); cfg.lon = parseFloat(String(cfg.weather_lon).replace(',', '.'));
  cfg.wHour = parseInt(cfg.weather_hour, 10) || 6;
  return cfg;
}

/** Thêm các khóa weather_* vào TQN_CONFIG nếu chưa có (không đụng giá trị anh đã sửa). */
function tqnEnsureWeatherConfig_() {
  var sh = tqnEnsureConfig_(), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_WEATHER_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}

/** Kiểm điều kiện trước khi gọi API. Trả về chuỗi lỗi hoặc ''. */
function tqnWeatherBlocked_(cfg) {
  if (!cfg.wOn) return 'weather_enabled chưa bật (Có)';
  if (isNaN(cfg.lat) || isNaN(cfg.lon) || Math.abs(cfg.lat) > 90 || Math.abs(cfg.lon) > 180) return 'Chưa có tọa độ hợp lệ (weather_lat / weather_lon)';
  if (!cfg.wGps) return 'Tọa độ chưa được xác nhận (weather_gps_confirmed = Có)';
  if (!cfg.wLic) return 'Chưa xác nhận điều khoản nguồn dữ liệu (weather_license_ok = Có)';
  return '';
}

function tqnWeatherUrl_(cfg) {
  var base = cfg.weather_api_key ? 'https://customer-api.open-meteo.com/v1/forecast' : 'https://api.open-meteo.com/v1/forecast';
  var q = {
    latitude: cfg.lat, longitude: cfg.lon, timezone: cfg.timezone, forecast_days: 1, wind_speed_unit: 'kmh',
    daily: 'weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,wind_gusts_10m_max',
    hourly: 'relative_humidity_2m'
  };
  if (cfg.weather_api_key) q.apikey = cfg.weather_api_key;
  return base + '?' + Object.keys(q).map(function (k) { return k + '=' + encodeURIComponent(q[k]); }).join('&');
}

/** Gọi API có thử lại; trả về {ok, data|err, raw, fetchedAt}. */
function tqnWeatherFetch_(cfg) {
  var url = tqnWeatherUrl_(cfg), last = '';
  for (var i = 0; i < 3; i++) {
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = res.getResponseCode(), raw = res.getContentText();
      if (code === 200) {
        var j = JSON.parse(raw), d = j.daily || {};
        var need = ['time', 'temperature_2m_min', 'temperature_2m_max', 'weather_code'];
        for (var n = 0; n < need.length; n++) if (!d[need[n]] || d[need[n]][0] === undefined || d[need[n]][0] === null) throw new Error('Thiếu trường ' + need[n]);
        var hum = (j.hourly && j.hourly.relative_humidity_2m || []).filter(function (x) { return typeof x === 'number'; });
        return {
          ok: true, raw: raw, fetchedAt: new Date(), data: {
            date: d.time[0], tmin: d.temperature_2m_min[0], tmax: d.temperature_2m_max[0], code: d.weather_code[0],
            rain: d.precipitation_sum ? d.precipitation_sum[0] : null, pop: d.precipitation_probability_max ? d.precipitation_probability_max[0] : null,
            wind: d.wind_speed_10m_max ? d.wind_speed_10m_max[0] : null, gust: d.wind_gusts_10m_max ? d.wind_gusts_10m_max[0] : null,
            hum: hum.length ? Math.round(hum.reduce(function (a, b) { return a + b; }, 0) / hum.length) : null
          }
        };
      }
      last = 'HTTP ' + code;
      if (code >= 400 && code < 500 && code !== 429) break;   // lỗi tham số: thử lại vô ích
    } catch (e) { last = String(e.message || e); }
    Utilities.sleep(Math.pow(2, i) * 2000);
  }
  return { ok: false, err: last };
}

/** Tab thời tiết: chỉ cho phép đúng tên trong cấu hình và đúng tiêu đề, để không ghi nhầm tab khác. */
function tqnWeatherSheet_(cfg) {
  var sh = tqnSs_().getSheetByName(cfg.weather_sheet);
  if (!sh) throw new Error('Không thấy tab ' + cfg.weather_sheet);
  var hdr = sh.getRange(1, 1, 1, Math.max(11, sh.getLastColumn())).getValues()[0].map(function (x) { return tqnNoAccent_(tqnClean_(x)); });
  var col = function (k) { var key = tqnNoAccent_(k); for (var i = 0; i < hdr.length; i++) if (hdr[i].indexOf(key) === 0) return i + 1; return 0; };
  var m = { date: col('Ngày'), thu: col('Thứ'), tmin: col('Nhiệt độ thấp nhất'), tmax: col('Nhiệt độ cao nhất'), wx: col('Thời tiết'), rain: col('Lượng mưa'), hum: col('Độ ẩm'), wind: col('Tốc độ gió'), note: col('Ghi chú') };
  if (!m.date || !m.tmin || !m.tmax || !m.note) throw new Error('Tiêu đề tab ' + cfg.weather_sheet + ' không đúng format (cần Ngày / Nhiệt độ thấp nhất / Nhiệt độ cao nhất / Ghi chú)');
  return { sh: sh, m: m };
}

/** Ghi / cập nhật dòng của ngày. Trả 'THEM' | 'CAP_NHAT' | 'BO_QUA_NHAP_TAY'. */
function tqnWeatherWrite_(cfg, w) {
  var t = tqnWeatherSheet_(cfg), sh = t.sh, m = t.m;
  var p = w.data.date.split('-'), day = new Date(+p[0], +p[1] - 1, +p[2]);
  var last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var dates = sh.getRange(2, m.date, last - 1, 1).getValues(), notes = sh.getRange(2, m.note, last - 1, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      var dv = dates[i][0];
      if (tqnIsDate_(dv) && dv.getFullYear() === day.getFullYear() && dv.getMonth() === day.getMonth() && dv.getDate() === day.getDate()) {
        if (String(notes[i][0]).indexOf('[TQN]') !== 0) return 'BO_QUA_NHAP_TAY';
        row = i + 2; break;
      }
    }
  }
  var action = row ? 'CAP_NHAT' : 'THEM';
  if (!row) row = Math.max(2, last + 1);
  var r1 = function (x) { return (x === null || x === undefined) ? '' : Math.round(x * 10) / 10; };
  var d = w.data;
  var note = '[TQN] Dự báo lúc ' + Utilities.formatDate(w.fetchedAt, cfg.timezone, 'HH:mm dd/MM') +
    (d.pop !== null ? ' · khả năng mưa ' + d.pop + '%' : '') + (d.gust !== null ? ' · gió giật ' + r1(d.gust) + ' km/h' : '') +
    ' · Nguồn: Open-Meteo (CC BY 4.0). Số dự báo, không phải quan trắc tại công trường';
  var set = function (c, v) { if (c) sh.getRange(row, c).setValue(v); };
  set(m.date, day); set(m.thu, TQN_THU[day.getDay()]); set(m.tmin, r1(d.tmin)); set(m.tmax, r1(d.tmax));
  set(m.wx, TQN_WMO_VI[d.code] || ('Mã WMO ' + d.code)); set(m.rain, r1(d.rain)); set(m.hum, d.hum === null ? '' : d.hum);
  set(m.wind, r1(d.wind)); set(m.note, note);
  if (m.date) sh.getRange(row, m.date).setNumberFormat('dd/MM/yyyy');
  return action;
}

/** Lưu bản gốc phản hồi API làm bằng chứng (tab riêng TQN_). */
function tqnWeatherRaw_(cfg, w, status) {
  var sh = tqnOwnSheet_('TQN_THOI_TIET_RAW', ['Lúc lấy', 'Dự án', 'Ngày dự báo', 'Trạng thái', 'Vĩ độ', 'Kinh độ', 'Phản hồi gốc (JSON)']);
  sh.appendRow([new Date(), cfg.project_id, w.ok ? w.data.date : '', status, cfg.lat, cfg.lon, w.ok ? String(w.raw).slice(0, 45000) : String(w.err)]);
}

/** Hàm cho trigger 06:00. manual=true khi bấm tay (vẫn chống trùng: cập nhật lại dòng TQN của hôm nay). */
function tqnWeatherJob(manual) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) { tqnAudit_('THOI_TIET', 'Bỏ qua: đang có lần chạy khác', {}); return 'BAN'; }
  try {
    var cfg = tqnWeatherCfg_(), why = tqnWeatherBlocked_(cfg);
    if (why) { tqnAudit_('THOI_TIET', 'Chặn: ' + why, {}); return 'CHAN: ' + why; }
    var key = cfg.project_id + ':' + Utilities.formatDate(new Date(), cfg.timezone, 'yyyy-MM-dd') + ':weather06';
    var props = PropertiesService.getDocumentProperties();
    if (manual !== true && props.getProperty('lastWeather') === key) { tqnAudit_('THOI_TIET', 'Bỏ qua: hôm nay đã chạy ' + key, {}); return 'DA_CHAY'; }
    var w = tqnWeatherFetch_(cfg);
    if (!w.ok) { tqnWeatherRaw_(cfg, w, 'FAILED'); tqnAudit_('THOI_TIET', 'FAILED: ' + w.err, {}); return 'FAILED: ' + w.err; }
    var act = tqnWeatherWrite_(cfg, w);
    tqnWeatherRaw_(cfg, w, act);
    props.setProperty('lastWeather', key);
    tqnAudit_('THOI_TIET', act + ' ' + w.data.date, w.data);
    return act;
  } finally { lock.releaseLock(); }
}

function tqnInstallWeather() {
  var cfg = tqnWeatherCfg_();
  tqnRemoveWeather(true);
  ScriptApp.newTrigger('tqnWeatherJob').timeBased().atHour(cfg.wHour).nearMinute(0).everyDays(1).inTimezone(cfg.timezone).create();
  tqnAudit_('BAT_THOI_TIET', 'Hằng ngày khoảng ' + cfg.wHour + ':00 (' + cfg.timezone + ')', {});
  return cfg.wHour;
}
function tqnRemoveWeather(silent) {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnWeatherJob') { ScriptApp.deleteTrigger(t); n++; } });
  if (!silent) tqnAudit_('TAT_THOI_TIET', 'Đã xóa ' + n + ' trigger', {});
  return n;
}

/* menu */
function tqnWeatherMsg_(r) {
  r = String(r || '');
  if (/^CHAN: /.test(r)) return 'Chưa cập nhật được thời tiết.\nLý do: ' + r.slice(6) + '\n\nCách bật: mở tab TQN_CONFIG, đổi weather_enabled = Có và weather_license_ok = Có (sau khi chốt nguồn dữ liệu), rồi bấm lại.';
  if (r === 'THEM' || r === 'CAP_NHAT') return '✔ Đã ' + (r === 'THEM' ? 'thêm' : 'cập nhật') + ' dòng thời tiết hôm nay vào tab Theo dõi thời tiết.';
  if (r === 'BO_QUA_NHAP_TAY') return 'Dòng hôm nay đã có người nhập tay – giữ nguyên, không ghi đè.';
  if (r === 'DA_CHAY') return 'Hôm nay đã cập nhật thời tiết rồi.';
  if (/^FAILED/.test(r)) return 'Lấy dữ liệu thời tiết lỗi: ' + r.replace(/^FAILED:\s*/, '') + '\nĐã thử lại 3 lần. Bấm lại sau ít phút.';
  return r;
}
function tqnMenuWeatherNow() { tqnEnsureWeatherConfig_(); SpreadsheetApp.getUi().alert('Thời tiết', tqnWeatherMsg_(tqnWeatherJob(true)), SpreadsheetApp.getUi().ButtonSet.OK); }
function tqnMenuWeatherOn() {
  tqnEnsureWeatherConfig_();
  var why = tqnWeatherBlocked_(tqnWeatherCfg_());
  if (why) { SpreadsheetApp.getUi().alert('Chưa bật được: ' + why + '\nSửa ở tab TQN_CONFIG.'); return; }
  SpreadsheetApp.getUi().alert('Đã bật cập nhật thời tiết hằng ngày khoảng ' + tqnInstallWeather() + ':00 (Google có thể lệch ±15 phút).');
}
function tqnMenuWeatherOff() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveWeather(false) + ' lịch thời tiết.'); }

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

/**
 * Ai.gs – TRỢ LÝ AI ĐỌC NỘI DUNG hồ sơ (S2).
 * Đọc nội dung + ý kiến TVGS/TVTK/CĐT của 1 hồ sơ → tóm tắt, mức độ, bên phải xử lý, việc tiếp theo,
 * danh sách cần bổ sung, nháp trả lời. KẾT QUẢ CHỈ LÀ BẢN NHÁP: ghi vào tab TQN_AI, không ghi vào sổ gốc,
 * người phụ trách đánh dấu "Đã duyệt" sau khi kiểm.
 * Khóa API để ở Script Properties (ANTHROPIC_API_KEY) – không để trong ô bảng tính, không gửi qua chat.
 */
var TQN_AI_DEFAULTS = [
  ['ai_enabled', 'Không', 'Có = bật trợ lý AI đọc nội dung (cần khóa ANTHROPIC_API_KEY trong Script Properties)'],
  ['ai_model', 'claude-sonnet-5-5', 'Mô hình AI'],
  ['ai_batch_max', '10', 'Số hồ sơ tối đa mỗi lần bấm "AI rà hồ sơ cần xử lý"']
];
var TQN_AI_HEADERS = ['Thời điểm', 'Vị trí', 'Số hiệu', 'Loại', 'Mức độ', 'Bên cần xử lý', 'Tóm tắt', 'Việc tiếp theo', 'Cần bổ sung',
  'Nháp trả lời', 'Mô hình', 'Đã duyệt bởi (người ghi)', 'Ghi chú'];
var TQN_AI_LEVELS = ['CAO', 'TRUNG_BINH', 'THAP'];

function tqnEnsureAiConfig_() {
  tqnEnsureConfig_();
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_AI_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}
function tqnAiCfg_() {
  var cfg = tqnConfig_();
  TQN_AI_DEFAULTS.forEach(function (r) { if (cfg[r[0]] === undefined || cfg[r[0]] === '') cfg[r[0]] = r[1]; });
  cfg.aiOn = /^c|^y|^1/i.test(cfg.ai_enabled);
  cfg.aiMax = Math.max(1, Math.min(30, parseInt(cfg.ai_batch_max, 10) || 10));
  return cfg;
}
function tqnAiBlocked_(cfg) {
  if (!cfg.aiOn) return 'ai_enabled chưa bật (Có) trong TQN_CONFIG';
  var k = PropertiesService.getScriptProperties ? PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') : null;
  if (!k) return 'Chưa có khóa ANTHROPIC_API_KEY trong Script Properties (Apps Script → Cài đặt dự án → Thuộc tính tập lệnh)';
  return '';
}

/** Lấy 1 hồ sơ từ TQN_DATA theo vị trí "Tab!dòng". */
function tqnAiRecord_(at) {
  var sh = tqnSs_().getSheetByName(TQN.SH_DATA);
  if (!sh || sh.getLastRow() < 2) throw new Error('Chưa có TQN_DATA – bấm "Nhập & kiểm tra" trước.');
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TQN_DATA_HEADERS.length).getValues(), C = TQN_COL;
  var key = String(at).split(',')[0].trim();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][C['Nguồn']]) === key) {
    var r = vals[i], o = {};
    TQN_DATA_HEADERS.forEach(function (h, j) { o[h] = r[j] instanceof Date ? Utilities.formatDate(r[j], 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : r[j]; });
    return o;
  }
  throw new Error('Không thấy hồ sơ ở ' + key);
}

/** Câu lệnh gửi AI: chỉ dữ liệu của đúng 1 hồ sơ, không gửi email/SĐT. Dùng chung cho Apps Script và bản xem thử. */
function tqnAiPrompt_(rec) {
  var keep = ['Loại', 'Số hiệu (nguyên văn)', 'Rev', 'Nội dung', 'Tình trạng (sổ chi tiết)', 'Ngày trình', 'TVGS ngày', 'TVGS ý kiến',
    'TVTK ngày', 'TVTK ý kiến', 'CĐT ngày', 'CĐT ý kiến', 'Trình đến', 'Hạn phản hồi', 'Ngày phản hồi', 'Kết quả', 'Tình trạng (Sổ trình duyệt)'];
  var data = keep.map(function (k) { return k + ': ' + String(rec[k] === undefined || rec[k] === '' ? '(trống)' : rec[k]).replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, '[email]').slice(0, 1500); }).join('\n');
  return 'Bạn là thư ký kiểm soát hồ sơ trình duyệt của công trường xây dựng tại Việt Nam (dự án Coastal Quảng Ngãi; các bên: Nhà thầu TBCONS, '
    + 'Tư vấn giám sát TVGS, Tư vấn thiết kế TVTK, Chủ đầu tư CĐT). Đọc dữ liệu 1 hồ sơ dưới đây và trả lời ĐÚNG MỘT đối tượng JSON, không thêm chữ nào khác, theo mẫu:\n'
    + '{"tom_tat":"≤2 câu","muc_do":"CAO|TRUNG_BINH|THAP","ben_xu_ly":"Nhà thầu|TVGS|TVTK|CĐT|Không ai","viec_tiep_theo":"1 câu, động từ đầu câu",'
    + '"can_bo_sung":["mục cần bổ sung/sửa theo ý kiến, nếu có"],"nhap_tra_loi":"nháp văn bản trả lời/đệ trình lại ngắn gọn, lịch sự, tiếng Việt"}\n'
    + 'Quy tắc: chỉ dựa vào dữ liệu đã cho, không bịa số liệu, ngày, tên người. Thiếu thông tin thì ghi rõ "chưa có thông tin". '
    + 'Mức độ CAO nếu quá hạn, bị trả về/không duyệt, hoặc ý kiến yêu cầu sửa quan trọng; THAP nếu đã duyệt xong.\n\n--- DỮ LIỆU HỒ SƠ ---\n' + data;
}

/** Đọc JSON từ câu trả lời AI, kiểm khuôn; lỗi thì ném ra (không ghi bậy). */
function tqnAiParse_(text) {
  var s = String(text || ''), a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a < 0 || b <= a) throw new Error('AI không trả về JSON');
  var o = JSON.parse(s.slice(a, b + 1));
  if (!o.tom_tat || !o.viec_tiep_theo) throw new Error('AI trả về thiếu trường tom_tat / viec_tiep_theo');
  o.muc_do = TQN_AI_LEVELS.indexOf(String(o.muc_do).toUpperCase()) >= 0 ? String(o.muc_do).toUpperCase() : 'TRUNG_BINH';
  o.can_bo_sung = Array.isArray(o.can_bo_sung) ? o.can_bo_sung.map(String).slice(0, 10) : [];
  o.nhap_tra_loi = String(o.nhap_tra_loi || '');
  o.ben_xu_ly = String(o.ben_xu_ly || '');
  return o;
}

function tqnAiSave_(rec, o, model) {
  var sh = tqnOwnSheet_(TQN.PREFIX + 'AI', TQN_AI_HEADERS);
  sh.appendRow([new Date(), rec['Nguồn'], rec['Số hiệu (nguyên văn)'], rec['Loại'], o.muc_do, o.ben_xu_ly, o.tom_tat, o.viec_tiep_theo,
    o.can_bo_sung.join('; '), o.nhap_tra_loi, model, '', 'Bản nháp do AI soạn – cần người kiểm tra trước khi dùng']);
  tqnAudit_('AI_DOC_HO_SO', rec['Nguồn'] + ' ' + rec['Số hiệu (nguyên văn)'], { muc_do: o.muc_do, model: model });
}

/** Gọi Claude qua API Anthropic (thử lại khi 429/5xx). */
function tqnAiCall_(cfg, prompt) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'), last = '';
  for (var i = 0; i < 3; i++) {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: cfg.ai_model, max_tokens: 1200, messages: [{ role: 'user', content: prompt }] })
    });
    var code = res.getResponseCode(), body = res.getContentText();
    if (code === 200) {
      var j = JSON.parse(body);
      return (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    }
    last = 'HTTP ' + code + ' ' + body.slice(0, 200);
    if (code !== 429 && code < 500) break;
    Utilities.sleep(Math.pow(2, i) * 1500);
  }
  throw new Error('Gọi AI lỗi: ' + last);
}

/** Phân tích 1 hồ sơ (gọi từ bảng quy trình). */
function tqnAiAnalyze(at) {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error('Chưa bật AI: ' + why);
  var rec = tqnAiRecord_(at), o = tqnAiParse_(tqnAiCall_(cfg, tqnAiPrompt_(rec)));
  tqnAiSave_(rec, o, cfg.ai_model);
  o.so = rec['Số hiệu (nguyên văn)']; o.at = rec['Nguồn']; o.model = cfg.ai_model;
  return o;
}

/** AI rà các hồ sơ cần xử lý trước (quá hạn / trả về / sắp hạn), tối đa ai_batch_max. */
function tqnAiBatch() {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error('Chưa bật AI: ' + why);
  var sh = tqnSs_().getSheetByName(TQN.PREFIX + 'QUY_TRINH');
  if (!sh || sh.getLastRow() < 2) throw new Error('Chưa có TQN_QUY_TRINH – bấm "Nhập & kiểm tra" trước.');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, TQN_WF_HEADERS.length).getValues()
    .filter(function (r) { return r[11] || /^4/.test(r[0]); }).slice(0, cfg.aiMax);
  var ok = 0, err = [];
  rows.forEach(function (r) { try { tqnAiAnalyze(r[13]); ok++; } catch (e) { err.push(r[3] + ': ' + e.message); } });
  return { ok: ok, err: err, total: rows.length };
}
function tqnMenuAiBatch() {
  tqnEnsureAiConfig_();
  var r = tqnAiBatch();
  SpreadsheetApp.getUi().alert('AI rà hồ sơ', 'Đã đọc ' + r.ok + '/' + r.total + ' hồ sơ → tab TQN_AI (bản nháp, cần người kiểm).' + (r.err.length ? '\nLỗi: ' + r.err.join('\n') : ''), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Theme.gs – GIAO DIỆN DỄ NHÌN cho các sổ của Register (chỉ đổi định dạng: màu nền, màu chữ, đậm; KHÔNG đổi dữ liệu).
 *  - Màu tối / quá gắt → màu nhạt cùng tông; chữ xanh dương gắt → xanh than; dòng tiêu đề → nền xanh nhạt chữ đậm.
 *  - Trước khi đổi, lưu định dạng gốc vào tab TQN_GIAO_DIEN_GOC (ẩn). Nút "Trả lại giao diện gốc" khôi phục y nguyên.
 */
var TQN_THEME_TAB = 'TQN_GIAO_DIEN_GOC', TQN_THEME_ROWS = 600, TQN_THEME_COLS = 40, TQN_THEME_CHUNK = 40;

function tqnHexToHsl_(hex) {
  var m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim()); if (!m) return null;
  var n = parseInt(m[1], 16), r = (n >> 16 & 255) / 255, g = (n >> 8 & 255) / 255, b = (n & 255) / 255;
  var mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, h = 0, s = 0;
  if (mx !== mn) {
    var d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h *= 60;
  }
  return { h: h, s: s, l: l };
}
function tqnHslToHex_(h, s, l) {
  var f = function (n) { var k = (n + h / 30) % 12, a = s * Math.min(l, 1 - l); return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1)))); };
  return '#' + [f(0), f(8), f(4)].map(function (v) { return ('0' + v.toString(16)).slice(-2); }).join('').toUpperCase();
}
/** Bảng màu TQN (cùng tông bảng điều khiển: xanh than #1B3F7A + xanh ngọc #0F6E66). Màu gắt → màu theo VAI TRÒ của ô. */
var TQN_TH = { head: '#E8EEF8', headFc: '#1B3F7A', band: '#DDF0EE', bandFc: '#0F5E58', input: '#FFF8E1', calc: '#F7F8FA', line: '#EEF1F5' };
/** Trả [nền, màu chữ mới hoặc '' (giữ)]. */
function tqnRoleColor_(hex) {
  var c = tqnHexToHsl_(hex); if (!c) return [hex, ''];
  if (c.l >= 0.97) return [hex, ''];                                        // trắng
  if (c.s < 0.12) return [c.l < 0.5 ? TQN_TH.line : TQN_TH.calc, ''];       // xám (ô công thức) / đen
  if (c.l < 0.35) return [TQN_TH.line, ''];                                 // cột ngăn xanh than
  if (c.h >= 170 && c.h <= 215) return [TQN_TH.band, TQN_TH.bandFc];        // dải xanh cyan (tiêu đề khối)
  if (c.h > 50 && c.h <= 70) return [TQN_TH.input, ''];                     // vàng chanh = ô nhập liệu
  if (c.h >= 15 && c.h <= 50) return [TQN_TH.head, TQN_TH.headFc];          // vàng cam / cam = dòng tiêu đề, dòng tổng
  return [tqnHslToHex_(c.h, 0.55, 0.94), ''];                               // màu khác → pastel cùng tông
}
function tqnSoftBg_(hex) { return tqnRoleColor_(hex)[0]; }
function tqnSoftFont_(hex, bg) {
  var c = tqnHexToHsl_(hex), b = tqnHexToHsl_(bg);
  if (!c) return hex;
  if (c.l > 0.85 && b && b.l > 0.8) return '#1B3F7A';           // chữ trắng trên nền đã thành nhạt
  if (c.s > 0.5 && c.h >= 200 && c.h <= 250 && c.l >= 0.25) return '#1B3F7A';   // xanh dương gắt → xanh than TQN
  if (c.s > 0.5 && (c.h <= 10 || c.h >= 350)) return '#B42318';  // đỏ gắt → đỏ trầm
  return hex;
}

function tqnThemeSheets_(cfg) {
  return tqnSs_().getSheets().filter(function (sh) { return sh.getName().indexOf(TQN.PREFIX) !== 0 && !(sh.isSheetHidden && sh.isSheetHidden()); })
    .map(function (sh) { return sh.getName(); });
}

function tqnThemeBackup_(sh, R, C) {
  var bk = tqnOwnSheet_(TQN_THEME_TAB, ['Tab', 'Dòng từ', 'Số dòng', 'Số cột', 'Nền', 'Chữ', 'Đậm', 'Thời điểm']);
  if (bk.getLastRow() >= 2 && bk.getRange(2, 1, bk.getLastRow() - 1, 1).getValues().some(function (r) { return r[0] === sh.getName(); })) return false;  // đã có bản gốc
  for (var r0 = 1; r0 <= R; r0 += TQN_THEME_CHUNK) {
    var n = Math.min(TQN_THEME_CHUNK, R - r0 + 1), rg = sh.getRange(r0, 1, n, C);
    bk.appendRow([sh.getName(), r0, n, C, JSON.stringify(rg.getBackgrounds()), JSON.stringify(rg.getFontColors()), JSON.stringify(rg.getFontWeights()), new Date()]);
  }
  try { if (bk.hideSheet) bk.hideSheet(); } catch (e) { }
  return true;
}

/** Áp giao diện dễ nhìn cho 1 tab. Trả số ô đổi màu. */
function tqnThemeSheet_(sh, cfg) {
  var R = Math.min(Math.max(sh.getLastRow(), 1), TQN_THEME_ROWS), C = Math.min(Math.max(sh.getLastColumn(), 1), TQN_THEME_COLS);
  tqnThemeBackup_(sh, R, C);
  var rg = sh.getRange(1, 1, R, C), bg = rg.getBackgrounds(), fc = rg.getFontColors(), fw = rg.getFontWeights(), vals = rg.getValues(), n = 0;
  // dòng tiêu đề: dòng có "Mã hồ sơ" và các dòng tiêu đề con ngay dưới (khối Nhà thầu/TVGS/CĐT…)
  var hdr = {};
  for (var r = 0; r < Math.min(R, 60); r++) {
    var t = vals[r].map(function (x) { return tqnNoAccent_(tqnClean_(x)); }).join('|');
    if (/ma ho so/.test(t)) { hdr[r] = 1; if (r + 1 < R && /ke hoach|thuc te|ngay trinh|vuong mac/.test(vals[r + 1].map(function (x) { return tqnNoAccent_(tqnClean_(x)); }).join('|'))) hdr[r + 1] = 1; }
  }
  for (var i = 0; i < R; i++) for (var j = 0; j < C; j++) {
    var b0 = bg[i][j], f0 = fc[i][j], b1, f1, w1 = fw[i][j];
    if (hdr[i]) { b1 = TQN_TH.head; f1 = TQN_TH.headFc; w1 = 'bold'; }
    else { var rc = tqnRoleColor_(b0); b1 = rc[0]; f1 = rc[1] || tqnSoftFont_(f0, b1); }
    if (String(b1).toUpperCase() !== String(b0).toUpperCase() || String(f1).toUpperCase() !== String(f0).toUpperCase() || w1 !== fw[i][j]) n++;
    bg[i][j] = b1; fc[i][j] = f1; fw[i][j] = w1;
  }
  rg.setBackgrounds(bg); rg.setFontColors(fc); rg.setFontWeights(fw);
  return n;
}

function tqnApplyTheme() {
  var cfg = tqnConfig_(), out = [], total = 0;
  tqnThemeSheets_(cfg).forEach(function (name) {
    var sh = tqnSs_().getSheetByName(name); if (!sh) return;
    var n = tqnThemeSheet_(sh, cfg); total += n; out.push(name + ': ' + n + ' ô');
  });
  try { tqnApplyStatusColors(); } catch (e) { }
  try { tqnApplyOptionColors(); } catch (e) { }
  tqnAudit_('GIAO_DIEN', 'Áp giao diện dễ nhìn', { o_doi: total, tab: out.length });
  return { total: total, tabs: out };
}

/** Trả lại định dạng gốc từ TQN_GIAO_DIEN_GOC. */
function tqnRestoreTheme() {
  var bk = tqnSs_().getSheetByName(TQN_THEME_TAB);
  if (!bk || bk.getLastRow() < 2) throw new Error('Chưa có bản lưu giao diện gốc.');
  var rows = bk.getRange(2, 1, bk.getLastRow() - 1, 7).getValues(), done = {};
  rows.forEach(function (r) {
    var sh = tqnSs_().getSheetByName(r[0]); if (!sh) return;
    var rg = sh.getRange(+r[1], 1, +r[2], +r[3]);
    rg.setBackgrounds(JSON.parse(r[4])); rg.setFontColors(JSON.parse(r[5])); rg.setFontWeights(JSON.parse(r[6]));
    done[r[0]] = 1;
  });
  tqnAudit_('GIAO_DIEN', 'Trả lại giao diện gốc', { tab: Object.keys(done).length });
  return Object.keys(done).length;
}

function tqnMenuTheme() {
  var r = tqnApplyTheme();
  SpreadsheetApp.getUi().alert('Giao diện dễ nhìn', 'Đã đổi màu ' + r.total + ' ô trên ' + r.tabs.length + ' sổ (chỉ màu, không đổi dữ liệu).\nBản gốc lưu ở tab ẩn TQN_GIAO_DIEN_GOC – bấm "Trả lại giao diện gốc" nếu muốn quay lại.', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRestoreTheme() { SpreadsheetApp.getUi().alert('Đã trả lại giao diện gốc cho ' + tqnRestoreTheme() + ' sổ.'); }

/* ---------------- Ô CHỌN (dropdown) đổi màu theo lựa chọn ---------------- */
var TQN_OPT_COLORS = [   // [từ khóa (không dấu), nền, chữ]
  ['khong duyet', '#FDE9DD', '#A8430A'], ['tra ve', '#FDE9DD', '#A8430A'], ['qua han', '#FCE4E4', '#B42318'], ['duyet co y kien', '#EAF5E3', '#3F6E12'],
  ['da duyet', '#E1F4EF', '#0F6E56'], ['da hoan thanh', '#E1F4EF', '#0F6E56'], ['da phan hoi', '#E5EEFB', '#1B3F7A'],
  ['da trinh', '#E5EEFB', '#1B3F7A'], ['dang', '#FEF3D7', '#8A5A00'], ['chua', '#F1F3F6', '#5F6B7A'], ['huy', '#F1F3F6', '#8A94A3'],
  ['tvgs + cdt', '#EAEFFB', '#3B4FA0'], ['tvgs', '#E5EEFB', '#1B3F7A'], ['cdt', '#EFE9FB', '#5B3FB0'], ['tvtk', '#E3F2F7', '#0B6B8A'], ['tbcons', '#E1F4EF', '#0F6E56']
];
function tqnOptColor_(v) {
  var t = tqnNoAccent_(String(v || '')).trim();
  for (var i = 0; i < TQN_OPT_COLORS.length; i++) if (t.indexOf(TQN_OPT_COLORS[i][0]) === 0 || t === TQN_OPT_COLORS[i][0]) return TQN_OPT_COLORS[i];
  return null;
}
/** Lấy danh sách lựa chọn từ ô chọn sẵn có của Register (Data!B3:B10, danh sách gõ tay…). */
function tqnListValues_(rule) {
  if (!rule) return [];
  try {
    var t = rule.getCriteriaType(), a = rule.getCriteriaValues();
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return (a[0] || []).map(String);
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return a[0].getValues().map(function (r) { return String(r[0]); }).filter(String);
  } catch (e) { }
  return [];
}
/** Thêm màu cho mọi ô chọn (Kết quả, Trình đến, Tình trạng hồ sơ…) của các sổ: chọn giá trị nào, ô đổi màu theo giá trị đó. Chỉ định dạng. */
function tqnApplyOptionColors() {
  var cfg = tqnConfig_(), added = 0, cols = [];
  tqnThemeSheets_(cfg).forEach(function (name) {
    var sh = tqnSs_().getSheetByName(name); if (!sh || !sh.getDataValidations) return;
    var R = Math.min(Math.max(sh.getLastRow(), 1), TQN_THEME_ROWS), C = Math.min(Math.max(sh.getLastColumn(), 1), TQN_THEME_COLS);
    var dv = sh.getRange(1, 1, R, C).getDataValidations(), seen = {}, rules = sh.getConditionalFormatRules(), mine = {}, add = [];
    rules.forEach(function (ru) { var c = ru.getBooleanCondition && ru.getBooleanCondition(); if (c) (c.getCriteriaValues() || []).forEach(function (v) { mine[String(v) + '|' + String(c.getBackground ? c.getBackground() : '').toUpperCase() + '|' + (ru.getRanges ? ru.getRanges()[0].getColumn() : '')] = 1; }); });
    for (var j = 0; j < C; j++) {
      var first = -1, last = -1, opts = null;
      for (var i = 0; i < R; i++) { var v = tqnListValues_(dv[i][j]); if (v.length) { if (first < 0) { first = i; opts = v; } last = i; } }
      if (first < 0) continue;
      var rng = sh.getRange(first + 1, j + 1, Math.max(sh.getMaxRows() - first, 1), 1);
      opts.forEach(function (o) {
        var c = tqnOptColor_(o); if (!c || mine[o + '|' + c[1].toUpperCase() + '|' + (j + 1)]) return;
        add.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(o).setBackground(c[1]).setFontColor(c[2]).setRanges([rng]).build());
      });
      cols.push(name + '!' + tqnColLetter_(j + 1));
    }
    if (add.length) { sh.setConditionalFormatRules(add.concat(rules)); added += add.length; }
  });
  tqnAudit_('MAU_O_CHON', cols.join(', '), { them: added });
  return { added: added, cols: cols };
}

/**
 * Auto.gs – bật/tắt toàn bộ tự động bằng 1 nút. Chạy trên máy chủ Google: máy anh tắt vẫn chạy.
 *  - tqnHourlyJob : mỗi giờ nhập & kiểm lại (chỉ đọc tab gốc, ghi tab TQN_*), không gửi mail.
 *  - tqnDailyJob  : nhắc hạn ~07:00 (Remind.gs).
 *  - tqnWeatherJob: thời tiết ~06:00 (Weather.gs), chỉ khi đủ điều kiện GPS + điều khoản nguồn.
 */
var TQN_AUTO_HANDLERS = ['tqnHourlyJob', 'tqnDailyJob', 'tqnWeatherJob'];

function tqnHourlyJob() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { tqnAudit_('TU_DONG_GIO', 'Bỏ qua: đang có lần chạy khác', {}); return 'BAN'; }
  try {
    tqnEnsureConfig_();
    var cfg = tqnConfig_(), imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), rem = tqnRemindList_(cfg, imp, new Date());
    tqnWriteRemind_(cfg, rem);
    var wf = tqnBuildWorkflow_(cfg, imp, new Date());
    PropertiesService.getDocumentProperties().setProperty('lastHourly', Utilities.formatDate(new Date(), cfg.timezone, 'dd/MM/yyyy HH:mm'));
    tqnAudit_('TU_DONG_GIO', 'Nhập & kiểm tự động', { du_lieu: imp.counts, qc: qc.byRule, nhac: rem.length, quy_trinh: wf.count });
    return 'OK';
  } finally { lock.releaseLock(); }
}

function tqnRemoveHourly_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnHourlyJob') { ScriptApp.deleteTrigger(t); n++; } });
  return n;
}

/** Bật tất cả. Bấm nhiều lần không sinh lịch trùng (xóa lịch cũ cùng tên trước). Trả danh sách dòng báo cáo. */
function tqnSetupAll() {
  tqnEnsureConfig_(); tqnEnsureWeatherConfig_(); tqnEnsureCodeConfig_();
  var out = [];
  tqnRemoveHourly_();
  ScriptApp.newTrigger('tqnHourlyJob').timeBased().everyHours(1).create();
  out.push('✔ Nhập & kiểm tự động: mỗi giờ');
  out.push('✔ Nhắc hạn: hằng ngày khoảng ' + tqnInstallDaily() + ':00' + (tqnConfig_().emailOn ? ' (có gửi mail)' : ' (chưa gửi mail – bật remind_email_enabled = Có để gửi)'));
  var wc = tqnWeatherCfg_(), why = tqnWeatherBlocked_(wc);
  if (why) { tqnRemoveWeather(true); out.push('✖ Thời tiết 06:00 CHƯA bật: ' + why + ' → sửa ở tab TQN_CONFIG rồi bấm lại nút này'); }
  else out.push('✔ Thời tiết: hằng ngày khoảng ' + tqnInstallWeather() + ':00');
  try { var th = tqnApplyTheme(); out.push('✔ Giao diện dễ nhìn: ' + th.tabs.length + ' sổ (bản gốc lưu ở TQN_GIAO_DIEN_GOC)'); } catch (e) { out.push('✖ Giao diện: ' + e.message); }
  try { tqnHourlyJob(); out.push('✔ Đã chạy nhập & kiểm lần đầu'); } catch (e) { out.push('✖ Chạy lần đầu lỗi: ' + e.message); }
  tqnAudit_('BAT_TU_DONG', out.join(' | '), {});
  return out;
}

function tqnRemoveAll() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (TQN_AUTO_HANDLERS.indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); n++; } });
  tqnAudit_('TAT_TU_DONG', 'Đã xóa ' + n + ' lịch', {});
  return n;
}

/** Trạng thái các lịch đang bật + lần chạy gần nhất. */
function tqnAutoStatus() {
  var on = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { on[t.getHandlerFunction()] = (on[t.getHandlerFunction()] || 0) + 1; });
  var p = PropertiesService.getDocumentProperties();
  return [
    (on.tqnHourlyJob ? '✔' : '✖') + ' Nhập & kiểm mỗi giờ' + (p.getProperty('lastHourly') ? ' – lần gần nhất ' + p.getProperty('lastHourly') : ''),
    (on.tqnDailyJob ? '✔' : '✖') + ' Nhắc hạn hằng ngày' + (p.getProperty('lastRemind') ? ' – ' + p.getProperty('lastRemind') : ''),
    (on.tqnWeatherJob ? '✔' : '✖') + ' Thời tiết 06:00' + (p.getProperty('lastWeather') ? ' – ' + p.getProperty('lastWeather') : '')
  ];
}

/* menu */
function tqnMenuSetupAll() {
  SpreadsheetApp.getUi().alert('Bật tự động TQN Register', tqnSetupAll().join('\n') + '\n\nGoogle tự chạy trên máy chủ, máy anh tắt vẫn chạy (có thể lệch ±15 phút).', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRemoveAll() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveAll() + ' lịch tự động.'); }
function tqnMenuAutoStatus() { SpreadsheetApp.getUi().alert('Trạng thái tự động', tqnAutoStatus().join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }

/**
 * Menu.gs – menu "TQN Register" và thanh bên bảng điều khiển.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('TQN Register')
    .addItem('★ Bật tự động tất cả (bấm 1 lần)', 'tqnMenuSetupAll')
    .addItem('Trạng thái tự động', 'tqnMenuAutoStatus')
    .addItem('Tắt tự động tất cả', 'tqnMenuRemoveAll')
    .addSeparator()
    .addItem('Bảng điều khiển', 'tqnShowSidebar')
    .addItem('Quy trình hồ sơ (ai đang giữ việc)', 'tqnMenuWorkflow')
    .addItem('Tạo hồ sơ mới / Trình lại – cấp mã', 'tqnMenuNewDoc')
    .addItem('✦ AI rà hồ sơ cần xử lý (bản nháp)', 'tqnMenuAiBatch')
    .addItem('Giao diện dễ nhìn (đổi màu các sổ)', 'tqnMenuTheme')
    .addItem('Trả lại giao diện gốc', 'tqnMenuRestoreTheme')
    .addItem('Tô màu cột Tình trạng', 'tqnMenuStatusColors')
    .addSeparator()
    .addItem('1. Nhập & kiểm tra ngay', 'tqnMenuRun')
    .addItem('2. Xem nhắc hạn hôm nay', 'tqnMenuRemindPreview')
    .addSeparator()
    .addItem('Bật nhắc hạn hằng ngày', 'tqnMenuInstallDaily')
    .addItem('Tắt nhắc hạn hằng ngày', 'tqnMenuRemoveDaily')
    .addSeparator()
    .addItem('Thời tiết: cập nhật ngay', 'tqnMenuWeatherNow')
    .addItem('Thời tiết: bật 06:00 hằng ngày', 'tqnMenuWeatherOn')
    .addItem('Thời tiết: tắt', 'tqnMenuWeatherOff')
    .addSeparator()
    .addItem('Cài đặt (tạo TQN_CONFIG)', 'tqnMenuSetup')
    .addToUi();
}

/** Chạy nhập + kiểm, có khóa chống 2 người bấm cùng lúc. Trả số liệu cho sidebar. */
function tqnRunAll_() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) throw new Error('Đang có người chạy nhập/kiểm, thử lại sau ít phút.');
  try {
    tqnEnsureConfig_();
    var cfg = tqnConfig_(), imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), rem = tqnRemindList_(cfg, imp, new Date());
    tqnWriteRemind_(cfg, rem);
    var wf = tqnBuildWorkflow_(cfg, imp, new Date());
    tqnAudit_('NHAP_KIEM', 'Chạy tay', { du_lieu: imp.counts, qc: qc.byRule, nhac: rem.length, quy_trinh: wf.count });
    return { cfg: cfg, imp: imp, qc: qc, rem: rem, wf: wf };
  } finally { lock.releaseLock(); }
}

function tqnMenuRun() {
  var r = tqnRunAll_(), c = r.imp.counts;
  var loi = r.qc.rows.filter(function (x) { return x[0] === 'LỖI'; }).length;
  var cb = r.qc.rows.filter(function (x) { return x[0] === 'CẢNH BÁO'; }).length;
  SpreadsheetApp.getUi().alert('TQN Register',
    'Đã nhập (chỉ đọc tab gốc):\n' +
    '• Sổ chi tiết: ' + c.chi_tiet + ' hồ sơ, ghép được với Sổ trình duyệt: ' + c.ghep_duoc + '\n' +
    '• Sổ trình duyệt: ' + c.so_tong + ' dòng (' + c.chi_co_so_tong + ' dòng chỉ có ở sổ tổng)\n' +
    '• TQN_DATA: ' + c.tong_dong + ' dòng\n\n' +
    'Kiểm lỗi: ' + loi + ' lỗi, ' + cb + ' cảnh báo → xem tab TQN_QC\n' +
    'Nhắc hạn: ' + r.rem.length + ' hồ sơ → tab TQN_NHAC_HAN', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRemindPreview() {
  var r = tqnRunAll_();
  tqnSs_().setActiveSheet(tqnSs_().getSheetByName(TQN.SH_REMIND));
  SpreadsheetApp.getUi().alert('Nhắc hạn hôm nay: ' + r.rem.length + ' hồ sơ. Chưa gửi mail (chỉ xem).');
}
function tqnMenuInstallDaily() {
  var h = tqnInstallDaily();
  SpreadsheetApp.getUi().alert('Đã bật nhắc hạn hằng ngày khoảng ' + h + 'h. Google có thể chạy lệch ±15 phút.\nGửi mail chỉ khi TQN_CONFIG: remind_email_enabled = Có và có remind_recipients.');
}
function tqnMenuRemoveDaily() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveDaily(false) + ' lịch nhắc hạn.'); }
function tqnMenuSetup() { tqnEnsureConfig_(); tqnEnsureWeatherConfig_(); tqnEnsureCodeConfig_(); tqnEnsureAiConfig_(); tqnSs_().setActiveSheet(tqnSs_().getSheetByName(TQN.SH_CONFIG)); }

/* ---------------- Thanh bên ---------------- */
function tqnShowSidebar() {
  SpreadsheetApp.getUi().showSidebar(tqnHtml_('Sidebar').setTitle('TQN Register'));
}

/** Dữ liệu cho thanh bên (gọi từ google.script.run). */
function tqnSidebarData() {
  var r = tqnRunAll_(), cfg = r.cfg, tz = cfg.timezone;
  var log = r.imp.src.log, cnt = function (f) { return log.filter(f).length; };
  var st = function (L) { return tqnNoAccent_(L.status); };
  var rules = r.qc.byRule;
  return {
    project: cfg.project_name, version: TQN.VERSION, at: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'),
    kpi: {
      tong: log.length,  // số dòng Sổ trình duyệt
      cho: cnt(function (L) { return st(L) === 'dang cho'; }),
      quaHan: cnt(function (L) { return st(L) === 'qua han'; }),
      duyet: cnt(function (L) { return /da duyet|duyet co y kien/.test(st(L)); }),
      loi: r.qc.rows.filter(function (x) { return x[0] === 'LỖI'; }).length,
      canhBao: r.qc.rows.filter(function (x) { return x[0] === 'CẢNH BÁO'; }).length,
      chiTiet: r.imp.counts.chi_tiet, ghep: r.imp.counts.ghep_duoc
    },
    remindTotal: r.rem.length,
    remind: r.rem.slice(0, 12).map(function (x) { return { g: x[0], han: tqnFmtDate_(x[1], tz), d: x[2], ma: x[4] || x[3], nd: x[6], at: x[9] }; }),
    qc: tqnGroupQC_(r.qc.rows.filter(function (x) { return x[0] !== 'THÔNG TIN'; })).slice(0, 15),
    qcTotal: r.qc.rows.length,
    rules: rules
  };
}

/** Nhảy tới ô "Tab!Dòng" khi bấm trong thanh bên. */
function tqnGoto(at) {
  var m = /^(.+)!([A-Z]*)(\d+)/.exec(String(at).split(',')[0].split(' ↔ ')[0].trim());
  if (!m) return;
  var sh = tqnSs_().getSheetByName(m[1]); if (!sh) return;
  sh.activate(); sh.setActiveRange(sh.getRange(+m[3], 1, 1, Math.min(sh.getLastColumn(), 20)));
}

/** Gom lỗi cùng luật + cùng nội dung cho thanh bên (vd 7 dòng khoảng trắng M0S → 1 mục). */
function tqnGroupQC_(rows) {
  var g = {}, order = [];
  rows.forEach(function (x) {
    var k = x[0] + '|' + x[1] + '|' + (x[1] === 'Q01_TRUNG_MA' || x[1] === 'Q09_CONG_THUC' || x[1] === 'Q08_Y_KIEN_LAP' ? x[2] : x[4]);
    if (!g[k]) { g[k] = { lv: x[0], rule: x[1], at: x[2], msg: x[4], n: 0, pos: [] }; order.push(k); }
    g[k].n++; g[k].pos.push(x[2]);
  });
  return order.map(function (k) { var o = g[k]; o.where = o.pos.join(', '); delete o.pos; return o; });
}

/** Lấy giao diện HTML: ưu tiên file HTML cùng tên trong dự án; nếu không có thì dùng bản gộp sẵn trong code (bản 1 file). */
function tqnHtml_(name) {
  try { return HtmlService.createHtmlOutputFromFile(name); }
  catch (e) {
    if (typeof TQN_HTML_EMBED !== 'undefined' && TQN_HTML_EMBED[name]) return HtmlService.createHtmlOutput(TQN_HTML_EMBED[name]);
    throw e;
  }
}

