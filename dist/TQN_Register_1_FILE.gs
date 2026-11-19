// TQN Register — BẢN 1 FILE: chỉ cần dán toàn bộ file này vào Code.gs (không cần tạo file HTML).

// ===== Config.gs =====
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
  ['han_mac_dinh', '4', 'Số ngày chờ phản hồi để tự tính hạn khi Sổ trình duyệt không có hạn (ngày trình + số ngày)'],
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

// ===== Read.gs =====
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

// ===== Import.gs =====
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
    // Không có hạn từ Sổ trình duyệt (sổ thiếu / chưa ghép) → tự tính: ngày trình + han_mac_dinh ngày (TQN_CONFIG)
    if (!r[21] && tqnIsDate_(d.submitted)) {
      var nd = parseInt(cfg.han_mac_dinh, 10); if (isNaN(nd)) nd = 4;
      r[21] = new Date(d.submitted.getFullYear(), d.submitted.getMonth(), d.submitted.getDate() + nd);
      r[26] = (r[26] ? r[26] + ' · ' : '') + 'Hạn tự tính (+' + nd + ' ngày)';
    }
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

// ===== QC.gs =====
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

// ===== Remind.gs =====
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

// ===== Weather.gs =====
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

// ===== Workflow.gs =====
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
var TQN_STATUS_COLORS = [   // cùng bảng màu ô chọn (TQN_OPT_COLORS)
  ['Quá hạn', '#FEE2E2', '#B91C1C'], ['Không duyệt - trình lại', '#FFEDD5', '#9A3412'], ['Không duyệt', '#FFEDD5', '#9A3412'], ['Đang chờ', '#FEF3C7', '#92400E'],
  ['Chưa trình', '#F1F5F9', '#475569'], ['Đã phản hồi', '#E0E7FF', '#3730A3'], ['Duyệt có ý kiến', '#ECFCCB', '#3F6212'], ['Đã duyệt', '#DCFCE7', '#166534']
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

// ===== Ai.gs =====
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
      payload: JSON.stringify({ model: cfg.ai_model, max_tokens: 1200, messages: [{ role: 'user', content: prompt }] })   // prompt: chữ, hoặc mảng [ảnh/PDF, chữ] khi đọc phiếu
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

// ===== Theme.gs =====
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
var TQN_TH = { head: '#EEF1F5', headFc: '#1E293B', band: '#E2E8F0', bandFc: '#1E293B', input: '#FFF8E1', calc: '#F7F8FA', line: '#EEF1F5' };
/** MÀU THEO BÊN – dùng chung cho tiêu đề khối các sổ, cột "Trình đến", bảng quản lý. Mỗi bên 1 tông, cách xa nhau. */
var TQN_PARTY = [   // [từ khóa đầu nhãn (không dấu), tên, nền tiêu đề, chữ, nền dòng tiêu đề con]
  ['nha thau', 'Nhà thầu', '#C7DCFB', '#1E3A8A', '#E8F0FD'],
  ['tu van giam sat', 'TVGS', '#B7EDE0', '#115E59', '#E3F7F1'],
  ['bql du an', 'TVGS', '#B7EDE0', '#115E59', '#E3F7F1'],
  ['tvqlda', 'TVGS', '#B7EDE0', '#115E59', '#E3F7F1'],
  ['tu van thiet ke', 'TVTK', '#F7D0E2', '#9D174D', '#FCEAF2'],
  ['chu dau tu', 'CĐT', '#FBDDB8', '#8A4B0F', '#FDEEDC'],
  ['cdt phan hoi', 'CĐT', '#FBDDB8', '#8A4B0F', '#FDEEDC']
];
function tqnPartyOf_(label) {
  var t = tqnNoAccent_(tqnClean_(label));
  if (!t) return null;
  for (var i = 0; i < TQN_PARTY.length; i++) if (t.indexOf(TQN_PARTY[i][0]) === 0) return TQN_PARTY[i];
  return null;
}
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
  // khối theo bên trên dòng tiêu đề: từ cột có nhãn bên tới trước nhãn kế tiếp, bỏ cột ngăn trống ở cuối khối
  var party = {};   // 'r,c' → [nền, chữ]
  Object.keys(hdr).forEach(function (k) {
    var r = +k, row = vals[r], sub = hdr[r + 1] ? vals[r + 1] : null;
    for (var c = 0; c < C; c++) {
      var p = tqnPartyOf_(row[c]); if (!p) continue;
      var e = c + 1; while (e < C && !tqnClean_(row[e])) e++;
      var last = e - 1; if (sub) while (last > c && !tqnClean_(sub[last])) last--;
      for (var cc = c; cc <= last; cc++) { party[r + ',' + cc] = [p[2], p[3]]; if (sub) party[(r + 1) + ',' + cc] = [p[4], p[3]]; }
    }
  });
  for (var i = 0; i < R; i++) for (var j = 0; j < C; j++) {
    var b0 = bg[i][j], f0 = fc[i][j], b1, f1, w1 = fw[i][j];
    if (party[i + ',' + j]) { b1 = party[i + ',' + j][0]; f1 = party[i + ',' + j][1]; w1 = 'bold'; }
    else if (hdr[i]) { b1 = TQN_TH.head; f1 = TQN_TH.headFc; w1 = 'bold'; }
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
  var loi = [];
  try { tqnApplyStatusColors(); } catch (e) { loi.push('Tô màu Sổ trình duyệt: ' + e.message); }
  try { tqnApplyOptionColors(); } catch (e) { loi.push('Màu ô chọn: ' + e.message); }
  try { tqnApplyTabColors(); } catch (e) { loi.push('Màu tab: ' + e.message); }
  if (loi.length) tqnAudit_('GIAO_DIEN_LOI', loi.join(' | '), { loi: loi.length });
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
  try { tqnRestoreTabColors(); } catch (e) { }
  tqnAudit_('GIAO_DIEN', 'Trả lại giao diện gốc', { tab: Object.keys(done).length });
  return Object.keys(done).length;
}

function tqnMenuTheme() {
  var r = tqnApplyTheme();
  SpreadsheetApp.getUi().alert('Giao diện dễ nhìn', 'Đã đổi màu ' + r.total + ' ô trên ' + r.tabs.length + ' sổ (chỉ màu, không đổi dữ liệu).\nBản gốc lưu ở tab ẩn TQN_GIAO_DIEN_GOC – bấm "Trả lại giao diện gốc" nếu muốn quay lại.', SpreadsheetApp.getUi().ButtonSet.OK);
}
function tqnMenuRestoreTheme() { SpreadsheetApp.getUi().alert('Đã trả lại giao diện gốc cho ' + tqnRestoreTheme() + ' sổ.'); }

/* ---------------- Ô CHỌN (dropdown) đổi màu theo lựa chọn ---------------- */
var TQN_OPT_COLORS = [   // [từ khóa (không dấu), nền, chữ] – TÌNH TRẠNG theo đèn giao thông, BÊN theo TQN_PARTY; 2 nhóm không trùng tông
  ['qua han', '#FEE2E2', '#B91C1C'],
  ['khong duyet', '#FFEDD5', '#9A3412'], ['tra ve', '#FFEDD5', '#9A3412'], ['trinh lai', '#FFEDD5', '#9A3412'],
  ['duyet co y kien', '#ECFCCB', '#3F6212'],
  ['da duyet', '#DCFCE7', '#166534'], ['da hoan thanh', '#DCFCE7', '#166534'], ['hoan thanh', '#DCFCE7', '#166534'],
  ['da phan hoi', '#E0E7FF', '#3730A3'],
  ['da trinh', '#FEF3C7', '#92400E'], ['dang', '#FEF3C7', '#92400E'], ['gui email', '#FEF3C7', '#92400E'],
  ['chua', '#F1F5F9', '#475569'], ['huy', '#F8FAFC', '#94A3B8'],
  ['tvgs + cdt', '#E2E8F0', '#334155'], ['tvgs', '#B7EDE0', '#115E59'], ['cdt', '#FBDDB8', '#8A4B0F'], ['tvtk', '#F7D0E2', '#9D174D'],
  ['tbcons', '#C7DCFB', '#1E3A8A'], ['nha thau', '#C7DCFB', '#1E3A8A']
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
  var loi = [];
  tqnThemeSheets_(cfg).forEach(function (name) { try {
    var sh = tqnSs_().getSheetByName(name); if (!sh) return;
    var R = Math.min(Math.max(sh.getLastRow(), 1), TQN_THEME_ROWS), C = Math.min(Math.max(sh.getLastColumn(), 1), TQN_THEME_COLS);
    var dv = sh.getRange(1, 1, R, C).getDataValidations(), seen = {}, rules = sh.getConditionalFormatRules(), mine = {}, add = [];
    rules.forEach(function (ru) { var c = ru.getBooleanCondition && ru.getBooleanCondition(); if (c) (c.getCriteriaValues() || []).forEach(function (v) { mine[String(v) + '|' + String(c.getBackground ? c.getBackground() : '').toUpperCase() + '|' + (ru.getRanges && ru.getRanges().length ? ru.getRanges()[0].getColumn() : '')] = 1; }); });
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
  } catch (e) { loi.push(name + ': ' + e.message); } });
  tqnAudit_('MAU_O_CHON', cols.join(', ') + (loi.length ? ' | LỖI: ' + loi.join('; ') : ''), { them: added, loi: loi.length });
  return { added: added, cols: cols };
}


/* ---------------- Màu thanh tab theo nhóm (cùng màu bảng điều khiển) ---------------- */
var TQN_TAB_GROUPS = [
  { ten: 'Sổ tổng', mau: '#1B3F7A', tabs: ['Sổ trình duyệt'] },
  { ten: 'Trình duyệt', mau: '#0F6E66', tabs: ['SD', 'M0S', 'MAT', 'OTS', 'OTS-MEP', 'SCH', 'RFA', 'CRP- Báo cáo', 'Hồ sơ an toàn'] },
  { ten: 'Hỏi · đổi', mau: '#2563EB', tabs: ['RFI', 'RFC'] },
  { ten: 'Giao nhận · văn bản', mau: '#D97706', tabs: ['TRN-GNHS', 'CV-TB đi', 'CV-TB đến', 'Theo dõi phát hành thiết kế'] },
  { ten: 'Hiện trường', mau: '#7C3AED', tabs: ['NCR', 'Chỉ thị - Phát sinh', 'Chỉ thị NT - Phát sinh', 'Hiện trường', 'Hao hụt bê tông', 'Cấp phối bê tông', 'Theo dõi thời tiết', 'Danh mục thẩm tra BPTC'] },
  { ten: 'Kế hoạch · nhân sự', mau: '#94A3B8', tabs: ['SCH-Kế hoạch', 'Lịch họp', 'Cập nhật kế hoạch công việc', 'DS BCH', 'DS Trực CN T4', 'DS Trực CN T7', 'Data', 'Check'] },
  { ten: 'Tool TQN', mau: '#334155', prefix: 'TQN_' }
];
function tqnTabColorOf_(name) {
  var n = String(name).trim();
  for (var i = 0; i < TQN_TAB_GROUPS.length; i++) {
    var g = TQN_TAB_GROUPS[i];
    if (g.prefix && n.indexOf(g.prefix) === 0) return g.mau;
    if (g.tabs && g.tabs.some(function (t) { return t.trim() === n; })) return g.mau;
  }
  return '';
}
/** Chỉ đổi MÀU tab (không đổi tên, thứ tự, dữ liệu). Lưu màu gốc để trả lại. */
function tqnApplyTabColors() {
  var props = PropertiesService.getDocumentProperties(), goc = {};
  try { goc = JSON.parse(props.getProperty('tqn_tab_goc') || '{}'); } catch (e) { goc = {}; }
  var n = 0;
  tqnSs_().getSheets().forEach(function (sh) {
    var c = tqnTabColorOf_(sh.getName()); if (!c) return;
    try { if (!(sh.getName() in goc)) goc[sh.getName()] = sh.getTabColor() || ''; sh.setTabColor(c); n++; } catch (e) { }
  });
  props.setProperty('tqn_tab_goc', JSON.stringify(goc));
  return n;
}
function tqnRestoreTabColors() {
  var props = PropertiesService.getDocumentProperties(), goc = {};
  try { goc = JSON.parse(props.getProperty('tqn_tab_goc') || '{}'); } catch (e) { goc = {}; }
  tqnSs_().getSheets().forEach(function (sh) { if (sh.getName() in goc) sh.setTabColor(goc[sh.getName()] || null); });
  props.setProperty('tqn_tab_goc', '{}');
}

// ===== Docs.gs =====
/**
 * Docs.gs – Đọc sổ VĂN BẢN: CV-TB đi, CV-TB đến, Chỉ thị, NCR. CHỈ ĐỌC, không ghi tab gốc.
 * Mỗi tab dò dòng tiêu đề theo từ khóa, nên tab đổi vị trí cột vẫn đọc được.
 */
var TQN_DOC_TABS = [
  { tab: 'CV-TB đi', loai: 'CV đi', code: ['Mã hồ sơ', 'Số văn bản'], subj: ['Nội dung chi tiết'], date: ['Ngày phát hành', 'Ngày gửi'], reply: ['Ngày phản hồi'], note: ['Nội dung phản hồi'], who: ['Người ký'] },
  { tab: 'CV-TB đến', loai: 'CV đến', code: ['Số văn bản'], subj: ['Nội dung chi tiết'], date: ['Ngày nhận'], reply: ['Ngày phản hồi'], note: ['Nội dung phản hồi'], who: ['Phụ trách', 'Tên người gửi'] },
  { tab: 'Chỉ thị - Phát sinh', loai: 'Chỉ thị', code: ['Số chỉ thị'], subj: ['Danh mục'], date: ['Ngày phát hành'], reply: ['Thực tế'], note: [], who: [] },
  { tab: 'Chỉ thị NT - Phát sinh', loai: 'Chỉ thị NT', code: ['Số chỉ thị'], subj: ['Danh mục'], date: ['Ngày phát hành'], reply: ['Thực tế'], note: [], who: [] },
  { tab: 'NCR', loai: 'NCR', code: ['Mã hồ sơ'], subj: ['Nội dung chi tiết'], date: ['Ngày phát hành'], reply: ['Thời gian kết thúc'], note: ['Vướng mắc/ Trả lời'], who: ['Người phát hành', 'Nhân sự phụ trách'], status: ['Tình trạng'] }
];

/** Ngày từ ô: Date hoặc chuỗi dd/mm/yyyy. */
function tqnAnyDate_(v) {
  if (tqnIsDate_(v)) return v;
  var m = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/.exec(String(v || ''));
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : '';
}

function tqnReadDocTab_(def, today) {
  var sh = tqnSs_().getSheetByName(def.tab);
  if (!sh || sh.getLastRow() < 4) return [];
  var lastR = sh.getLastRow(), lastC = Math.min(sh.getLastColumn(), 30);
  var all = sh.getRange(1, 1, lastR, lastC).getValues();
  var hdrTo = Math.min(6, all.length);
  var find = function (keys) {
    if (!keys || !keys.length) return -1;
    for (var k = 0; k < keys.length; k++) {
      var key = tqnNoAccent_(keys[k]);
      for (var r = 0; r < hdrTo; r++) for (var c = 0; c < lastC; c++)
        if (tqnNoAccent_(tqnClean_(all[r][c])).indexOf(key) === 0) return c;
    }
    return -1;
  };
  var c = { code: find(def.code), subj: find(def.subj), date: find(def.date), reply: find(def.reply), note: find(def.note), who: find(def.who), status: find(def.status) };
  if (c.code < 0) return [];
  var first = 0;
  for (var r = 0; r < hdrTo; r++) if (all[r].some(function (x) { return /^stt$/i.test(String(tqnClean_(x))); })) first = r + 1;
  var t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate()), out = [];
  for (var i = first; i < all.length; i++) {
    var row = all[i], code = String(tqnClean_(tqnCell_(row, c.code)));
    var subj = String(tqnClean_(tqnCell_(row, c.subj)));
    if (!code || !subj) continue;                    // dòng trống / dòng mục "NĂM 2026"
    if (/^(nam \d{4}|[ivx]+\.?)$/i.test(tqnNoAccent_(code))) continue;
    var d = tqnAnyDate_(tqnCell_(row, c.date)), rp = tqnAnyDate_(tqnCell_(row, c.reply));
    var note = String(tqnClean_(tqnCell_(row, c.note))), st = String(tqnClean_(tqnCell_(row, c.status)));
    var done = !!rp || /dong|xong|hoan thanh|da xu ly|da phan hoi/.test(tqnNoAccent_(st));
    out.push({
      loai: def.loai, tab: def.tab, row: i + 1, ma: code, nd: subj.replace(/\s+/g, ' '),
      ngay: d ? Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '',
      phanhoi: rp ? Utilities.formatDate(rp, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '',
      nguoi: String(tqnClean_(tqnCell_(row, c.who))), ykien: note, tt: done ? 'Đã phản hồi' : (st || 'Chờ phản hồi'),
      cho: (!done && d) ? Math.max(0, Math.round((t0 - d) / 86400000)) : ''
    });
  }
  return out;
}

/** Mọi văn bản / NCR / chỉ thị. Gọi từ bảng quản lý. */
function tqnDocsData(today) {
  today = today || new Date();
  var rows = [];
  TQN_DOC_TABS.forEach(function (def) { try { rows = rows.concat(tqnReadDocTab_(def, today)); } catch (e) { } });
  return rows;
}

// ===== Manager.gs =====
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

// ===== KhoiPhuc.gs =====

/* ===== Khôi phục tab "Sổ trình duyệt" từ bản chụp Register 19/11/2026 10:28 (dùng cho bản xem thử). Chạy 1 lần. ===== */
var TQN_SO_TD_CHUP = {"v": [["SỔ THEO DÕI TRÌNH DUYỆT HỒ SƠ", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["Dự án: COASTAL QUẢNG NGÃI", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["Chữ xanh = ô nhập liệu. Cột nền xám = tự tính, không cần nhập. Tình trạng tự cập nhật theo ngày trình, hạn phản hồi và kết quả.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["THỐNG KÊ THEO LOẠI HỒ SƠ", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["STT", "Loại hồ sơ", "Tổng số", "Chưa trình", "Đang chờ", "Quá hạn", "Đã duyệt", "Duyệt có ý kiến", "Không duyệt - trình lại", "Đã phản hồi", "", "", "", "", "", "", ""], [1.0, "RFI", 3, 0, 2, 0, 1, 0, 0, 0, "", "", "", "", "", "", ""], [2.0, "RFA", 2, 0, 1, 0, 0, 0, 1, 0, "", "", "", "", "", "", ""], [3.0, "RFC", 3, 0, 2, 0, 1, 0, 0, 0, "", "", "", "", "", "", ""], [4.0, "SD", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], [5.0, "MAT", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], [6.0, "MS", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], [7.0, "OTS", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], [8.0, "SCH", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], [9.0, "Thanh toán", 6, 0, 3, 1, 2, 0, 0, 0, "", "", "", "", "", "", ""], [10.0, "Mua sắm", 5, 0, 1, 0, 2, 0, 2, 0, "", "", "", "", "", "", ""], [11.0, "Thông báo", 14, 0, 0, 7, 7, 0, 0, 0, "", "", "", "", "", "", ""], [12.0, "Lấy ý kiến", 9, 0, 0, 1, 8, 0, 0, 0, "", "", "", "", "", "", ""], [13.0, "Khác", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", "", "", ""], ["", "TỔNG", 42, 0, 9, 9, 21, 0, 3, 0, "", "", "", "", "", "", ""], ["CHI TIẾT HỒ SƠ TRÌNH DUYỆT", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""], ["STT", "Loại hồ sơ", "Mã hồ sơ", "Rev", "Nội dung chi tiết", "Hạng mục", "Người thực hiện", "Trình đến", "Ngày trình", "Thời hạn phản hồi (ngày)", "Hạn phản hồi", "Ngày phản hồi", "Kết quả", "Số ngày xử lý", "Tình trạng", "Link hồ sơ", "Ghi chú / Vướng mắc"], [1, "RFC", "20261116.COT.TC38.COA6A.001.R00", "00", "Yêu cầu thay đổi - Thay bể tự hoại BTCT đổ tại chỗ bằng bồn tự hoại đúc sẵn", "12 căn mẫu", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, {"__d": "2026-11-16T00:00:00"}, "Đã duyệt", 0, "Đã duyệt", "260611-COA-TBC-RFC-001 Rev01.pdf", "Inhouse (B): Đồng ý đề xuất; đây là biện pháp thi công nên không phát hành lại thiết kế. BQLDA duyệt (A) 16/11."], [2, "RFI", "20261116.COT.TC38.COA2B.V1.001.R00", "00", "Yêu cầu làm rõ thông tin - Về gạch Porcelain và Kết cấu thép", "", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, {"__d": "2026-11-17T00:00:00"}, "Đã duyệt", 1, "Đã duyệt", "261116-COA-TBC-RFI-001 rev01.pdf", "BQLDA (B): Đề nghị phòng QLTK cho ý kiến sớm để TBCONS triển khai. File: 261116-COA-TBC-RFI-001 rev01"], [3, "RFI", "20261116.COT.TC38.COA2B.V1.002.R00", "00", "Yêu cầu làm rõ thông tin - Yêu cầu thông tin gạch Mosaic", "", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, "", "", 3, "Đang chờ", "261116-COA-TBC-RFI-002 rev01.pdf", ""], [4, "RFI", "20261116.COT.TC38.COA2B.V1.003.R00", "00", "Yêu cầu làm rõ thông tin - Cấu tạo nền công viên giai đoạn 2", "Công viên GĐ2", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, "", "", 3, "Đang chờ", "261116-COA-TBC-RFI-003 rev01.pdf", ""], [5, "RFA", "20261116.COT.TC38.COA3.V1.002.R00", "00", "Yêu cầu phê duyệt - Tiến độ 12 căn mẫu Phân khu 1 và Phân khu 3", "12 căn mẫu", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, {"__d": "2026-11-17T00:00:00"}, "Không duyệt - trình lại", 1, "Không duyệt - trình lại", "261116-COA-TBC-RFA-COS-002-rev02.pdf", "BQLDA (D): Thể hiện nhân lực, máy móc, thiết bị cho từng công việc theo khối lượng và định mức. (E): Sửa lại theo nội dung ghi trên bản cứng."], [6, "RFA", "20261116.COT.TC38.COA3.V1.001.R00", "00", "Yêu cầu phê duyệt - Tiến độ Gói thầu Công viên giai đoạn 2", "Công viên GĐ2", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, "", "", 3, "Đang chờ", "261116-COA-TBC-RFA-COS-001-rev02.pdf", ""], [7, "RFC", "20261116.COT.TC38.COA6A.002.R00", "00", "Yêu cầu thay đổi - Thay đổi chủng loại bulong neo", "", "Lý Mai Lương", "CĐT", {"__d": "2026-11-16T00:00:00"}, 4.0, {"__d": "2026-11-20T00:00:00"}, "", "", 3, "Đang chờ", "261116-COA-TBC-RFC-002 rev01.pdf", ""], [8, "RFC", "20261118.COT.TC38.COA6A.001.R00", "00", "Yêu cầu thay đổi - Thay đổi bề dày thép tấm rọ cọc Gói thầu 12 căn mẫu", "12 căn mẫu", "Lý Mai Lương", "CĐT", {"__d": "2026-11-18T00:00:00"}, 4.0, {"__d": "2026-11-22T00:00:00"}, "", "", 1, "Đang chờ", "261118-COA-TBC-RFC-003 rev01-Yêu cầu thay đổi bề dày thép tấm rọ cọc.pdf", ""], [9, "Lấy ý kiến", "20260804.DES.DES01.HAUS04.001.R00", "00", "QLTK gửi Bản rà soát Quy hoạch Dự án Coastal Quảng Ngãi", "", "Trần Thị Dung", "", {"__d": "2026-08-04T00:00:00"}, 2.0, {"__d": "2026-08-06T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [10, "Thông báo", "20260807.DESCOS.DB01.HAUS07.002.R00", "00", "Thông báo phát hành hồ sơ TK BVTC phần Kiến trúc công trình Thương mại TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-07T00:00:00"}, 1.0, {"__d": "2026-08-08T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [11, "Thông báo", "20260807.DESCOS.DB01.HAUS07.001.R00", "00", "Thông báo phát hành hồ sơ vật liệu Kiến trúc công trình Thương mại TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-07T00:00:00"}, 1.0, {"__d": "2026-08-08T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [12, "Thông báo", "20260807.DESCOS.DB01.HAUS07.003.R00", "00", "Thông báo phát hành hồ sơ vật liệu nội thất công trình Thương mại TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-07T00:00:00"}, 1.0, {"__d": "2026-08-08T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [13, "Thông báo", "20260811.DESCOS.DB01.HAUS07.001.R00", "00", "Thông báo phát hành hồ sơ TK BVTC phần Nội thất công trình Thương mại TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-11T00:00:00"}, 1.0, {"__d": "2026-08-12T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [14, "Thông báo", "20260814.DESCOS.DB01.COA07.001.R00", "00", "Cập nhật nội dung chỉnh sửa BVTC nhà TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-14T00:00:00"}, 1.0, {"__d": "2026-08-15T00:00:00"}, "", "", 97, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [15, "Thông báo", "20260819.DESCOS.DB01.HAUS07.001.R00", "00", "Thông báo nhà thầu Catinat trình duyệt vật liệu, thiết bị còn thiếu của công trình nhà TM3", "", "Nguyễn Hữu Công", "", {"__d": "2026-08-19T00:00:00"}, 1.0, {"__d": "2026-08-20T00:00:00"}, "", "", 92, "Quá hạn", "", "Quy trình: HAUS | 07. NOTICE"], [16, "Lấy ý kiến", "20260822.GEN.HAUS04.001.R00", "00", "260811.QNG.COA.NHSL.MOM_Weekly Meeting_REV10_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-08-22T00:00:00"}, 2.0, {"__d": "2026-08-24T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [17, "Lấy ý kiến", "20260822.GEN.HAUS04.002.R00", "00", "260818.QNG.COA.NHSL.MOM_Weekly Meeting_REV11_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-08-22T00:00:00"}, 2.0, {"__d": "2026-08-24T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [18, "Thông báo", "20260827.DES.HAUS07.001.R00", "00", "Hồ sơ cấp điện ngoài ranh - tuyến điện cấp Cocoland", "", "Dương Việt Hùng", "", {"__d": "2026-08-27T00:00:00"}, 1.0, {"__d": "2026-08-28T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [19, "Lấy ý kiến", "20260829.GEN.HAUS04.001.R00", "00", "260827.QNG.COA.NHSL.MOM_Weekly Meeting_REV12_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-08-29T00:00:00"}, 2.0, {"__d": "2026-08-31T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [20, "Thông báo", "20260909.GEN.HAUS07.001.R00", "00", "260908.QNG.COA.NHS_Công văn yêu cầu khắc phục tồn tại, xử lý vướng mắc của các Nhà thầu", "", "Đinh Mạnh Cường", "", {"__d": "2026-09-09T00:00:00"}, 0.0, {"__d": "2026-09-09T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [21, "Lấy ý kiến", "20260913.GEN.HAUS04.001.R00", "00", "260908.QNG.COA.NHSL.MOM_Weekly Meeting_REV13_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-09-13T00:00:00"}, 2.0, {"__d": "2026-09-15T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [22, "Thông báo", "20260922.DESCOS.DB01.DE08.COA07.001.R00", "00", "Công văn Yêu cầu sự có mặt của Giám đốc nhà thầu tại dự án để trực tiếp điều hành thi công", "", "Nguyễn Thị Thu Tuyền", "", {"__d": "2026-09-22T00:00:00"}, 1.0, {"__d": "2026-09-23T00:00:00"}, "", "", 58, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [23, "Thông báo", "20260922.DESCOS.DB01.DE08.COA07.002.R00", "00", "Công văn Cảnh báo chậm trễ đệ trình hồ sơ phê duyệt vật liệu đầu vào nhà Thương mại TM3", "", "Nguyễn Thị Thu Tuyền", "", {"__d": "2026-09-22T00:00:00"}, 1.0, {"__d": "2026-09-23T00:00:00"}, "", "", 58, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [24, "Lấy ý kiến", "20260925.PMCS.PMCS01.COA01.002.R00", "00", "Xin ý kiến phương án xử lý về việc sự cố Cọc BT-542 dừng ép khi chưa đạt Pmin & Lmin", "", "Lê Tấn Thanh Bình", "", {"__d": "2026-09-25T00:00:00"}, 1.0, {"__d": "2026-09-26T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: COA | 01 Lấy ý kiến các phòng ban"], [25, "Lấy ý kiến", "20260928.GEN.HAUS04.001.R00", "00", "260922.QNG.COA.NHSL.MOM_Weekly Meeting_REV14_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-09-28T00:00:00"}, 2.0, {"__d": "2026-09-30T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [26, "Lấy ý kiến", "20260930.PMCS.PMCS01.COA01.001.R00", "00", "Xin ý kiến phương án xử lý về việc sự cố Cọc BT-483 bị vỡ thân cọc khi lực ép đạt Pmax=140T", "", "Lê Tấn Thanh Bình", "", {"__d": "2026-09-30T00:00:00"}, 1.0, {"__d": "2026-10-01T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: COA | 01 Lấy ý kiến các phòng ban"], [27, "Thông báo", "20261003.DES.DES01.HAUS07.001.R00", "00", "Phát hành hồ sơ thiết kế thi công phần cọc toàn bộ cọc của Liền kề PK1 (cập nhật lần 2)", "", "Ngô Anh Tuấn", "", {"__d": "2026-10-03T00:00:00"}, 1.0, {"__d": "2026-10-04T00:00:00"}, "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HAUS | 07. NOTICE"], [28, "Lấy ý kiến", "20261011.GEN.HAUS04.001.R00", "00", "261006.QNG.COA.NHSL.MOM_Weekly Meeting_REV16_Biên bản cuộc họp Nội bộ Tuần", "", "Đinh Mạnh Cường", "", {"__d": "2026-10-11T00:00:00"}, 2.0, {"__d": "2026-10-13T00:00:00"}, "", "", 39, "Quá hạn", "", "Quy trình: HAUS | 04. Lấy ý kiến các phòng ban"], [29, "Thông báo", "20261012.DESCOS.DB01.DE08.COA07.001.R00", "00", "Công văn chấn chỉnh tiến độ thi công, phạt chậm trễ tiến độ theo điều khoản hợp đồng", "", "Nguyễn Thị Thu Tuyền", "", {"__d": "2026-10-12T00:00:00"}, 1.0, {"__d": "2026-10-13T00:00:00"}, "", "", 38, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [30, "Thông báo", "20261015.DESCOS.DB01.DE08.COA07.001.R00", "00", "Công văn chấn chỉnh tiến độ thi công, yêu cầu thay thế nhân sự và bổ sung nhân sự theo hợp đồng", "", "Nguyễn Thị Thu Tuyền", "", {"__d": "2026-10-15T00:00:00"}, 1.0, {"__d": "2026-10-16T00:00:00"}, "", "", 35, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [31, "Thông báo", "20261024.DESCOS.DB01.DE08.COA07.001.R00", "00", "Công văn chậm trễ hoàn thiện và thi công hệ thống MEP phòng kỹ thuật nhà điều hành bán hàng lần 2", "", "Nguyễn Thị Thu Tuyền", "", {"__d": "2026-10-24T00:00:00"}, 1.0, {"__d": "2026-10-25T00:00:00"}, "", "", 26, "Quá hạn", "", "Quy trình: QLDA COA | 07. NOTICE"], [32, "Mua sắm", "20261104.ADTBCO.HA.CH.03.001.R00", "00", "Đề nghị mua bảo hộ lao động", "", "Hồ Thanh Hải", "TBCONS (nội bộ)", {"__d": "2026-11-04T00:00:00"}, "", "", "", "Không duyệt - trình lại", "", "Không duyệt - trình lại", "", "Quy trình: HC 03 | Mua sắm trang thiết bị <= 150tr"], [33, "Mua sắm", "20261106.ADTBCO.HA.CH.03.001.R00", "00", "Đề nghị mua bảo hộ lao động - văn phòng phẩm", "", "Hồ Thanh Hải", "TBCONS (nội bộ)", {"__d": "2026-11-06T00:00:00"}, "", "", "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HC 03 | Mua sắm trang thiết bị <= 150tr"], [34, "Mua sắm", "20261106.ADTBCO.HA.CH.03.003.R00", "00", "Đề nghị mua vật tư văn phòng phẩm", "", "Hồ Thanh Hải", "TBCONS (nội bộ)", {"__d": "2026-11-06T00:00:00"}, "", "", "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: HC 03 | Mua sắm trang thiết bị <= 150tr"], [35, "Thanh toán", "20261109.COT.TC04.C07.TD10.001.R00", "00", "Thanh toán tạm ứng 30% GTHĐ trước VAT - HĐ 0611/2026/HĐTC/QUANGNGAI/TBCONS-VIETPHONG - Thi công Gift Shop Truck: 221.059.686 vnđ", "", "Vũ Huy Cao", "", {"__d": "2026-11-09T00:00:00"}, "", "", "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: TĐ|10. Đề nghị thanh toán (Công ty TBcons)"], [36, "Mua sắm", "20261110.COT.TC38.C07.TD10.001.R00", "00", "COA_TBC_Đệ trình Đề xuất mua sắm vật tư_Thi công điện nước tạm phục vụ thi công", "", "Ngô Văn Điểu", "", {"__d": "2026-11-10T00:00:00"}, "", "", "", "Không duyệt - trình lại", "", "Không duyệt - trình lại", "", "Quy trình: TĐ|10. Đề nghị thanh toán (Công ty TBcons)"], [37, "Thanh toán", "20261112.COT.TC39.NTP01.C07.TBCONS.02C.001.R00", "00", "COA_TBC_Đệ trình Hồ sơ tạm ứng vật tư_Long Thịnh Phát", "", "Ngô Văn Điểu", "TBCONS (nội bộ)", {"__d": "2026-11-12T00:00:00"}, 6.0, {"__d": "2026-11-18T00:00:00"}, "", "", 7, "Quá hạn", "", "Quy trình: TBCONS | 02C. Tạm ứng, thanh quyết toán cho nhà thầu cung cấp"], [38, "Thanh toán", "20261113.COT.TC39.NTP01.C07.TD10.001.R00", "00", "Thanh toán tạm ứng 20% GTĐH số 01 trước VAT - HĐ 1111/HĐMB/QNG/TBC-LTP - Thép xây dựng và vật tư khác: 177.854.487 vnđ", "", "Vũ Huy Cao", "", {"__d": "2026-11-13T00:00:00"}, "", "", "", "Đã duyệt", "", "Đã duyệt", "", "Quy trình: TĐ|10. Đề nghị thanh toán (Công ty TBcons)"], [39, "Thanh toán", "20261116.COT.TC39.NTP04.C07.TD10.001.R00", "00", "Thanh toán tạm ứng 30% GTHĐ sau VAT - HĐ 1211/HĐMB/TBCONS-BLSG - Bulong neo và phụ kiện: 39.415.572 vnđ", "", "Vũ Huy Cao", "", {"__d": "2026-11-16T00:00:00"}, "", "", "", "", 3, "Đang chờ", "", "Quy trình: TĐ|10. Đề nghị thanh toán (Công ty TBcons)"], [40, "Thanh toán", "20261117.COT.TC39.NTP05.C07.TD10.001.R00", "00", "Thanh toán đợt - HĐ 0307/2026/HĐTC/QN/TBC-PA - Thi công văn phòng nhà điều hành BQLDA và nhà kho: 1.729.617.164 vnđ", "", "Vũ Huy Cao", "", {"__d": "2026-11-17T00:00:00"}, "", "", "", "", 2, "Đang chờ", "", "Quy trình: TĐ|10. Đề nghị thanh toán (Công ty TBcons)"], [41, "Thanh toán", "20261117.COT.TC39.NTP01.C07.TBCONS.02C.001.R01", "01", "COA_TBC_Đệ trình Hồ sơ thanh toán vật tư_Long Thịnh Phát", "", "Ngô Văn Điểu", "TBCONS (nội bộ)", {"__d": "2026-11-17T00:00:00"}, 5.0, {"__d": "2026-11-22T00:00:00"}, "", "", 2, "Đang chờ", "", "Quy trình: TBCONS | 02C. Tạm ứng, thanh quyết toán cho nhà thầu cung cấp"], [42, "Mua sắm", "20261119.COT.TC38.C01.TBCONS01.001.R00", "00", "COA_TBC_Đệ trình Đề xuất chủ trương Chỉ định NCC_Bản mã_Công viên GĐ2", "", "Ngô Văn Điểu", "TBCONS (nội bộ)", {"__d": "2026-11-19T00:00:00"}, 2.0, {"__d": "2026-11-21T00:00:00"}, "", "", 0, "Đang chờ", "", "Quy trình: TBCONS | 01. Phê duyệt chủ trương"]], "f": [[7, 3, "=COUNTIFS($B$23:$B$522,B7,$C$23:$C$522,\"<>\")"], [7, 4, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,D$6)"], [7, 5, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,E$6)"], [7, 6, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,F$6)"], [7, 7, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,G$6)"], [7, 8, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,H$6)"], [7, 9, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,I$6)"], [7, 10, "=COUNTIFS($B$23:$B$522,$B7,$O$23:$O$522,J$6)"], [8, 3, "=COUNTIFS($B$23:$B$522,B8,$C$23:$C$522,\"<>\")"], [8, 4, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,D$6)"], [8, 5, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,E$6)"], [8, 6, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,F$6)"], [8, 7, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,G$6)"], [8, 8, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,H$6)"], [8, 9, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,I$6)"], [8, 10, "=COUNTIFS($B$23:$B$522,$B8,$O$23:$O$522,J$6)"], [9, 3, "=COUNTIFS($B$23:$B$522,B9,$C$23:$C$522,\"<>\")"], [9, 4, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,D$6)"], [9, 5, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,E$6)"], [9, 6, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,F$6)"], [9, 7, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,G$6)"], [9, 8, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,H$6)"], [9, 9, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,I$6)"], [9, 10, "=COUNTIFS($B$23:$B$522,$B9,$O$23:$O$522,J$6)"], [10, 3, "=COUNTIFS($B$23:$B$522,B10,$C$23:$C$522,\"<>\")"], [10, 4, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,D$6)"], [10, 5, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,E$6)"], [10, 6, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,F$6)"], [10, 7, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,G$6)"], [10, 8, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,H$6)"], [10, 9, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,I$6)"], [10, 10, "=COUNTIFS($B$23:$B$522,$B10,$O$23:$O$522,J$6)"], [11, 3, "=COUNTIFS($B$23:$B$522,B11,$C$23:$C$522,\"<>\")"], [11, 4, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,D$6)"], [11, 5, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,E$6)"], [11, 6, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,F$6)"], [11, 7, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,G$6)"], [11, 8, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,H$6)"], [11, 9, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,I$6)"], [11, 10, "=COUNTIFS($B$23:$B$522,$B11,$O$23:$O$522,J$6)"], [12, 3, "=COUNTIFS($B$23:$B$522,B12,$C$23:$C$522,\"<>\")"], [12, 4, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,D$6)"], [12, 5, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,E$6)"], [12, 6, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,F$6)"], [12, 7, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,G$6)"], [12, 8, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,H$6)"], [12, 9, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,I$6)"], [12, 10, "=COUNTIFS($B$23:$B$522,$B12,$O$23:$O$522,J$6)"], [13, 3, "=COUNTIFS($B$23:$B$522,B13,$C$23:$C$522,\"<>\")"], [13, 4, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,D$6)"], [13, 5, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,E$6)"], [13, 6, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,F$6)"], [13, 7, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,G$6)"], [13, 8, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,H$6)"], [13, 9, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,I$6)"], [13, 10, "=COUNTIFS($B$23:$B$522,$B13,$O$23:$O$522,J$6)"], [14, 3, "=COUNTIFS($B$23:$B$522,B14,$C$23:$C$522,\"<>\")"], [14, 4, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,D$6)"], [14, 5, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,E$6)"], [14, 6, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,F$6)"], [14, 7, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,G$6)"], [14, 8, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,H$6)"], [14, 9, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,I$6)"], [14, 10, "=COUNTIFS($B$23:$B$522,$B14,$O$23:$O$522,J$6)"], [15, 3, "=COUNTIFS($B$23:$B$522,B15,$C$23:$C$522,\"<>\")"], [15, 4, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,D$6)"], [15, 5, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,E$6)"], [15, 6, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,F$6)"], [15, 7, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,G$6)"], [15, 8, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,H$6)"], [15, 9, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,I$6)"], [15, 10, "=COUNTIFS($B$23:$B$522,$B15,$O$23:$O$522,J$6)"], [16, 3, "=COUNTIFS($B$23:$B$522,B16,$C$23:$C$522,\"<>\")"], [16, 4, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,D$6)"], [16, 5, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,E$6)"], [16, 6, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,F$6)"], [16, 7, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,G$6)"], [16, 8, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,H$6)"], [16, 9, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,I$6)"], [16, 10, "=COUNTIFS($B$23:$B$522,$B16,$O$23:$O$522,J$6)"], [17, 3, "=COUNTIFS($B$23:$B$522,B17,$C$23:$C$522,\"<>\")"], [17, 4, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,D$6)"], [17, 5, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,E$6)"], [17, 6, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,F$6)"], [17, 7, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,G$6)"], [17, 8, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,H$6)"], [17, 9, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,I$6)"], [17, 10, "=COUNTIFS($B$23:$B$522,$B17,$O$23:$O$522,J$6)"], [18, 3, "=COUNTIFS($B$23:$B$522,B18,$C$23:$C$522,\"<>\")"], [18, 4, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,D$6)"], [18, 5, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,E$6)"], [18, 6, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,F$6)"], [18, 7, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,G$6)"], [18, 8, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,H$6)"], [18, 9, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,I$6)"], [18, 10, "=COUNTIFS($B$23:$B$522,$B18,$O$23:$O$522,J$6)"], [19, 3, "=COUNTIFS($B$23:$B$522,B19,$C$23:$C$522,\"<>\")"], [19, 4, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,D$6)"], [19, 5, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,E$6)"], [19, 6, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,F$6)"], [19, 7, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,G$6)"], [19, 8, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,H$6)"], [19, 9, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,I$6)"], [19, 10, "=COUNTIFS($B$23:$B$522,$B19,$O$23:$O$522,J$6)"], [20, 3, "=SUM(C7:C19)"], [20, 4, "=SUM(D7:D19)"], [20, 5, "=SUM(E7:E19)"], [20, 6, "=SUM(F7:F19)"], [20, 7, "=SUM(G7:G19)"], [20, 8, "=SUM(H7:H19)"], [20, 9, "=SUM(I7:I19)"], [20, 10, "=SUM(J7:J19)"], [23, 1, "=IF(C23=\"\",\"\",COUNTA($C$23:C23))"], [23, 11, "=IF(OR(I23=\"\",J23=\"\"),\"\",I23+J23)"], [23, 14, "=IF(I23=\"\",\"\",IF(L23<>\"\",L23-I23,IF(M23<>\"\",\"\",TODAY()-I23)))"], [23, 15, "=IF(C23=\"\",\"\",IF(M23<>\"\",M23,IF(I23=\"\",\"Chưa trình\",IF(L23<>\"\",\"Đã phản hồi\",IF(AND(K23<>\"\",TODAY()>K23),\"Quá hạn\",\"Đang chờ\")))))"], [24, 1, "=IF(C24=\"\",\"\",COUNTA($C$23:C24))"], [24, 11, "=IF(OR(I24=\"\",J24=\"\"),\"\",I24+J24)"], [24, 14, "=IF(I24=\"\",\"\",IF(L24<>\"\",L24-I24,IF(M24<>\"\",\"\",TODAY()-I24)))"], [24, 15, "=IF(C24=\"\",\"\",IF(M24<>\"\",M24,IF(I24=\"\",\"Chưa trình\",IF(L24<>\"\",\"Đã phản hồi\",IF(AND(K24<>\"\",TODAY()>K24),\"Quá hạn\",\"Đang chờ\")))))"], [25, 1, "=IF(C25=\"\",\"\",COUNTA($C$23:C25))"], [25, 11, "=IF(OR(I25=\"\",J25=\"\"),\"\",I25+J25)"], [25, 14, "=IF(I25=\"\",\"\",IF(L25<>\"\",L25-I25,IF(M25<>\"\",\"\",TODAY()-I25)))"], [25, 15, "=IF(C25=\"\",\"\",IF(M25<>\"\",M25,IF(I25=\"\",\"Chưa trình\",IF(L25<>\"\",\"Đã phản hồi\",IF(AND(K25<>\"\",TODAY()>K25),\"Quá hạn\",\"Đang chờ\")))))"], [26, 1, "=IF(C26=\"\",\"\",COUNTA($C$23:C26))"], [26, 11, "=IF(OR(I26=\"\",J26=\"\"),\"\",I26+J26)"], [26, 14, "=IF(I26=\"\",\"\",IF(L26<>\"\",L26-I26,IF(M26<>\"\",\"\",TODAY()-I26)))"], [26, 15, "=IF(C26=\"\",\"\",IF(M26<>\"\",M26,IF(I26=\"\",\"Chưa trình\",IF(L26<>\"\",\"Đã phản hồi\",IF(AND(K26<>\"\",TODAY()>K26),\"Quá hạn\",\"Đang chờ\")))))"], [27, 1, "=IF(C27=\"\",\"\",COUNTA($C$23:C27))"], [27, 11, "=IF(OR(I27=\"\",J27=\"\"),\"\",I27+J27)"], [27, 14, "=IF(I27=\"\",\"\",IF(L27<>\"\",L27-I27,IF(M27<>\"\",\"\",TODAY()-I27)))"], [27, 15, "=IF(C27=\"\",\"\",IF(M27<>\"\",M27,IF(I27=\"\",\"Chưa trình\",IF(L27<>\"\",\"Đã phản hồi\",IF(AND(K27<>\"\",TODAY()>K27),\"Quá hạn\",\"Đang chờ\")))))"], [28, 1, "=IF(C28=\"\",\"\",COUNTA($C$23:C28))"], [28, 11, "=IF(OR(I28=\"\",J28=\"\"),\"\",I28+J28)"], [28, 14, "=IF(I28=\"\",\"\",IF(L28<>\"\",L28-I28,IF(M28<>\"\",\"\",TODAY()-I28)))"], [28, 15, "=IF(C28=\"\",\"\",IF(M28<>\"\",M28,IF(I28=\"\",\"Chưa trình\",IF(L28<>\"\",\"Đã phản hồi\",IF(AND(K28<>\"\",TODAY()>K28),\"Quá hạn\",\"Đang chờ\")))))"], [29, 1, "=IF(C29=\"\",\"\",COUNTA($C$23:C29))"], [29, 11, "=IF(OR(I29=\"\",J29=\"\"),\"\",I29+J29)"], [29, 14, "=IF(I29=\"\",\"\",IF(L29<>\"\",L29-I29,IF(M29<>\"\",\"\",TODAY()-I29)))"], [29, 15, "=IF(C29=\"\",\"\",IF(M29<>\"\",M29,IF(I29=\"\",\"Chưa trình\",IF(L29<>\"\",\"Đã phản hồi\",IF(AND(K29<>\"\",TODAY()>K29),\"Quá hạn\",\"Đang chờ\")))))"], [30, 1, "=IF(C30=\"\",\"\",COUNTA($C$23:C30))"], [30, 11, "=IF(OR(I30=\"\",J30=\"\"),\"\",I30+J30)"], [30, 14, "=IF(I30=\"\",\"\",IF(L30<>\"\",L30-I30,IF(M30<>\"\",\"\",TODAY()-I30)))"], [30, 15, "=IF(C30=\"\",\"\",IF(M30<>\"\",M30,IF(I30=\"\",\"Chưa trình\",IF(L30<>\"\",\"Đã phản hồi\",IF(AND(K30<>\"\",TODAY()>K30),\"Quá hạn\",\"Đang chờ\")))))"], [31, 1, "=IF(C31=\"\",\"\",COUNTA($C$23:C31))"], [31, 11, "=IF(OR(I31=\"\",J31=\"\"),\"\",I31+J31)"], [31, 14, "=IF(I31=\"\",\"\",IF(L31<>\"\",L31-I31,IF(M31<>\"\",\"\",TODAY()-I31)))"], [31, 15, "=IF(C31=\"\",\"\",IF(M31<>\"\",M31,IF(I31=\"\",\"Chưa trình\",IF(L31<>\"\",\"Đã phản hồi\",IF(AND(K31<>\"\",TODAY()>K31),\"Quá hạn\",\"Đang chờ\")))))"], [32, 1, "=IF(C32=\"\",\"\",COUNTA($C$23:C32))"], [32, 11, "=IF(OR(I32=\"\",J32=\"\"),\"\",I32+J32)"], [32, 14, "=IF(I32=\"\",\"\",IF(L32<>\"\",L32-I32,IF(M32<>\"\",\"\",TODAY()-I32)))"], [32, 15, "=IF(C32=\"\",\"\",IF(M32<>\"\",M32,IF(I32=\"\",\"Chưa trình\",IF(L32<>\"\",\"Đã phản hồi\",IF(AND(K32<>\"\",TODAY()>K32),\"Quá hạn\",\"Đang chờ\")))))"], [33, 1, "=IF(C33=\"\",\"\",COUNTA($C$23:C33))"], [33, 11, "=IF(OR(I33=\"\",J33=\"\"),\"\",I33+J33)"], [33, 14, "=IF(I33=\"\",\"\",IF(L33<>\"\",L33-I33,IF(M33<>\"\",\"\",TODAY()-I33)))"], [33, 15, "=IF(C33=\"\",\"\",IF(M33<>\"\",M33,IF(I33=\"\",\"Chưa trình\",IF(L33<>\"\",\"Đã phản hồi\",IF(AND(K33<>\"\",TODAY()>K33),\"Quá hạn\",\"Đang chờ\")))))"], [34, 1, "=IF(C34=\"\",\"\",COUNTA($C$23:C34))"], [34, 11, "=IF(OR(I34=\"\",J34=\"\"),\"\",I34+J34)"], [34, 14, "=IF(I34=\"\",\"\",IF(L34<>\"\",L34-I34,IF(M34<>\"\",\"\",TODAY()-I34)))"], [34, 15, "=IF(C34=\"\",\"\",IF(M34<>\"\",M34,IF(I34=\"\",\"Chưa trình\",IF(L34<>\"\",\"Đã phản hồi\",IF(AND(K34<>\"\",TODAY()>K34),\"Quá hạn\",\"Đang chờ\")))))"], [35, 1, "=IF(C35=\"\",\"\",COUNTA($C$23:C35))"], [35, 11, "=IF(OR(I35=\"\",J35=\"\"),\"\",I35+J35)"], [35, 14, "=IF(I35=\"\",\"\",IF(L35<>\"\",L35-I35,IF(M35<>\"\",\"\",TODAY()-I35)))"], [35, 15, "=IF(C35=\"\",\"\",IF(M35<>\"\",M35,IF(I35=\"\",\"Chưa trình\",IF(L35<>\"\",\"Đã phản hồi\",IF(AND(K35<>\"\",TODAY()>K35),\"Quá hạn\",\"Đang chờ\")))))"], [36, 1, "=IF(C36=\"\",\"\",COUNTA($C$23:C36))"], [36, 11, "=IF(OR(I36=\"\",J36=\"\"),\"\",I36+J36)"], [36, 14, "=IF(I36=\"\",\"\",IF(L36<>\"\",L36-I36,IF(M36<>\"\",\"\",TODAY()-I36)))"], [36, 15, "=IF(C36=\"\",\"\",IF(M36<>\"\",M36,IF(I36=\"\",\"Chưa trình\",IF(L36<>\"\",\"Đã phản hồi\",IF(AND(K36<>\"\",TODAY()>K36),\"Quá hạn\",\"Đang chờ\")))))"], [37, 1, "=IF(C37=\"\",\"\",COUNTA($C$23:C37))"], [37, 11, "=IF(OR(I37=\"\",J37=\"\"),\"\",I37+J37)"], [37, 14, "=IF(I37=\"\",\"\",IF(L37<>\"\",L37-I37,IF(M37<>\"\",\"\",TODAY()-I37)))"], [37, 15, "=IF(C37=\"\",\"\",IF(M37<>\"\",M37,IF(I37=\"\",\"Chưa trình\",IF(L37<>\"\",\"Đã phản hồi\",IF(AND(K37<>\"\",TODAY()>K37),\"Quá hạn\",\"Đang chờ\")))))"], [38, 1, "=IF(C38=\"\",\"\",COUNTA($C$23:C38))"], [38, 11, "=IF(OR(I38=\"\",J38=\"\"),\"\",I38+J38)"], [38, 14, "=IF(I38=\"\",\"\",IF(L38<>\"\",L38-I38,IF(M38<>\"\",\"\",TODAY()-I38)))"], [38, 15, "=IF(C38=\"\",\"\",IF(M38<>\"\",M38,IF(I38=\"\",\"Chưa trình\",IF(L38<>\"\",\"Đã phản hồi\",IF(AND(K38<>\"\",TODAY()>K38),\"Quá hạn\",\"Đang chờ\")))))"], [39, 1, "=IF(C39=\"\",\"\",COUNTA($C$23:C39))"], [39, 11, "=IF(OR(I39=\"\",J39=\"\"),\"\",I39+J39)"], [39, 14, "=IF(I39=\"\",\"\",IF(L39<>\"\",L39-I39,IF(M39<>\"\",\"\",TODAY()-I39)))"], [39, 15, "=IF(C39=\"\",\"\",IF(M39<>\"\",M39,IF(I39=\"\",\"Chưa trình\",IF(L39<>\"\",\"Đã phản hồi\",IF(AND(K39<>\"\",TODAY()>K39),\"Quá hạn\",\"Đang chờ\")))))"], [40, 1, "=IF(C40=\"\",\"\",COUNTA($C$23:C40))"], [40, 11, "=IF(OR(I40=\"\",J40=\"\"),\"\",I40+J40)"], [40, 14, "=IF(I40=\"\",\"\",IF(L40<>\"\",L40-I40,IF(M40<>\"\",\"\",TODAY()-I40)))"], [40, 15, "=IF(C40=\"\",\"\",IF(M40<>\"\",M40,IF(I40=\"\",\"Chưa trình\",IF(L40<>\"\",\"Đã phản hồi\",IF(AND(K40<>\"\",TODAY()>K40),\"Quá hạn\",\"Đang chờ\")))))"], [41, 1, "=IF(C41=\"\",\"\",COUNTA($C$23:C41))"], [41, 11, "=IF(OR(I41=\"\",J41=\"\"),\"\",I41+J41)"], [41, 14, "=IF(I41=\"\",\"\",IF(L41<>\"\",L41-I41,IF(M41<>\"\",\"\",TODAY()-I41)))"], [41, 15, "=IF(C41=\"\",\"\",IF(M41<>\"\",M41,IF(I41=\"\",\"Chưa trình\",IF(L41<>\"\",\"Đã phản hồi\",IF(AND(K41<>\"\",TODAY()>K41),\"Quá hạn\",\"Đang chờ\")))))"], [42, 1, "=IF(C42=\"\",\"\",COUNTA($C$23:C42))"], [42, 11, "=IF(OR(I42=\"\",J42=\"\"),\"\",I42+J42)"], [42, 14, "=IF(I42=\"\",\"\",IF(L42<>\"\",L42-I42,IF(M42<>\"\",\"\",TODAY()-I42)))"], [42, 15, "=IF(C42=\"\",\"\",IF(M42<>\"\",M42,IF(I42=\"\",\"Chưa trình\",IF(L42<>\"\",\"Đã phản hồi\",IF(AND(K42<>\"\",TODAY()>K42),\"Quá hạn\",\"Đang chờ\")))))"], [43, 1, "=IF(C43=\"\",\"\",COUNTA($C$23:C43))"], [43, 11, "=IF(OR(I43=\"\",J43=\"\"),\"\",I43+J43)"], [43, 14, "=IF(I43=\"\",\"\",IF(L43<>\"\",L43-I43,IF(M43<>\"\",\"\",TODAY()-I43)))"], [43, 15, "=IF(C43=\"\",\"\",IF(M43<>\"\",M43,IF(I43=\"\",\"Chưa trình\",IF(L43<>\"\",\"Đã phản hồi\",IF(AND(K43<>\"\",TODAY()>K43),\"Quá hạn\",\"Đang chờ\")))))"], [44, 1, "=IF(C44=\"\",\"\",COUNTA($C$23:C44))"], [44, 11, "=IF(OR(I44=\"\",J44=\"\"),\"\",I44+J44)"], [44, 14, "=IF(I44=\"\",\"\",IF(L44<>\"\",L44-I44,IF(M44<>\"\",\"\",TODAY()-I44)))"], [44, 15, "=IF(C44=\"\",\"\",IF(M44<>\"\",M44,IF(I44=\"\",\"Chưa trình\",IF(L44<>\"\",\"Đã phản hồi\",IF(AND(K44<>\"\",TODAY()>K44),\"Quá hạn\",\"Đang chờ\")))))"], [45, 1, "=IF(C45=\"\",\"\",COUNTA($C$23:C45))"], [45, 11, "=IF(OR(I45=\"\",J45=\"\"),\"\",I45+J45)"], [45, 14, "=IF(I45=\"\",\"\",IF(L45<>\"\",L45-I45,IF(M45<>\"\",\"\",TODAY()-I45)))"], [45, 15, "=IF(C45=\"\",\"\",IF(M45<>\"\",M45,IF(I45=\"\",\"Chưa trình\",IF(L45<>\"\",\"Đã phản hồi\",IF(AND(K45<>\"\",TODAY()>K45),\"Quá hạn\",\"Đang chờ\")))))"], [46, 1, "=IF(C46=\"\",\"\",COUNTA($C$23:C46))"], [46, 11, "=IF(OR(I46=\"\",J46=\"\"),\"\",I46+J46)"], [46, 14, "=IF(I46=\"\",\"\",IF(L46<>\"\",L46-I46,IF(M46<>\"\",\"\",TODAY()-I46)))"], [46, 15, "=IF(C46=\"\",\"\",IF(M46<>\"\",M46,IF(I46=\"\",\"Chưa trình\",IF(L46<>\"\",\"Đã phản hồi\",IF(AND(K46<>\"\",TODAY()>K46),\"Quá hạn\",\"Đang chờ\")))))"], [47, 1, "=IF(C47=\"\",\"\",COUNTA($C$23:C47))"], [47, 11, "=IF(OR(I47=\"\",J47=\"\"),\"\",I47+J47)"], [47, 14, "=IF(I47=\"\",\"\",IF(L47<>\"\",L47-I47,IF(M47<>\"\",\"\",TODAY()-I47)))"], [47, 15, "=IF(C47=\"\",\"\",IF(M47<>\"\",M47,IF(I47=\"\",\"Chưa trình\",IF(L47<>\"\",\"Đã phản hồi\",IF(AND(K47<>\"\",TODAY()>K47),\"Quá hạn\",\"Đang chờ\")))))"], [48, 1, "=IF(C48=\"\",\"\",COUNTA($C$23:C48))"], [48, 11, "=IF(OR(I48=\"\",J48=\"\"),\"\",I48+J48)"], [48, 14, "=IF(I48=\"\",\"\",IF(L48<>\"\",L48-I48,IF(M48<>\"\",\"\",TODAY()-I48)))"], [48, 15, "=IF(C48=\"\",\"\",IF(M48<>\"\",M48,IF(I48=\"\",\"Chưa trình\",IF(L48<>\"\",\"Đã phản hồi\",IF(AND(K48<>\"\",TODAY()>K48),\"Quá hạn\",\"Đang chờ\")))))"], [49, 1, "=IF(C49=\"\",\"\",COUNTA($C$23:C49))"], [49, 11, "=IF(OR(I49=\"\",J49=\"\"),\"\",I49+J49)"], [49, 14, "=IF(I49=\"\",\"\",IF(L49<>\"\",L49-I49,IF(M49<>\"\",\"\",TODAY()-I49)))"], [49, 15, "=IF(C49=\"\",\"\",IF(M49<>\"\",M49,IF(I49=\"\",\"Chưa trình\",IF(L49<>\"\",\"Đã phản hồi\",IF(AND(K49<>\"\",TODAY()>K49),\"Quá hạn\",\"Đang chờ\")))))"], [50, 1, "=IF(C50=\"\",\"\",COUNTA($C$23:C50))"], [50, 11, "=IF(OR(I50=\"\",J50=\"\"),\"\",I50+J50)"], [50, 14, "=IF(I50=\"\",\"\",IF(L50<>\"\",L50-I50,IF(M50<>\"\",\"\",TODAY()-I50)))"], [50, 15, "=IF(C50=\"\",\"\",IF(M50<>\"\",M50,IF(I50=\"\",\"Chưa trình\",IF(L50<>\"\",\"Đã phản hồi\",IF(AND(K50<>\"\",TODAY()>K50),\"Quá hạn\",\"Đang chờ\")))))"], [51, 1, "=IF(C51=\"\",\"\",COUNTA($C$23:C51))"], [51, 11, "=IF(OR(I51=\"\",J51=\"\"),\"\",I51+J51)"], [51, 14, "=IF(I51=\"\",\"\",IF(L51<>\"\",L51-I51,IF(M51<>\"\",\"\",TODAY()-I51)))"], [51, 15, "=IF(C51=\"\",\"\",IF(M51<>\"\",M51,IF(I51=\"\",\"Chưa trình\",IF(L51<>\"\",\"Đã phản hồi\",IF(AND(K51<>\"\",TODAY()>K51),\"Quá hạn\",\"Đang chờ\")))))"], [52, 1, "=IF(C52=\"\",\"\",COUNTA($C$23:C52))"], [52, 11, "=IF(OR(I52=\"\",J52=\"\"),\"\",I52+J52)"], [52, 14, "=IF(I52=\"\",\"\",IF(L52<>\"\",L52-I52,IF(M52<>\"\",\"\",TODAY()-I52)))"], [52, 15, "=IF(C52=\"\",\"\",IF(M52<>\"\",M52,IF(I52=\"\",\"Chưa trình\",IF(L52<>\"\",\"Đã phản hồi\",IF(AND(K52<>\"\",TODAY()>K52),\"Quá hạn\",\"Đang chờ\")))))"], [53, 1, "=IF(C53=\"\",\"\",COUNTA($C$23:C53))"], [53, 11, "=IF(OR(I53=\"\",J53=\"\"),\"\",I53+J53)"], [53, 14, "=IF(I53=\"\",\"\",IF(L53<>\"\",L53-I53,IF(M53<>\"\",\"\",TODAY()-I53)))"], [53, 15, "=IF(C53=\"\",\"\",IF(M53<>\"\",M53,IF(I53=\"\",\"Chưa trình\",IF(L53<>\"\",\"Đã phản hồi\",IF(AND(K53<>\"\",TODAY()>K53),\"Quá hạn\",\"Đang chờ\")))))"], [54, 1, "=IF(C54=\"\",\"\",COUNTA($C$23:C54))"], [54, 11, "=IF(OR(I54=\"\",J54=\"\"),\"\",I54+J54)"], [54, 14, "=IF(I54=\"\",\"\",IF(L54<>\"\",L54-I54,IF(M54<>\"\",\"\",TODAY()-I54)))"], [54, 15, "=IF(C54=\"\",\"\",IF(M54<>\"\",M54,IF(I54=\"\",\"Chưa trình\",IF(L54<>\"\",\"Đã phản hồi\",IF(AND(K54<>\"\",TODAY()>K54),\"Quá hạn\",\"Đang chờ\")))))"], [55, 1, "=IF(C55=\"\",\"\",COUNTA($C$23:C55))"], [55, 11, "=IF(OR(I55=\"\",J55=\"\"),\"\",I55+J55)"], [55, 14, "=IF(I55=\"\",\"\",IF(L55<>\"\",L55-I55,IF(M55<>\"\",\"\",TODAY()-I55)))"], [55, 15, "=IF(C55=\"\",\"\",IF(M55<>\"\",M55,IF(I55=\"\",\"Chưa trình\",IF(L55<>\"\",\"Đã phản hồi\",IF(AND(K55<>\"\",TODAY()>K55),\"Quá hạn\",\"Đang chờ\")))))"], [56, 1, "=IF(C56=\"\",\"\",COUNTA($C$23:C56))"], [56, 11, "=IF(OR(I56=\"\",J56=\"\"),\"\",I56+J56)"], [56, 14, "=IF(I56=\"\",\"\",IF(L56<>\"\",L56-I56,IF(M56<>\"\",\"\",TODAY()-I56)))"], [56, 15, "=IF(C56=\"\",\"\",IF(M56<>\"\",M56,IF(I56=\"\",\"Chưa trình\",IF(L56<>\"\",\"Đã phản hồi\",IF(AND(K56<>\"\",TODAY()>K56),\"Quá hạn\",\"Đang chờ\")))))"], [57, 1, "=IF(C57=\"\",\"\",COUNTA($C$23:C57))"], [57, 11, "=IF(OR(I57=\"\",J57=\"\"),\"\",I57+J57)"], [57, 14, "=IF(I57=\"\",\"\",IF(L57<>\"\",L57-I57,IF(M57<>\"\",\"\",TODAY()-I57)))"], [57, 15, "=IF(C57=\"\",\"\",IF(M57<>\"\",M57,IF(I57=\"\",\"Chưa trình\",IF(L57<>\"\",\"Đã phản hồi\",IF(AND(K57<>\"\",TODAY()>K57),\"Quá hạn\",\"Đang chờ\")))))"], [58, 1, "=IF(C58=\"\",\"\",COUNTA($C$23:C58))"], [58, 11, "=IF(OR(I58=\"\",J58=\"\"),\"\",I58+J58)"], [58, 14, "=IF(I58=\"\",\"\",IF(L58<>\"\",L58-I58,IF(M58<>\"\",\"\",TODAY()-I58)))"], [58, 15, "=IF(C58=\"\",\"\",IF(M58<>\"\",M58,IF(I58=\"\",\"Chưa trình\",IF(L58<>\"\",\"Đã phản hồi\",IF(AND(K58<>\"\",TODAY()>K58),\"Quá hạn\",\"Đang chờ\")))))"], [59, 1, "=IF(C59=\"\",\"\",COUNTA($C$23:C59))"], [59, 11, "=IF(OR(I59=\"\",J59=\"\"),\"\",I59+J59)"], [59, 14, "=IF(I59=\"\",\"\",IF(L59<>\"\",L59-I59,IF(M59<>\"\",\"\",TODAY()-I59)))"], [59, 15, "=IF(C59=\"\",\"\",IF(M59<>\"\",M59,IF(I59=\"\",\"Chưa trình\",IF(L59<>\"\",\"Đã phản hồi\",IF(AND(K59<>\"\",TODAY()>K59),\"Quá hạn\",\"Đang chờ\")))))"], [60, 1, "=IF(C60=\"\",\"\",COUNTA($C$23:C60))"], [60, 11, "=IF(OR(I60=\"\",J60=\"\"),\"\",I60+J60)"], [60, 14, "=IF(I60=\"\",\"\",IF(L60<>\"\",L60-I60,IF(M60<>\"\",\"\",TODAY()-I60)))"], [60, 15, "=IF(C60=\"\",\"\",IF(M60<>\"\",M60,IF(I60=\"\",\"Chưa trình\",IF(L60<>\"\",\"Đã phản hồi\",IF(AND(K60<>\"\",TODAY()>K60),\"Quá hạn\",\"Đang chờ\")))))"], [61, 1, "=IF(C61=\"\",\"\",COUNTA($C$23:C61))"], [61, 11, "=IF(OR(I61=\"\",J61=\"\"),\"\",I61+J61)"], [61, 14, "=IF(I61=\"\",\"\",IF(L61<>\"\",L61-I61,IF(M61<>\"\",\"\",TODAY()-I61)))"], [61, 15, "=IF(C61=\"\",\"\",IF(M61<>\"\",M61,IF(I61=\"\",\"Chưa trình\",IF(L61<>\"\",\"Đã phản hồi\",IF(AND(K61<>\"\",TODAY()>K61),\"Quá hạn\",\"Đang chờ\")))))"], [62, 1, "=IF(C62=\"\",\"\",COUNTA($C$23:C62))"], [62, 11, "=IF(OR(I62=\"\",J62=\"\"),\"\",I62+J62)"], [62, 14, "=IF(I62=\"\",\"\",IF(L62<>\"\",L62-I62,IF(M62<>\"\",\"\",TODAY()-I62)))"], [62, 15, "=IF(C62=\"\",\"\",IF(M62<>\"\",M62,IF(I62=\"\",\"Chưa trình\",IF(L62<>\"\",\"Đã phản hồi\",IF(AND(K62<>\"\",TODAY()>K62),\"Quá hạn\",\"Đang chờ\")))))"], [63, 1, "=IF(C63=\"\",\"\",COUNTA($C$23:C63))"], [63, 11, "=IF(OR(I63=\"\",J63=\"\"),\"\",I63+J63)"], [63, 14, "=IF(I63=\"\",\"\",IF(L63<>\"\",L63-I63,IF(M63<>\"\",\"\",TODAY()-I63)))"], [63, 15, "=IF(C63=\"\",\"\",IF(M63<>\"\",M63,IF(I63=\"\",\"Chưa trình\",IF(L63<>\"\",\"Đã phản hồi\",IF(AND(K63<>\"\",TODAY()>K63),\"Quá hạn\",\"Đang chờ\")))))"], [64, 1, "=IF(C64=\"\",\"\",COUNTA($C$23:C64))"], [64, 11, "=IF(OR(I64=\"\",J64=\"\"),\"\",I64+J64)"], [64, 14, "=IF(I64=\"\",\"\",IF(L64<>\"\",L64-I64,IF(M64<>\"\",\"\",TODAY()-I64)))"], [64, 15, "=IF(C64=\"\",\"\",IF(M64<>\"\",M64,IF(I64=\"\",\"Chưa trình\",IF(L64<>\"\",\"Đã phản hồi\",IF(AND(K64<>\"\",TODAY()>K64),\"Quá hạn\",\"Đang chờ\")))))"]]};
function tqnKhoiPhucSoTrinhDuyet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), name = 'Sổ trình duyệt';
  if (ss.getSheetByName(name)) throw new Error('Đã có tab ' + name + ' – không ghi đè.');
  var at = ss.getSheetByName('RFI') ? ss.getSheetByName('RFI').getIndex() : ss.getNumSheets();
  var sh = ss.insertSheet(name, at);
  var v = TQN_SO_TD_CHUP.v.map(function (r) { return r.map(function (x) { return (x && typeof x === 'object' && x.__d) ? new Date(x.__d) : x; }); });
  sh.getRange(1, 1, v.length, v[0].length).setValues(v);
  tqnGhiCongThucSoTD_(sh);
  tqnDinhDangNgaySoTD();
  tqnAudit_('KHOI_PHUC_SO_TD', 'Tạo tab Sổ trình duyệt từ bản chụp 19/11 10:28 (' + v.length + ' dòng, ' + TQN_SO_TD_CHUP.f.length + ' công thức)', {});
  return v.length;
}

/** Định dạng ngày cho các cột Ngày trình / Hạn phản hồi / Ngày phản hồi của tab Sổ trình duyệt (công thức I+J ra số nếu cột chưa định dạng ngày). */
function tqnDinhDangNgaySoTD() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sổ trình duyệt'); if (!sh) return 0;
  var n = Math.max(sh.getMaxRows() - 22, 1);
  ['I', 'K', 'L'].forEach(function (c) { sh.getRange(c + '23:' + c + (22 + n)).setNumberFormat('dd/MM/yyyy'); });
  return n;
}

/** Đổi dấu ngăn tham số "," → ";" (ngoài chuỗi) khi bảng tính dùng kiểu số Việt Nam / châu Âu. */
function tqnCongThucTheoLocale_(f, locale) {
  if (!/^(vi|de|fr|es|it|pt|ru|nl|pl|tr|id)/i.test(String(locale || ''))) return f;
  var out = '', q = false;
  for (var i = 0; i < f.length; i++) { var ch = f[i]; if (ch === '"') q = !q; out += (!q && ch === ',') ? ';' : ch; }
  return out;
}
function tqnGhiCongThucSoTD_(sh) {
  var loc = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  TQN_SO_TD_CHUP.f.forEach(function (x) { sh.getRange(x[0], x[1]).setFormula(tqnCongThucTheoLocale_(x[2], loc)); });
}
/** Sửa lại công thức tab Sổ trình duyệt đã khôi phục (khi trước đó bị #ERROR! vì dấu ngăn). */
function tqnSuaCongThucSoTD() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sổ trình duyệt'); if (!sh) throw new Error('Chưa có tab Sổ trình duyệt');
  tqnGhiCongThucSoTD_(sh); tqnDinhDangNgaySoTD(); SpreadsheetApp.flush();
  var k = sh.getRange('K36').getValue(), o = sh.getRange('O36').getValue();
  console.log('K36=' + k + ' | O36=' + o);
  return [String(k), String(o)];
}

// ===== Auto.gs =====
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

// ===== Menu.gs =====
/**
 * Menu.gs – menu "TQN Register" và thanh bên bảng điều khiển.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('TQN Register')
    .addItem('▣ Quản lý hồ sơ (bảng chính)', 'tqnMenuManager')
    .addSeparator()
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

// ===== Giao diện HTML gộp sẵn =====
var TQN_HTML_EMBED = {"Sidebar": "<!DOCTYPE html>\n<html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\n:root{--bg:#f4f6f9;--card:#fff;--bd:#e3e8ef;--tx:#18212d;--tx2:#4d5a6b;--tx3:#8792a2;--navy:#1B3F7A;\n--red:#c5221f;--redbg:#fce8e6;--amb:#8a6100;--ambbg:#fef7e0;--grn:#1e7b3a;--grnbg:#e6f4ea;--blu:#1a56b8;--blubg:#e8f0fe}\n*{box-sizing:border-box}body{margin:0;font:12.5px/1.45 Inter,\"Segoe UI\",Roboto,Arial,sans-serif;color:var(--tx);background:var(--bg)}\n.hd{display:flex;align-items:center;gap:9px;padding:12px;background:var(--card);border-bottom:1px solid var(--bd)}\n.lg{width:34px;height:26px;border-radius:7px;background:var(--navy);display:flex;align-items:center;justify-content:center;color:#fff;font:800 10.5px Inter;flex:none}.lg b{color:#2DD4BF}\n.hd .t b{display:block;font-size:13.5px}.hd .t span{font-size:11px;color:var(--tx3)}\n.tabs{display:flex;background:var(--card);border-bottom:1px solid var(--bd);padding:0 6px}\n.tabs button{flex:1;border:0;background:none;padding:9px 4px;font:600 12px inherit;color:var(--tx2);cursor:pointer;border-bottom:2px solid transparent}\n.tabs button.on{color:var(--navy);border-bottom-color:var(--navy)}\n.tabs i{font-style:normal;font-size:10.5px;background:var(--bg);border-radius:999px;padding:0 6px;margin-left:3px;color:var(--tx2)}\n.tabs button.on i{background:var(--blubg);color:var(--navy)}\n.wrap{padding:12px}\n.kp{display:grid;grid-template-columns:1fr 1fr;gap:8px}\n.k{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:9px 10px;border-left:4px solid var(--c,#cbd5e1)}\n.k .l{color:var(--tx2);font-size:11px;font-weight:600}.k .v{font-size:22px;font-weight:700;line-height:1.2;color:var(--vc,var(--tx))}\n.note{color:var(--tx3);font-size:11px;margin-top:8px}\n.grp{margin-top:14px}\n.grp h5{margin:0 0 6px;font-size:10.5px;letter-spacing:.6px;text-transform:uppercase;color:var(--tx3);font-weight:700}\n.act{display:flex;align-items:center;gap:10px;width:100%;text-align:left;border:1px solid var(--bd);background:var(--card);border-radius:10px;padding:8px 10px;margin-bottom:6px;cursor:pointer;font:inherit;color:var(--tx);transition:border-color .15s,box-shadow .15s}\n.act:hover{border-color:var(--ac);box-shadow:0 2px 8px rgba(16,24,40,.06)}\n.act .ic{width:30px;height:30px;border-radius:8px;display:grid;place-items:center;background:var(--ab);color:var(--ac);flex:none}\n.act b{display:block;font-size:12.5px}.act span{display:block;font-size:11px;color:var(--tx3)}\n.act.pri{background:var(--navy);border-color:var(--navy);color:#fff}.act.pri span{color:#c9d6ee}.act.pri .ic{background:rgba(255,255,255,.15);color:#fff}\n.act:disabled{opacity:.6;cursor:wait}\n.it{background:var(--card);border:1px solid var(--bd);border-left:4px solid var(--c);border-radius:9px;padding:8px 10px;margin-bottom:6px;cursor:pointer}\n.it:hover{border-color:var(--c)}\n.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:10.5px;font-weight:700;background:var(--b);color:var(--c)}\n.it .m{font-size:11px;color:var(--tx3)}.it .c2{font:11px ui-monospace,Consolas,monospace;color:var(--tx3);word-break:break-all;margin-top:2px}\n.it .d{margin-top:2px}\n.load{padding:28px;text-align:center;color:var(--tx3)}.err{background:var(--redbg);color:var(--red);padding:9px;border-radius:9px}\n.toast{position:fixed;left:10px;right:10px;bottom:10px;background:#18212d;color:#fff;border-radius:9px;padding:9px 11px;font-size:12px;display:none}\n</style></head>\n<body>\n<div class=\"hd\"><div class=\"lg\">T<b>Q</b>N</div><div class=\"t\"><b>TQN Register</b><span id=\"pj\">Đang tải…</span></div></div>\n<div class=\"tabs\" id=\"tabs\"></div>\n<div class=\"wrap\" id=\"app\"><div class=\"load\">Đang đọc Register và kiểm tra…</div></div>\n<div class=\"toast\" id=\"toast\"></div>\n<script>\nvar ICO={\"refresh\": \"<path d=\\\"M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7\\\"/>\", \"board\": \"<rect x=\\\"3\\\" y=\\\"4\\\" width=\\\"18\\\" height=\\\"16\\\" rx=\\\"2\\\"/><path d=\\\"M9 4v16M15 4v16\\\"/>\", \"plus\": \"<path d=\\\"M12 5v14M5 12h14\\\"/>\", \"panel\": \"<rect x=\\\"3\\\" y=\\\"4\\\" width=\\\"18\\\" height=\\\"16\\\" rx=\\\"2\\\"/><path d=\\\"M15 4v16\\\"/>\", \"star\": \"<path d=\\\"M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z\\\"/>\", \"clock\": \"<circle cx=\\\"12\\\" cy=\\\"12\\\" r=\\\"9\\\"/><path d=\\\"M12 7v5l3 2\\\"/>\", \"cloud\": \"<path d=\\\"M7 18h10a4 4 0 0 0 .5-8 6 6 0 0 0-11.5 1.5A3.5 3.5 0 0 0 7 18z\\\"/>\", \"palette\": \"<path d=\\\"M12 3a9 9 0 1 0 0 18c1.1 0 1.5-.8 1.5-1.5 0-.9-.7-1.2-.7-2s.7-1.5 1.7-1.5H17a4 4 0 0 0 4-4c0-4.4-4-9-9-9z\\\"/><circle cx=\\\"7.5\\\" cy=\\\"11\\\" r=\\\"1\\\"/><circle cx=\\\"10\\\" cy=\\\"7\\\" r=\\\"1\\\"/><circle cx=\\\"15\\\" cy=\\\"7.5\\\" r=\\\"1\\\"/>\", \"undo\": \"<path d=\\\"M9 14L4 9l5-5\\\"/><path d=\\\"M4 9h10a6 6 0 0 1 0 12h-3\\\"/>\", \"spark\": \"<path d=\\\"M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z\\\"/><path d=\\\"M19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z\\\"/>\"};function svg(n){return '<svg width=\"15\" height=\"15\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\">'+(ICO[n]||\"\")+\"</svg>\"}\nvar D=null,TAB='home';\nvar LV={'QUÁ HẠN':['#c5221f','#fce8e6'],'HẠN HÔM NAY':['#8a6100','#fef7e0'],'SẮP ĐẾN HẠN':['#1a56b8','#e8f0fe'],'LỖI':['#c5221f','#fce8e6'],'CẢNH BÁO':['#8a6100','#fef7e0']};\nvar GROUPS=[\n ['Việc hằng ngày',[['tqnMenuManager','board','Quản lý hồ sơ','Hôm nay cần làm gì · phiếu · giao nhận',1],['tqnSidebarRefresh','refresh','Nhập & kiểm tra','Đọc lại các sổ, báo lỗi, nhắc hạn',1],['tqnMenuWorkflow','board','Kiểm soát quy trình','Ai đang giữ việc, hồ sơ quá hạn',1],['tqnMenuNewDoc','plus','Tạo hồ sơ / cấp mã','Cấp số hiệu, trình lại Rev+1',1]],'#1B3F7A','#e8f0fe'],\n ['Trợ lý AI',[['tqnMenuAiBatch','spark','AI rà hồ sơ cần xử lý','Tóm tắt, việc tiếp theo, nháp trả lời']],'#5b3fd1','#f1edfc'],\n ['Tự động',[['tqnMenuSetupAll','star','Bật tự động tất cả','Mỗi giờ · nhắc 07:00 · thời tiết 06:00'],['tqnMenuAutoStatus','clock','Trạng thái tự động','Xem lịch đang chạy']],'#0f7a6d','#e6f6f4'],\n ['Thời tiết',[['tqnMenuWeatherNow','cloud','Cập nhật thời tiết ngay','Ghi vào tab Theo dõi thời tiết']],'#8a6100','#fef7e0'],\n ['Giao diện',[['tqnMenuTheme','palette','Giao diện dễ nhìn','Đổi màu các sổ (chỉ màu)'],['tqnMenuRestoreTheme','undo','Trả lại giao diện gốc','Khôi phục màu ban đầu']],'#4d5a6b','#f1f3f6']];\nfunction esc(s){return String(s==null?'':s).replace(/[&<>\"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]})}\nfunction toast(t){var e=document.getElementById('toast');e.textContent=t;e.style.display='block';clearTimeout(e._t);e._t=setTimeout(function(){e.style.display='none'},2800)}\nfunction load(){document.getElementById('app').innerHTML='<div class=\"load\">Đang đọc Register và kiểm tra…</div>';\n google.script.run.withSuccessHandler(function(d){D=d;draw()}).withFailureHandler(function(e){document.getElementById('app').innerHTML='<div class=\"err\">'+esc(e.message||e)+'</div>'}).tqnSidebarData();}\nfunction go(at){google.script.run.tqnGoto(at)}\nfunction run(fn,btn){if(fn==='tqnSidebarRefresh'){load();return}\n btn.disabled=true;google.script.run.withSuccessHandler(function(){btn.disabled=false;if(/Theme|SetupAll|WeatherNow|AiBatch/.test(fn))load()}).withFailureHandler(function(e){btn.disabled=false;toast(e.message||String(e))})[fn]()}\nfunction draw(){\n var d=D,k=d.kpi;document.getElementById('pj').textContent=d.project+' · '+d.at;\n document.getElementById('tabs').innerHTML=[['home','Tổng quan',''],['rem','Nhắc hạn',d.remindTotal||d.remind.length],['qc','Lỗi dữ liệu',k.loi+k.canhBao]].map(function(t){return '<button data-t=\"'+t[0]+'\" class=\"'+(TAB===t[0]?'on':'')+'\">'+t[1]+(t[2]!==''?'<i>'+t[2]+'</i>':'')+'</button>'}).join('');\n Array.prototype.forEach.call(document.querySelectorAll('#tabs button'),function(b){b.onclick=function(){TAB=b.getAttribute('data-t');draw()}});\n var h='';\n if(TAB==='home'){\n  h+='<div class=\"kp\">'+kp('Sổ trình duyệt',k.tong,'#94a3b8')+kp('Đang chờ',k.cho,'#e0a100','#8a6100')+kp('Quá hạn',k.quaHan,'#c5221f','#c5221f')+kp('Đã duyệt',k.duyet,'#1e7b3a','#1e7b3a')+'</div>';\n  h+='<div class=\"note\">Sổ chi tiết '+k.chiTiet+' hồ sơ · ghép được với sổ tổng '+k.ghep+' · '+(k.loi+k.canhBao)+' lỗi/cảnh báo</div>';\n  GROUPS.forEach(function(g){h+='<div class=\"grp\"><h5>'+g[0]+'</h5>'+g[1].map(function(a){return '<button class=\"act'+(a[4]&&a[0]==='tqnSidebarRefresh'?' pri':'')+'\" data-fn=\"'+a[0]+'\" style=\"--ac:'+g[2]+';--ab:'+g[3]+'\"><span class=\"ic\">'+svg(a[1])+'</span><span><b>'+esc(a[2])+'</b><span>'+esc(a[3])+'</span></span></button>'}).join('')+'</div>'});\n }else if(TAB==='rem'){\n  h+=d.remind.length?d.remind.map(function(x){var c=LV[x.g]||['#4d5a6b','#f1f3f6'];return '<div class=\"it\" style=\"--c:'+c[0]+';--b:'+c[1]+'\" data-at=\"'+esc(x.at)+'\"><span class=\"pill\">'+esc(x.g)+(x.d!==''?' · '+esc(Math.abs(x.d))+' ngày':'')+'</span> <span class=\"m\">hạn '+esc(x.han)+'</span><div class=\"d\">'+esc(String(x.nd).slice(0,110))+'</div><div class=\"c2\">'+esc(x.ma)+'</div></div>'}).join(''):'<div class=\"load\">✓ Không có hồ sơ quá hạn hoặc sắp đến hạn</div>';\n  if(d.remindTotal>d.remind.length)h+='<div class=\"note\">Hiện '+d.remind.length+'/'+d.remindTotal+' việc gần nhất · đủ ở tab TQN_NHAC_HAN</div>';\n }else{\n  h+=d.qc.map(function(x){var c=LV[x.lv]||['#4d5a6b','#f1f3f6'];return '<div class=\"it\" style=\"--c:'+c[0]+';--b:'+c[1]+'\" data-at=\"'+esc(x.at)+'\"><span class=\"pill\">'+esc(x.lv)+(x.n>1?' · '+x.n+' mục':'')+'</span> <span class=\"m\">'+esc(x.rule)+'</span><div class=\"d\">'+esc(String(x.msg).slice(0,160))+'</div><div class=\"m\">'+esc(String(x.where).slice(0,80))+'</div></div>'}).join('')+'<div class=\"note\">Đủ ở tab TQN_QC</div>';\n }\n h+='<div class=\"note\" style=\"margin-top:14px\">Bấm một mục để nhảy tới dòng trong sổ · bản '+esc(d.version)+'</div>';\n document.getElementById('app').innerHTML=h;\n Array.prototype.forEach.call(document.querySelectorAll('.act'),function(b){b.onclick=function(){run(b.getAttribute('data-fn'),b)}});\n Array.prototype.forEach.call(document.querySelectorAll('.it[data-at]'),function(b){b.onclick=function(){go(b.getAttribute('data-at'))}});\n}\nfunction kp(l,v,c,vc){return '<div class=\"k\" style=\"--c:'+c+(vc?';--vc:'+vc:'')+'\"><div class=\"l\">'+l+'</div><div class=\"v\">'+v+'</div></div>'}\nload();\n</script>\n</body></html>\n", "Workflow": "<!DOCTYPE html>\n<html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\n:root{--bg:#f4f6f9;--card:#fff;--bd:#e3e8ef;--bd2:#eef1f5;--tx:#18212d;--tx2:#4d5a6b;--tx3:#8792a2;--navy:#1B3F7A;--teal:#0f9d8f;\n--red:#d93a3a;--redbg:#fdeeee;--amber:#d98a00;--amberbg:#fff6e5;--ok:#0f9d8f;--okbg:#e6f6f4;--gray:#9aa5b4;--graybg:#f1f3f6;--green:#1f9d55;--greenbg:#e8f6ee}\n*{box-sizing:border-box}html,body{margin:0;height:100%}\nbody{font:13.5px/1.45 Inter,\"Segoe UI\",Roboto,Arial,sans-serif;color:var(--tx);background:var(--bg);display:flex;flex-direction:column;overflow:hidden}\n/* thanh trên */\n.hd{display:flex;align-items:center;gap:14px;padding:14px 22px;background:var(--card);border-bottom:1px solid var(--bd)}\n.lg{width:40px;height:30px;border-radius:8px;background:var(--navy);display:flex;align-items:center;justify-content:center;color:#fff;font:800 12px Inter;flex:none}\n.lg b{color:#2DD4BF}\n.tt{min-width:0}.tt b{display:block;font-size:16px;font-weight:700;white-space:nowrap}.tt span{font-size:12px;color:var(--tx3);white-space:nowrap}\n.sp{flex:1}\n.srch{display:flex;align-items:center;gap:7px;border:1px solid var(--bd);border-radius:9px;padding:7px 11px;background:var(--bg);width:250px}\n.srch input{border:0;outline:0;background:none;font:13px inherit;width:100%;color:var(--tx)}\n.btn-ai{display:flex;align-items:center;gap:6px;border:1px solid #d8cff8;background:#f5f2ff;color:#5b3fd1;border-radius:9px;padding:7px 12px;font:600 12.5px Inter;cursor:pointer;white-space:nowrap}\n.btn-ai:hover{background:#ebe5ff}\n.seg{display:flex;background:var(--bg);border:1px solid var(--bd);border-radius:9px;padding:3px}\n.seg button{border:0;background:none;padding:6px 14px;font:600 12.5px inherit;color:var(--tx2);cursor:pointer;border-radius:7px;white-space:nowrap}\n.seg button.on{background:var(--card);color:var(--navy);box-shadow:0 1px 3px rgba(16,24,40,.12)}\n.main{flex:1;min-height:0;overflow:auto;padding:18px 22px 24px}\n/* thẻ số */\n.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px}\n.kp{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px 16px;cursor:pointer;border-top:4px solid var(--c);transition:box-shadow .15s}\n.kp:hover{box-shadow:0 4px 14px rgba(16,24,40,.08)}\n.kp.on{box-shadow:0 0 0 2px var(--c)}\n.kp .l{font-size:12px;color:var(--tx2);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.kp .n{font-size:30px;font-weight:700;line-height:1.15;margin-top:6px;font-variant-numeric:tabular-nums;color:var(--tx)}\n.kp.alert{background:var(--redbg);border-color:#f5c7c7}.kp.alert .n,.kp.alert .l{color:var(--red)}\n.hint{font-size:12px;color:var(--tx3);margin:12px 2px 0}\n.flt{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:12px 0 0;font-size:12.5px;color:var(--tx2)}\n.tag{display:inline-flex;gap:6px;align-items:center;background:#eaf0fb;color:var(--navy);border-radius:999px;padding:3px 10px;font-weight:600}\n.tag button{border:0;background:none;color:var(--navy);cursor:pointer;font-size:15px;line-height:1;padding:0}\n/* khối */\n.grid2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.25fr);gap:16px;margin-top:16px}\n.box{background:var(--card);border:1px solid var(--bd);border-radius:12px}\n.bh{display:flex;align-items:baseline;gap:8px;padding:14px 16px 10px;border-bottom:1px solid var(--bd2)}\n.bh b{font-size:14px}.bh span{font-size:12px;color:var(--tx3)}\n.bb{padding:8px 16px 14px}\n/* theo bên */\n.prow{display:grid;grid-template-columns:150px 1fr 44px;gap:12px;align-items:center;padding:11px 0;border-bottom:1px solid var(--bd2)}\n.prow:last-child{border-bottom:0}\n.pn{display:flex;align-items:center;gap:9px;font-weight:600;white-space:nowrap;overflow:hidden}\n.av{width:28px;height:28px;border-radius:50%;display:grid;place-items:center;font:700 11px Inter;color:#fff;background:var(--ac);flex:none}\n.stk{display:flex;height:28px;border-radius:7px;overflow:hidden;background:var(--graybg)}\n.stk div{display:flex;align-items:center;justify-content:center;font:700 12px Inter;color:#fff;cursor:pointer;min-width:26px;transition:filter .15s}\n.stk div:hover{filter:brightness(.92)}.stk div.on{box-shadow:inset 0 0 0 2px #18212d}\n.ptot{text-align:right;font-weight:700;color:var(--tx2)}\n.lgd{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--tx2);padding:10px 16px 14px;border-top:1px solid var(--bd2)}\n.lgd i{display:inline-block;width:10px;height:10px;border-radius:3px;background:var(--c);margin-right:6px;vertical-align:-1px}\n/* bảng */\ntable.t{width:100%;border-collapse:collapse}\n.t th{text-align:left;font-size:11.5px;font-weight:600;color:var(--tx3);padding:10px 12px;border-bottom:1px solid var(--bd);white-space:nowrap;background:#fafbfc;position:sticky;top:0;cursor:default}\n.t th[data-o]{cursor:pointer}.t th[data-o]:hover{color:var(--navy)}\n.t td{padding:10px 12px;border-bottom:1px solid var(--bd2);vertical-align:middle}\n.t tr:hover td{background:#f8faff;cursor:pointer}\n.sj{font-weight:600;color:var(--tx);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}\n.co{font:11.5px ui-monospace,\"SF Mono\",Consolas,monospace;color:var(--tx3);margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:520px}\n.badge{display:inline-block;font:600 11.5px Inter;padding:3px 9px;border-radius:999px;white-space:nowrap;background:var(--b);color:var(--f)}\n.ty{display:inline-block;font:700 11px Inter;padding:2px 8px;border-radius:6px;background:var(--b);color:var(--f);white-space:nowrap}\n.who{display:inline-flex;align-items:center;gap:7px;white-space:nowrap}.who .av{width:22px;height:22px;font-size:9.5px}\n.aib{border:1px solid #d8cff8;background:#f5f2ff;color:#5b3fd1;border-radius:7px;padding:3px 9px;font:600 11.5px Inter;cursor:pointer;white-space:nowrap}\n.aib:hover{background:#ebe5ff}\n.empty{color:var(--tx3);text-align:center;padding:26px}\n.t td:first-child,.t td:first-child div{white-space:nowrap}.grid2>.box{overflow:auto}.co{max-width:380px}\n/* theo bước */\n.board{display:grid;grid-template-columns:repeat(5,minmax(230px,1fr));gap:14px;min-width:1180px;margin-top:16px}\n.col{background:#eef1f5;border-radius:12px;display:flex;flex-direction:column}\n.ch{display:flex;align-items:center;gap:8px;padding:12px 14px;font-weight:700;font-size:13px;border-top:4px solid var(--c);border-radius:12px 12px 0 0;background:var(--card);border-bottom:1px solid var(--bd)}\n.ch .cnt{margin-left:auto;font-size:12px;font-weight:700;color:var(--tx2);background:var(--bg);border-radius:999px;padding:1px 9px}\n.cl{padding:10px;display:flex;flex-direction:column;gap:10px}\n.cd{background:var(--card);border:1px solid var(--bd);border-left:4px solid var(--sc);border-radius:10px;padding:11px 12px;cursor:pointer}\n.cd:hover{box-shadow:0 4px 14px rgba(16,24,40,.08)}\n.cd .r1{display:flex;align-items:center;gap:6px;margin-bottom:6px}.cd .r1 .badge{margin-left:auto}\n.cd .r3{display:flex;align-items:center;gap:8px;margin-top:9px;font-size:12px;color:var(--tx2)}\n.more{font-size:12px;color:var(--tx3);text-align:center;padding:4px}\n/* AI */\n.dr{position:fixed;top:0;right:0;bottom:0;width:440px;max-width:100%;background:var(--card);border-left:1px solid var(--bd);box-shadow:-16px 0 40px rgba(16,24,40,.14);transform:translateX(105%);transition:transform .25s;z-index:20;display:flex;flex-direction:column}\n.dr.on{transform:none}\n.dh{display:flex;align-items:center;gap:8px;padding:14px 16px;border-bottom:1px solid var(--bd);font-weight:700}\n.dh button{margin-left:auto;border:0;background:none;color:var(--tx2);font-size:22px;cursor:pointer;line-height:1}\n.db{flex:1;overflow:auto;padding:16px}\n.sec{margin-top:16px}.sec h5{margin:0 0 6px;font-size:11.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--tx3)}\n.sec p,.sec li{margin:0;font-size:13.5px;line-height:1.6}.sec ul{margin:0;padding-left:18px}\n.draft{white-space:pre-wrap;background:var(--bg);border:1px solid var(--bd);border-radius:10px;padding:12px;font-size:13px}\n.cp{margin-top:8px;border:1px solid var(--bd);background:var(--card);border-radius:8px;padding:5px 12px;font:600 12px Inter;cursor:pointer}\n.note{margin-top:16px;font-size:12px;color:#8a5a00;background:var(--amberbg);border-radius:9px;padding:9px 11px}\n.sk{height:12px;border-radius:6px;background:linear-gradient(90deg,#eef1f5,#e2e7ee,#eef1f5);background-size:200% 100%;animation:sh 1.2s linear infinite;margin:10px 0}\n@keyframes sh{to{background-position:-200% 0}}\n.load{padding:40px;color:var(--tx3);text-align:center}.err{padding:14px;color:var(--red);background:var(--redbg);border-radius:9px}\n@media (max-width:1100px){.kpis{grid-template-columns:repeat(3,1fr)}.grid2{grid-template-columns:1fr}.srch{width:180px}}\n@media (prefers-reduced-motion:reduce){*{animation:none !important;transition:none !important}}\n</style></head><body>\n<div id=\"app\"><div class=\"load\">Đang đọc Register…</div></div>\n<div class=\"dr\" id=\"dr\"><div class=\"dh\"><span style=\"color:#5b3fd1\">✦</span><span id=\"drT\">Trợ lý AI</span><button onclick=\"closeDr()\" aria-label=\"Đóng\">×</button></div><div class=\"db\" id=\"drB\"></div></div>\n<script>\nvar D=null,S={view:'ctl',step:'',bic:'',sla:'',q:'',sort:'late'};\nvar STEP=[{c:'#9aa5b4',n:'Chưa trình'},{c:'#2f6fdb',n:'Đã trình – chờ phản hồi'},{c:'#7b5cd6',n:'Chờ CĐT / TVTK'},{c:'#d98a00',n:'Trả về – trình lại'},{c:'#1f9d55',n:'Đã có kết quả'}];\nvar SLA={QUA_HAN:{c:'#d93a3a',b:'#fdeeee',n:'Quá hạn'},SAP_HAN:{c:'#d98a00',b:'#fff6e5',n:'Đến hạn ≤ 2 ngày'},CON_HAN:{c:'#0f9d8f',b:'#e6f6f4',n:'Còn hạn'},KHONG_HAN:{c:'#9aa5b4',b:'#f1f3f6',n:'Chưa có hạn'},XONG:{c:'#1f9d55',b:'#e8f6ee',n:'Đã xong'}};\nvar SLAO=['QUA_HAN','SAP_HAN','CON_HAN','KHONG_HAN','XONG'];\nvar TY={RFI:['#e6f6f4','#0b7a6f'],RFA:['#e9f0fc','#2257b8'],RFC:['#f1edfc','#6342c0'],SDS:['#e7f3fb','#0c6aa6'],MSS:['#fff1e6','#b15a00'],MAS:['#fcebf3','#b02a6b'],OTS:['#f1f3f6','#4d5a6b'],SCH:['#eef7e6','#4f7a12']};\nvar PC={'Nhà thầu':'#e07a1f','TVGS':'#2f6fdb','CĐT':'#7b5cd6','TBCONS':'#0f9d8f','—':'#1f9d55'};\nfunction esc(s){return String(s==null?'':s).replace(/[&<>\"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]})}\nfunction party(b){return String(b||'').replace(/ \\(.*$/,'')}\nfunction person(b){var m=/\\((.+)\\)/.exec(b||'');return m?m[1]:''}\nfunction pc(p){return PC[p]||'#64748b'}\nfunction ini(p){if(p==='—')return '✓';return String(p).split(/\\s+/).filter(Boolean).slice(-2).map(function(w){return w[0]}).join('').toUpperCase().slice(0,2)}\nfunction all(){var a=[];D.steps.forEach(function(s,i){D.cols[s].forEach(function(x){x.si=i;a.push(x)})});return a}\nfunction pass(x,skip){skip=skip||{};\n if(!skip.step&&S.step!==''&&x.si!==+S.step)return false;if(!skip.bic&&S.bic&&party(x.bic)!==S.bic)return false;if(!skip.sla&&S.sla&&x.sla!==S.sla)return false;\n if(S.q){var q=S.q.toLowerCase();if((x.so+' '+x.nd+' '+x.bic+' '+x.loai).toLowerCase().indexOf(q)<0)return false}return true}\nfunction dueTxt(x){if(x.sla==='XONG')return 'Đã xong';if(x.d==='')return 'Chưa có hạn';if(x.d<0)return 'Trễ '+(-x.d)+' ngày';if(x.d===0)return 'Hạn hôm nay';return 'Còn '+x.d+' ngày'}\nfunction sbadge(x){var s=SLA[x.sla];return '<span class=\"badge\" style=\"--b:'+s.b+';--f:'+s.c+'\">'+esc(dueTxt(x))+'</span>'}\nfunction tbadge(t){var c=TY[t]||['#f1f3f6','#4d5a6b'];return '<span class=\"ty\" style=\"--b:'+c[0]+';--f:'+c[1]+'\">'+esc(t||'—')+'</span>'}\nfunction whoH(b){var p=party(b);if(p==='—')return '<span class=\"who\" style=\"color:#1f9d55\">✓ Không còn ai giữ</span>';var n=person(b);return '<span class=\"who\"><span class=\"av\" style=\"--ac:'+pc(p)+'\">'+esc(ini(p))+'</span><span><b style=\"font-weight:600\">'+esc(p)+'</b>'+(n?'<br><span style=\"font-size:11.5px;color:#8792a2\">'+esc(n)+'</span>':'')+'</span></span>'}\nfunction row(x,cols){var h='<tr data-at=\"'+esc(x.at)+'\">';\n if(cols.due)h+='<td>'+sbadge(x)+(x.han?'<div style=\"font-size:11.5px;color:#8792a2;margin-top:3px\">hạn '+esc(x.han)+'</div>':'')+'</td>';\n if(cols.step)h+='<td style=\"white-space:nowrap;color:'+STEP[x.si].c+';font-weight:600\">'+(x.si+1)+'. '+STEP[x.si].n+'</td>';\n h+='<td>'+tbadge(x.loai)+'</td><td><div class=\"sj\">'+esc(x.nd||'(chưa có nội dung)')+'</div><div class=\"co\" title=\"'+esc(x.so)+'\">'+esc(x.so)+(x.rev!==''?' · Rev '+esc(x.rev):'')+'</div></td><td>'+whoH(x.bic)+'</td>';\n if(cols.y)h+='<td style=\"max-width:280px;color:#4d5a6b;font-size:12.5px\">'+esc(x.y||'')+'</td>';\n return h+'<td><button class=\"aib\" data-ai=\"'+esc(x.at)+'\">✦ AI đọc</button></td></tr>'}\nfunction draw(){\n var A=all(),F=A.filter(function(x){return pass(x)}),open=A.filter(function(x){return x.sla!=='XONG'});\n var h='<div class=\"hd\"><div class=\"lg\">T<b>Q</b>N</div><div class=\"tt\"><b>Kiểm soát hồ sơ trình duyệt</b><span>'+D.total+' hồ sơ (bản Rev mới nhất) · cập nhật '+esc(D.at)+'</span></div><div class=\"sp\"></div>'+\n  '<div class=\"srch\"><svg width=\"14\" height=\"14\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"#8792a2\" stroke-width=\"2\"><circle cx=\"11\" cy=\"11\" r=\"7\"/><path d=\"M20 20l-3.5-3.5\"/></svg><input id=\"q\" placeholder=\"Tìm số hiệu, nội dung…\" value=\"'+esc(S.q)+'\"></div>'+\n  '<button class=\"btn-ai\" id=\"aiAll\">✦ AI rà hồ sơ cần xử lý</button>'+\n  '<div class=\"seg\"><button data-v=\"ctl\" class=\"'+(S.view==='ctl'?'on':'')+'\">Tổng quan</button><button data-v=\"board\" class=\"'+(S.view==='board'?'on':'')+'\">Theo bước</button><button data-v=\"list\" class=\"'+(S.view==='list'?'on':'')+'\">Danh sách</button></div></div><div class=\"main\">';\n h+='<div class=\"kpis\">';\n D.steps.forEach(function(s,i){h+='<div class=\"kp'+(S.step===String(i)?' on':'')+'\" data-s=\"'+i+'\" style=\"--c:'+STEP[i].c+'\"><div class=\"l\">'+(i+1)+'. '+STEP[i].n+'</div><div class=\"n\">'+D.cols[s].length+'</div></div>'});\n var nl=A.filter(function(x){return x.sla==='QUA_HAN'}).length;\n h+='<div class=\"kp alert'+(S.sla==='QUA_HAN'?' on':'')+'\" data-q=\"1\" style=\"--c:#d93a3a\"><div class=\"l\">⚠ Quá hạn – cần xử lý</div><div class=\"n\">'+nl+'</div></div></div>';\n var tg=[];if(S.step!=='')tg.push(['step','Bước '+(+S.step+1)+': '+STEP[+S.step].n]);if(S.bic)tg.push(['bic','Bên giữ: '+S.bic]);if(S.sla)tg.push(['sla',SLA[S.sla].n]);if(S.q)tg.push(['q','Tìm: '+S.q]);\n h+=tg.length?'<div class=\"flt\">Đang lọc '+tg.map(function(t){return '<span class=\"tag\">'+esc(t[1])+'<button data-x=\"'+t[0]+'\" aria-label=\"Bỏ lọc\">×</button></span>'}).join('')+'<span>→ <b>'+F.length+'</b> hồ sơ</span></div>':'<div class=\"hint\">Bấm vào thẻ số hoặc dòng để lọc · bấm hồ sơ để mở đúng dòng trong sổ</div>';\n if(S.view==='ctl'){\n  h+='<div class=\"grid2\" style=\"grid-template-columns:1fr\">';\n  var od=F.filter(function(x){return x.sla==='QUA_HAN'||x.sla==='SAP_HAN'||x.si===3}).sort(function(a,b){return SLAO.indexOf(a.sla)-SLAO.indexOf(b.sla)||(a.d===''?0:a.d)-(b.d===''?0:b.d)});\n  h+='<div class=\"box\"><div class=\"bh\"><b>Cần xử lý trước</b><span>quá hạn, sắp đến hạn, bị trả về · '+od.length+' hồ sơ</span></div>';\n  h+=od.length?'<table class=\"t\"><thead><tr><th>Hạn</th><th>Loại</th><th>Hồ sơ</th><th>Đang giữ</th><th></th></tr></thead><tbody>'+od.slice(0,15).map(function(x){return row(x,{due:1})}).join('')+'</tbody></table>':'<div class=\"empty\">✓ Không có hồ sơ quá hạn hay sắp đến hạn</div>';\n  h+='</div></div>';\n }else if(S.view==='board'){h+='<div class=\"board\">';\n  D.steps.forEach(function(s,i){var it=F.filter(function(x){return x.si===i}).sort(function(a,b){return SLAO.indexOf(a.sla)-SLAO.indexOf(b.sla)||(+b.cho||0)-(+a.cho||0)});\n   h+='<div class=\"col\"><div class=\"ch\" style=\"--c:'+STEP[i].c+'\">'+(i+1)+'. '+STEP[i].n+'<span class=\"cnt\">'+it.length+'</span></div><div class=\"cl\">';\n   h+=it.length?it.slice(0,40).map(function(x){return '<div class=\"cd\" style=\"--sc:'+SLA[x.sla].c+'\" data-at=\"'+esc(x.at)+'\"><div class=\"r1\">'+tbadge(x.loai)+(x.rev!==''?'<span style=\"font-size:11.5px;color:#8792a2\">Rev '+esc(x.rev)+'</span>':'')+sbadge(x)+'</div><div class=\"sj\">'+esc(x.nd||'(chưa có nội dung)')+'</div><div class=\"co\">'+esc(x.so)+'</div><div class=\"r3\">'+whoH(x.bic)+'<span class=\"sp\"></span><button class=\"aib\" data-ai=\"'+esc(x.at)+'\">✦ AI</button></div></div>'}).join(''):'<div class=\"empty\">Không có hồ sơ</div>';\n   if(it.length>40)h+='<div class=\"more\">+ '+(it.length-40)+' hồ sơ · xem Danh sách</div>';h+='</div></div>'});\n  h+='</div>';\n }else{\n  var key={late:function(x){return SLAO.indexOf(x.sla)*1e4+(x.d===''?5000:x.d)},step:function(x){return x.si},loai:function(x){return String(x.loai)}}[S.sort];\n  F.sort(function(a,b){var A1=key(a),B1=key(b);return A1<B1?-1:A1>B1?1:0});\n  h+='<div class=\"box\" style=\"margin-top:16px;overflow:auto\"><table class=\"t\" style=\"min-width:1050px\"><thead><tr><th data-o=\"late\">Hạn ↕</th><th data-o=\"step\">Bước ↕</th><th data-o=\"loai\">Loại ↕</th><th>Hồ sơ</th><th>Đang giữ</th><th>Ý kiến gần nhất</th><th></th></tr></thead><tbody>'+\n   F.map(function(x){return row(x,{due:1,step:1,y:1})}).join('')+'</tbody></table>'+(F.length?'':'<div class=\"empty\">Không có hồ sơ khớp bộ lọc</div>')+'</div>';\n }\n document.getElementById('app').innerHTML=h+'</div>';\n var q=document.getElementById('q');q.oninput=function(){S.q=q.value;var p=q.selectionStart;draw();var n=document.getElementById('q');n.focus();n.setSelectionRange(p,p)};\n each('.seg button',function(b){b.onclick=function(){S.view=b.getAttribute('data-v');draw()}});\n each('.kp[data-s]',function(b){b.onclick=function(){var v=b.getAttribute('data-s');S.step=S.step===v?'':v;draw()}});\n each('.kp[data-q]',function(b){b.onclick=function(){S.sla=S.sla==='QUA_HAN'?'':'QUA_HAN';draw()}});\n each('.stk div[data-k]',function(c){c.onclick=function(){var p=c.getAttribute('data-p'),k=c.getAttribute('data-k');if(S.bic===p&&S.sla===k){S.bic='';S.sla=''}else{S.bic=p;S.sla=k}draw()}});\n each('.tag button',function(b){b.onclick=function(){S[b.getAttribute('data-x')]='';draw()}});\n each('th[data-o]',function(b){b.onclick=function(){S.sort=b.getAttribute('data-o');draw()}});\n each('[data-at]',function(c){c.onclick=function(){google.script.run.tqnGoto(c.getAttribute('data-at'))}});\n each('[data-ai]',function(b){b.onclick=function(e){e.stopPropagation();aiOne(b.getAttribute('data-ai'))}});\n document.getElementById('aiAll').onclick=aiAll;\n}\nfunction each(sel,f){Array.prototype.forEach.call(document.querySelectorAll(sel),f)}\nvar LV={CAO:['#d93a3a','#fdeeee','Mức độ cao'],TRUNG_BINH:['#b36f00','#fff6e5','Mức độ trung bình'],THAP:['#1f9d55','#e8f6ee','Mức độ thấp']};\nfunction openDr(t,h){document.getElementById('drT').textContent=t;document.getElementById('drB').innerHTML=h;document.getElementById('dr').classList.add('on')}\nfunction closeDr(){document.getElementById('dr').classList.remove('on')}\nfunction aiHtml(o){var lv=LV[o.muc_do]||LV.TRUNG_BINH;\n return '<div class=\"co\" style=\"max-width:none\">'+esc(o.so||'')+'</div><div style=\"margin-top:10px;display:flex;gap:8px;flex-wrap:wrap\"><span class=\"badge\" style=\"--b:'+lv[1]+';--f:'+lv[0]+'\">'+lv[2]+'</span><span class=\"badge\" style=\"--b:#eaf0fb;--f:#1B3F7A\">Bên xử lý: '+esc(o.ben_xu_ly||'?')+'</span></div>'+\n '<div class=\"sec\"><h5>Tóm tắt</h5><p>'+esc(o.tom_tat)+'</p></div><div class=\"sec\"><h5>Việc tiếp theo</h5><p>'+esc(o.viec_tiep_theo)+'</p></div>'+\n (o.can_bo_sung&&o.can_bo_sung.length?'<div class=\"sec\"><h5>Cần bổ sung / sửa</h5><ul>'+o.can_bo_sung.map(function(x){return '<li>'+esc(x)+'</li>'}).join('')+'</ul></div>':'')+\n (o.nhap_tra_loi?'<div class=\"sec\"><h5>Nháp trả lời</h5><div class=\"draft\">'+esc(o.nhap_tra_loi)+'</div><button class=\"cp\" onclick=\"cp(this)\">Chép nháp</button></div>':'')+\n '<div class=\"note\">Bản nháp do AI soạn từ nội dung trong Register ('+esc(o.model||'AI')+'), đã lưu vào tab TQN_AI. Người phụ trách kiểm tra trước khi dùng.</div>'}\nfunction cp(b){var t=b.previousElementSibling.textContent;try{navigator.clipboard.writeText(t).then(function(){b.textContent='Đã chép'},function(){b.textContent='Bôi đen để chép'})}catch(e){b.textContent='Bôi đen để chép'}}\nfunction wait(t){return '<div style=\"color:#4d5a6b\">'+esc(t)+'</div><div class=\"sk\"></div><div class=\"sk\" style=\"width:82%\"></div><div class=\"sk\" style=\"width:64%\"></div>'}\nfunction aiOne(at){openDr('Trợ lý AI đọc nội dung',wait('AI đang đọc nội dung và ý kiến các bên…'));\n google.script.run.withSuccessHandler(function(o){openDr('Trợ lý AI',aiHtml(o))}).withFailureHandler(function(e){openDr('Trợ lý AI','<div class=\"err\">'+esc(e.message||e)+'</div>')}).tqnAiAnalyze(at)}\nfunction aiAll(){var it=all().filter(function(x){return x.sla==='QUA_HAN'||x.si===3}).slice(0,10);if(!it.length){openDr('AI rà hồ sơ','<div class=\"empty\">Không có hồ sơ quá hạn hay bị trả về.</div>');return}\n var out=[],i=0;function next(){if(i>=it.length){openDr('AI đã đọc '+it.length+' hồ sơ',out.map(function(o){return o.err?'<div class=\"err\" style=\"margin-bottom:10px\">'+esc(o.err)+'</div>':'<div style=\"border-bottom:1px solid #e3e8ef;padding-bottom:16px;margin-bottom:16px\">'+aiHtml(o)+'</div>'}).join(''));return}\n  openDr('AI rà hồ sơ cần xử lý',wait('Đang đọc hồ sơ '+(i+1)+'/'+it.length+': '+it[i].nd));\n  google.script.run.withSuccessHandler(function(o){out.push(o);i++;next()}).withFailureHandler(function(e){out.push({err:it[i].so+': '+(e.message||e)});i++;if(/Chưa bật AI|chưa được dùng AI|Chưa cho phép/.test(e.message||''))i=it.length;next()}).tqnAiAnalyze(it[i].at)}\n next()}\ngoogle.script.run.withSuccessHandler(function(d){D=d;draw()}).withFailureHandler(function(e){document.getElementById('app').innerHTML='<div class=\"err\" style=\"margin:20px\">'+esc(e.message||e)+'</div>'}).tqnWorkflowData();\n</script></body></html>\n", "NewDoc": "<!DOCTYPE html>\n<html><head><base target=\"_top\"><meta charset=\"utf-8\">\n<style>\n:root{--navy2:#1B3F7A;--teal:#14837f;--teal2:#2DD4BF;--line:#e4e8ee;--tx:#1d2733;--tx2:#5b6675;--tx3:#8a94a3;--bg:#f5f7fa}\n*{box-sizing:border-box}body{margin:0;font:13px/1.45 Inter,\"Segoe UI\",Roboto,Arial,sans-serif;color:var(--tx);background:#fff}\n.seg{display:flex;margin:14px 18px 4px;background:var(--bg);border-radius:10px;padding:3px}\n.seg button{flex:1;border:0;background:none;padding:8px;border-radius:8px;font:600 13px inherit;color:var(--tx2);cursor:pointer}\n.seg button.on{background:#fff;color:var(--navy2);box-shadow:0 1px 3px rgba(0,0,0,.08)}\n.p{padding:6px 18px 16px;display:none}.p.on{display:block}\n.g{display:grid;grid-template-columns:1fr 1fr;gap:10px 12px}.full{grid-column:1/-1}\nlabel{display:block;font-size:11.5px;font-weight:600;color:var(--tx2);margin:10px 0 4px}\ninput,select,textarea{width:100%;padding:8px 10px;border:1px solid #d3d9e1;border-radius:8px;font:13px inherit;color:var(--tx);background:#fff;transition:border-color .15s,box-shadow .15s}\ninput:focus,select:focus,textarea:focus{outline:0;border-color:var(--navy2);box-shadow:0 0 0 3px rgba(27,63,122,.12)}\ntextarea{height:64px;resize:vertical}\n.types{display:flex;gap:6px;flex-wrap:wrap}\n.types button{border:1px solid #d3d9e1;background:#fff;border-radius:8px;padding:6px 11px;font:700 12px inherit;color:var(--tx2);cursor:pointer}\n.types button.on{background:var(--navy2);border-color:var(--navy2);color:#fff}\n.types small{display:block;font:400 10px Inter;opacity:.8}\n.pv{margin-top:14px;border:1px solid var(--line);border-radius:12px;padding:12px 14px;background:linear-gradient(180deg,#f8fbff,#fff)}\n.pv .l{font-size:11px;color:var(--tx3);text-transform:uppercase;letter-spacing:.5px;font-weight:600}\n.pv .c{font:700 16px ui-monospace,\"SF Mono\",Consolas,monospace;color:var(--navy2);margin-top:4px;word-break:break-all}\n.pv .h{font-size:11.5px;color:var(--tx3);margin-top:4px}\n.ck{display:flex;gap:8px;align-items:flex-start;margin-top:12px;font-size:12.5px;color:var(--tx2)}.ck input{width:auto;margin-top:2px}\n.btn{margin-top:14px;width:100%;border:0;border-radius:10px;padding:11px;font:700 14px inherit;background:var(--navy2);color:#fff;cursor:pointer;transition:background .15s}\n.btn:hover{background:#163566}.btn:disabled{opacity:.6;cursor:default}\n.msg{margin-top:12px;border-radius:10px;padding:10px 12px;font-size:12.5px;display:none}\n.msg.ok{display:block;background:#e6f4ec;color:#1e7d4f}.msg.er{display:block;background:#fdecea;color:#c62828}.msg.wt{display:block;background:var(--bg);color:var(--tx2)}\n.msg b{font-family:ui-monospace,Consolas,monospace}\n</style></head><body>\n<div class=\"seg\"><button class=\"on\" data-p=\"a\">Hồ sơ mới</button><button data-p=\"b\">Trình lại (Rev + 1)</button></div>\n<div class=\"p on\" id=\"a\">\n <label>Loại hồ sơ</label>\n <div class=\"types\" id=\"types\"></div>\n <div class=\"g\">\n  <div id=\"wn\"><label>Nhóm (RFA)</label><input id=\"nhom\" value=\"COS\"></div>\n  <div id=\"wg\"><label>Gói</label><select id=\"goi\"><option value=\"NP\">NP – Nhà phố</option><option value=\"HTKT\">HTKT – Hạ tầng kỹ thuật</option></select></div>\n  <div class=\"full\"><label>Nội dung hồ sơ</label><textarea id=\"subject\" placeholder=\"Vd: Yêu cầu làm rõ thông tin – cao độ sàn tầng 2\"></textarea></div>\n  <div><label>Người thực hiện</label><input id=\"owner\" placeholder=\"Tên người lập\"></div>\n  <div><label>Ngày trình</label><input id=\"date\" type=\"date\"></div>\n </div>\n <div class=\"pv\"><div class=\"l\">Số hiệu sẽ cấp</div><div class=\"c\" id=\"pv\">—</div><div class=\"h\">Xem trước, chưa giữ số. Số chỉ được giữ khi bấm Cấp mã.</div></div>\n <label class=\"ck\"><input type=\"checkbox\" id=\"write\" checked><span>Ghi số hiệu và nội dung vào dòng trống tiếp theo của sổ chi tiết. Ô đã có dữ liệu sẽ không bị ghi đè.</span></label>\n <button class=\"btn\" id=\"issue\">Cấp mã</button>\n <div class=\"msg\" id=\"m1\"></div>\n</div>\n<div class=\"p\" id=\"b\">\n <label>Số hiệu lần trình trước</label><input id=\"from\" placeholder=\"Vd: 261116-COA-TBC-RFI-001 rev01\">\n <label>Nội dung (nếu thay đổi)</label><textarea id=\"subject2\"></textarea>\n <label class=\"ck\"><input type=\"checkbox\" id=\"write2\" checked><span>Ghi vào dòng trống tiếp theo của sổ chi tiết</span></label>\n <button class=\"btn\" id=\"revise\">Cấp mã trình lại</button>\n <div class=\"msg\" id=\"m2\"></div>\n</div>\n<script>\nvar T=[['RFI','Thông tin'],['RFA','Phê duyệt'],['RFC','Thay đổi'],['SDS','Bản vẽ'],['MSS','Biện pháp'],['MAS','Vật tư'],['OTS','Khác']],type='RFI';\nfunction $(i){return document.getElementById(i)}\nfunction esc(s){return String(s==null?'':s).replace(/[&<>]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;'}[c]})}\n$('types').innerHTML=T.map(function(t){return '<button data-t=\"'+t[0]+'\" class=\"'+(t[0]===type?'on':'')+'\">'+t[0]+'<small>'+t[1]+'</small></button>'}).join('');\nArray.prototype.forEach.call(document.querySelectorAll('#types button'),function(b){b.onclick=function(){type=b.getAttribute('data-t');Array.prototype.forEach.call(document.querySelectorAll('#types button'),function(x){x.classList.toggle('on',x===b)});pv()}});\nvar d=new Date();$('date').value=d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);\nArray.prototype.forEach.call(document.querySelectorAll('.seg button'),function(b){b.onclick=function(){\n Array.prototype.forEach.call(document.querySelectorAll('.seg button,.p'),function(x){x.classList.remove('on')});b.classList.add('on');$(b.getAttribute('data-p')).classList.add('on')}});\nfunction form(){return{type:type,nhom:$('nhom').value,goi:$('goi').value,subject:$('subject').value,owner:$('owner').value,date:$('date').value,write:$('write').checked}}\nfunction msg(id,cls,html){var m=$(id);m.className='msg '+cls;m.innerHTML=html}\nfunction pv(){$('wn').style.display=type==='RFA'?'':'none';$('wg').style.display=/SDS|MSS|MAS|OTS/.test(type)?'':'none';$('pv').textContent='…';\n google.script.run.withSuccessHandler(function(c){$('pv').textContent=c}).withFailureHandler(function(e){$('pv').textContent=e.message||e}).tqnPreviewCode(form())}\n['nhom','goi','date'].forEach(function(i){$(i).onchange=pv});pv();\n$('issue').onclick=function(){var b=this;b.disabled=true;msg('m1','wt','Đang cấp mã (có khóa chống trùng)…');\n google.script.run.withSuccessHandler(function(r){b.disabled=false;msg('m1','ok','✔ Đã cấp <b>'+esc(r.code)+'</b>'+(r.where?'<br>Đã ghi vào '+esc(r.where):'')+'<br>Lưu tại tab TQN_CAP_MA.');pv()})\n .withFailureHandler(function(e){b.disabled=false;msg('m1','er','✖ '+esc(e.message||e))}).tqnIssueCode(form())};\n$('revise').onclick=function(){var b=this;b.disabled=true;msg('m2','wt','Đang cấp…');\n google.script.run.withSuccessHandler(function(r){b.disabled=false;msg('m2','ok','✔ Số hiệu trình lại <b>'+esc(r.code)+'</b> (Rev '+esc(r.rev)+')'+(r.where?'<br>Đã ghi vào '+esc(r.where):''))})\n .withFailureHandler(function(e){b.disabled=false;msg('m2','er','✖ '+esc(e.message||e))}).tqnReviseCode({from:$('from').value,subject:$('subject2').value,write:$('write2').checked})};\n</script></body></html>\n", "Manager": "<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\"><base target=\"_top\">\n<style>\n:root{--navy:#1B3F7A;--teal:#0F6E66;--ink:#1E293B;--mut:#64748B;--line:#E2E8F0;--bg:#F6F8FB;--card:#fff;\n--red:#B91C1C;--redb:#FEE2E2;--org:#9A3412;--orgb:#FFEDD5;--amb:#92400E;--ambb:#FEF3C7;--grn:#166534;--grnb:#DCFCE7;--gry:#475569;--gryb:#F1F5F9;\n--nt:#1E3A8A;--ntb:#C7DCFB;--tv:#115E59;--tvb:#B7EDE0;--cdt:#8A4B0F;--cdtb:#FBDDB8;--tk:#9D174D;--tkb:#F7D0E2}\n*{box-sizing:border-box}body{margin:0;font:13px/1.45 \"Segoe UI\",Roboto,Arial,sans-serif;color:var(--ink);background:var(--bg)}\nheader{display:flex;align-items:center;gap:12px;padding:10px 16px;background:#fff;border-bottom:1px solid var(--line);position:sticky;top:0;z-index:5}\n.logo{width:34px;height:34px;border-radius:8px;background:var(--navy);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px;letter-spacing:.5px}\nh1{font-size:15px;margin:0;font-weight:600}.sub{font-size:11px;color:var(--mut)}\nnav{display:flex;gap:2px;margin-left:auto}nav a{padding:7px 12px;border-radius:7px;color:var(--mut);cursor:pointer;text-decoration:none;font-weight:500}\nnav a.on{color:var(--navy);background:#E8EEF8}nav a:hover{color:var(--ink)}nav .cnt{display:inline-block;min-width:16px;padding:0 4px;margin-left:4px;border-radius:8px;background:var(--redb);color:var(--red);font-size:10px;text-align:center}\nmain{padding:14px 16px 30px;max-width:1160px;margin:0 auto}.view{display:none}.view.on{display:block}\n.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.kpi{background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px 14px;cursor:pointer}\n.kpi span{font-size:12px;color:var(--mut)}.kpi b{display:block;font-size:28px;font-weight:600;line-height:1.2}.kpi:hover{border-color:#CBD5E1}\n.mini{display:flex;gap:16px;flex-wrap:wrap;margin:8px 2px 14px;font-size:12px;color:var(--mut)}.mini b{color:var(--ink);font-weight:600}\n.acts{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px}\nbutton{font:inherit;cursor:pointer;border:1px solid #CBD5E1;background:#fff;color:var(--ink);border-radius:8px;padding:7px 12px;display:inline-flex;align-items:center;gap:6px}\nbutton:hover{background:#F1F5F9}button.pri{background:var(--navy);border-color:var(--navy);color:#fff}button.pri:hover{background:#163566}\nbutton .n{width:18px;height:18px;border-radius:50%;background:#E8EEF8;color:var(--navy);font-size:11px;font-weight:700;display:inline-flex;align-items:center;justify-content:center}\nbutton.pri .n{background:#fff}button.sm{padding:4px 9px;font-size:12px}button:disabled{opacity:.5;cursor:default}\n.more{position:relative;margin-left:auto}.menu{display:none;position:absolute;right:0;top:110%;background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:0 8px 24px rgba(15,23,42,.12);padding:6px;min-width:230px;z-index:9}\n.menu.show{display:block}.menu button{display:flex;width:100%;border:0;justify-content:flex-start;padding:7px 10px}.menu hr{border:0;border-top:1px solid var(--line);margin:4px}\n.chips{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:6px}.chip{border:1px solid #CBD5E1;border-radius:99px;padding:3px 10px;font-size:12px;color:var(--mut);cursor:pointer;background:#fff;display:inline-flex;gap:6px;align-items:center}\n.chip.on{border-color:var(--navy);color:var(--navy);background:#E8EEF8}.chip i{width:8px;height:8px;border-radius:50%;display:inline-block}\n.sec{font-weight:600;font-size:13px;margin:0 8px 0 2px}\n.list{background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden}\n.row{display:grid;grid-template-columns:26px minmax(0,1fr) 130px 104px 132px;gap:10px;align-items:center;padding:9px 12px;border-top:1px solid var(--line)}\n.row:first-child{border-top:0}.row:hover{background:#FAFBFD}.row.hd{background:#F8FAFC;font-size:11px;color:var(--mut);font-weight:600;text-transform:none}\n.nd{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.ma{font-family:Consolas,monospace;font-size:11.5px;color:var(--mut);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.tag{font-size:11.5px;padding:3px 8px;border-radius:6px;text-align:center;white-space:nowrap;font-weight:500}\n.t-red{background:var(--redb);color:var(--red)}.t-org{background:var(--orgb);color:var(--org)}.t-amb{background:var(--ambb);color:var(--amb)}.t-grn{background:var(--grnb);color:var(--grn)}.t-gry{background:var(--gryb);color:var(--gry)}\n.p{font-size:11.5px;padding:2px 8px;border-radius:99px;font-weight:500;white-space:nowrap;display:inline-block;max-width:100%;overflow:hidden;text-overflow:ellipsis}\n.p-nt{background:var(--ntb);color:var(--nt)}.p-tv{background:var(--tvb);color:var(--tv)}.p-cdt{background:var(--cdtb);color:var(--cdt)}.p-tk{background:var(--tkb);color:var(--tk)}.p-x{background:var(--gryb);color:var(--gry)}\n.empty{padding:24px;text-align:center;color:var(--mut)}\n.tools{display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap}input[type=search],select{font:inherit;border:1px solid #CBD5E1;border-radius:8px;padding:7px 10px;background:#fff}input[type=search]{width:260px}\ntable.st{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden;margin-top:14px}\ntable.st th,table.st td{padding:7px 10px;border-top:1px solid var(--line);text-align:right;font-size:12px}table.st th{background:#F8FAFC;color:var(--mut);font-weight:600}\ntable.st td:first-child,table.st th:first-child{text-align:left}\n.legend{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-top:14px;font-size:11px;color:var(--mut)}\n.cfg td{padding:6px 10px;border-top:1px solid var(--line)}.cfg{width:100%;background:#fff;border:1px solid var(--line);border-radius:10px;border-collapse:collapse}\n.cfg td:first-child{color:var(--mut);width:34%}\n.log td{font-size:12px}\n#modal{display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:20;align-items:center;justify-content:center;padding:16px}#modal.show{display:flex}\n.box{background:#fff;border-radius:12px;max-width:900px;width:100%;max-height:92vh;display:flex;flex-direction:column;overflow:hidden}\n.box .bh{display:flex;align-items:center;gap:8px;padding:10px 14px;border-bottom:1px solid var(--line)}.box .bh b{flex:1}.box .bb{padding:14px;overflow:auto}\n#toast{position:fixed;left:50%;bottom:18px;transform:translateX(-50%);background:var(--ink);color:#fff;padding:9px 14px;border-radius:8px;font-size:12.5px;display:none;z-index:30;max-width:80%}\n.note{font-size:12px;color:var(--mut)}textarea{width:100%;min-height:200px;font:inherit;border:1px solid #CBD5E1;border-radius:8px;padding:10px}\n.cmp{width:100%;border-collapse:collapse;margin-top:8px}.cmp td,.cmp th{border-top:1px solid var(--line);padding:6px 8px;text-align:left;font-size:12px;vertical-align:top}\n.drop{border:2px dashed #CBD5E1;border-radius:10px;padding:22px;text-align:center;color:var(--mut);cursor:pointer}.drop:hover{border-color:var(--navy);color:var(--navy)}\n.spin{display:inline-block;width:14px;height:14px;border:2px solid #CBD5E1;border-top-color:var(--navy);border-radius:50%;animation:s .8s linear infinite}@keyframes s{to{transform:rotate(360deg)}}\n@media (max-width:760px){.row{grid-template-columns:22px minmax(0,1fr) 96px}.row>:nth-child(3),.row>:nth-child(5){display:none}input[type=search]{width:100%}.kpi b{font-size:22px}}\n</style></head><body>\n<header><div class=\"logo\">TQN</div><div><h1>Quản lý hồ sơ trình duyệt</h1><div class=\"sub\" id=\"at\">Đang tải…</div></div>\n<nav id=\"nav\"></nav></header>\n<main>\n<section class=\"view on\" id=\"v-hn\">\n <div class=\"kpis\">\n  <div class=\"kpi\" data-f=\"QUA_HAN\"><span>Quá hạn</span><b id=\"k1\" style=\"color:var(--red)\">–</b></div>\n  <div class=\"kpi\" data-f=\"SAP_HAN\"><span>Sắp hạn (≤ 2 ngày)</span><b id=\"k2\" style=\"color:var(--amb)\">–</b></div>\n  <div class=\"kpi\" data-f=\"CHO\"><span>Đang chờ TVGS / CĐT</span><b id=\"k3\">–</b></div>\n </div>\n <div class=\"mini\" id=\"mini\"></div>\n <div class=\"acts\">\n  <button class=\"pri\" onclick=\"act('nhap')\"><span class=\"n\">1</span>Nhập và kiểm</button>\n  <button onclick=\"act('phieu')\"><span class=\"n\">2</span>Tạo phiếu</button>\n  <button onclick=\"act('gn')\"><span class=\"n\">3</span>Giao nhận</button>\n  <button onclick=\"act('docphieu')\">✦ Đọc phiếu (AI)</button>\n  <div class=\"more\"><button onclick=\"event.stopPropagation();document.getElementById('menu').classList.toggle('show')\">Thêm ▾</button>\n   <div class=\"menu\" id=\"menu\">\n    <button onclick=\"act('mamoi')\">Cấp mã mới</button><button onclick=\"act('rev')\">Trình lại rev+1</button><hr>\n    <button onclick=\"act('inds')\">In danh sách đang lọc</button><button onclick=\"act('nhan')\">In nhãn gáy (dòng đã chọn)</button>\n    <button onclick=\"act('thumuc')\">Mở thư mục phiếu (Drive)</button><hr>\n    <button onclick=\"act('airaa')\">✦ AI rà hồ sơ cần xử lý</button><button onclick=\"act('bochon')\">Bỏ hết dấu chọn</button><button onclick=\"act('lammoi')\">Làm mới</button>\n   </div></div>\n </div>\n <div class=\"chips\"><span class=\"sec\">Cần xử lý trước</span><span id=\"chips\"></span><span class=\"note\" id=\"selN\"></span></div>\n <div class=\"list\" id=\"todo\"></div>\n</section>\n<section class=\"view\" id=\"v-hs\">\n <div class=\"tools\"><input type=\"search\" id=\"q\" placeholder=\"Tìm mã, nội dung, người…\" oninput=\"drawAll()\">\n  <select id=\"fb\" onchange=\"drawAll()\"><option value=\"\">Mọi bước</option><option value=\"1\">1. Chưa trình</option><option value=\"2\">2. Chờ phản hồi</option><option value=\"3\">3. Chờ CĐT / TVTK</option><option value=\"4\">4. Trả về</option><option value=\"5\">5. Đã có kết quả</option></select>\n  <span class=\"note\" id=\"cnt\"></span></div>\n <div class=\"list\" id=\"all\"></div>\n <table class=\"st\" id=\"stat\"></table>\n</section>\n<section class=\"view\" id=\"v-gn\"><div class=\"list\" id=\"gnl\"></div></section>\n<section class=\"view\" id=\"v-vb\"><div class=\"list\" id=\"vbl\"></div></section>\n<section class=\"view\" id=\"v-ch\">\n <p class=\"note\">Sửa các giá trị này ở tab <b>TQN_CONFIG</b> (cột B). Phiếu tạo sau đó sẽ dùng giá trị mới.</p>\n <table class=\"cfg\" id=\"cfg\"></table>\n <p class=\"sec\" style=\"margin-top:16px\">Nhật ký gần đây</p><table class=\"cfg log\" id=\"log\"></table>\n</section>\n<div class=\"legend\" id=\"legend\"></div>\n</main>\n<div id=\"modal\"><div class=\"box\"><div class=\"bh\"><b id=\"mt\"></b><span id=\"mx\"></span><button class=\"sm\" onclick=\"closeM()\">Đóng</button></div><div class=\"bb\" id=\"mb\"></div></div></div>\n<div id=\"toast\"></div>\n<script>\nvar D=null,TAB='hn',FIL={nhom:0,k:''},SEL={};\nvar TABS=[['hn','Hôm nay'],['hs','Hồ sơ'],['gn','Giao nhận'],['vb','Văn bản · NCR'],['ch','⚙']];\nvar GC=['', '#0F6E66','#2563EB','#D97706','#7C3AED'];\nfunction $(i){return document.getElementById(i)}\nfunction esc(s){return String(s==null?'':s).replace(/[&<>\"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]})}\nfunction toast(t,ms){var e=$('toast');e.textContent=t;e.style.display='block';clearTimeout(e._t);e._t=setTimeout(function(){e.style.display='none'},ms||3500)}\nfunction call(fn,args,ok,busy){if(busy)toast(busy,60000);var r=google.script.run.withSuccessHandler(function(x){if(busy)$('toast').style.display='none';ok&&ok(x)}).withFailureHandler(function(e){toast('Lỗi: '+(e&&e.message||e),7000)});r[fn].apply(r,args||[])}\nfunction openM(t,h,extra){$('mt').textContent=t;$('mb').innerHTML=h;$('mx').innerHTML=extra||'';$('modal').classList.add('show')}\nfunction closeM(){$('modal').classList.remove('show')}\nfunction nav(){$('nav').innerHTML=TABS.map(function(t){var c=t[0]==='vb'&&D?D.docs.filter(function(x){return x.tt!=='Đã phản hồi'}).length:0;return '<a class=\"'+(TAB===t[0]?'on':'')+'\" onclick=\"go(\\''+t[0]+'\\')\">'+t[1]+(c?'<span class=\"cnt\">'+c+'</span>':'')+'</a>'}).join('')}\nfunction go(t){TAB=t;document.querySelectorAll('.view').forEach(function(v){v.classList.toggle('on',v.id==='v-'+t)});nav()}\nfunction party(s){s=String(s||'');var c=/CĐT|chủ đầu tư/i.test(s)?'cdt':/TVTK/i.test(s)?'tk':/TVGS|giám sát/i.test(s)?'tv':/Nhà thầu|TBCONS/i.test(s)?'nt':'x';var n=c==='nt'?'Nhà thầu':(s.replace(/^Nhà thầu.*/,'Nhà thầu')||'—');return '<span class=\"p p-'+c+'\" title=\"'+esc(s)+'\">'+esc(n)+'</span>'}\nfunction due(x){\n if(x.sla==='QUA_HAN')return '<span class=\"tag t-red\">Quá '+(x.d!==''?Math.abs(x.d)+' ngày':'hạn')+'</span>';\n if(x.buoc===4)return '<span class=\"tag t-org\">Bị trả về</span>';\n if(x.sla==='SAP_HAN')return '<span class=\"tag t-amb\">'+(x.d===0?'Hạn hôm nay':'Còn '+x.d+' ngày')+'</span>';\n if(x.sla==='CON_HAN')return '<span class=\"tag t-gry\">Hạn '+esc(x.han)+'</span>';\n if(x.sla==='XONG')return '<span class=\"tag t-grn\">Đã xong</span>';\n if(x.buoc===1)return '<span class=\"tag t-gry\">'+(x.phieu?'Có phiếu':'Chưa trình')+'</span>';\n return '<span class=\"tag t-gry\">Chưa có hạn</span>'}\nvar NX={TAO_PHIEU:'Tạo phiếu',GIAO_NHAN:'Giao nhận',NHAC:'Soạn nhắc',TRINH_LAI:'Trình lại rev+1'};\nfunction rowH(x){var id=esc(x.so);return '<div class=\"row\"><input type=\"checkbox\" '+(SEL[x.so]?'checked':'')+' onchange=\"sel(this,\\''+id.replace(/'/g,\"\\\\'\")+'\\')\"><div><div class=\"nd\" title=\"'+esc(x.nd)+'\">'+esc(x.nd)+'</div><div class=\"ma\">'+esc(x.so)+(x.rev?' · rev '+esc(x.rev):'')+' · '+esc(x.loai)+'</div></div><div>'+party(x.buoc===5?'':x.bic)+'</div>'+due(x)+'<div>'+(x.next?'<button class=\"sm\" onclick=\"one(\\''+x.next+'\\',\\''+id.replace(/'/g,\"\\\\'\")+'\\')\">'+NX[x.next]+'</button>':'')+'</div></div>'}\nvar HD='<div class=\"row hd\"><span></span><span>Hồ sơ</span><span>Đang chờ ai</span><span>Hạn</span><span>Việc tiếp theo</span></div>';\nfunction sel(el,so){if(el.checked)SEL[so]=1;else delete SEL[so];selN()}\nfunction selN(){var n=Object.keys(SEL).length;$('selN').textContent=n?'· đã chọn '+n+' hồ sơ':''}\nvar PR={QUA_HAN:0,SAP_HAN:2,CON_HAN:3,KHONG_HAN:4,CHUA:5};\nfunction todoList(){return D.rows.filter(function(x){return x.buoc<5&&(!FIL.nhom||x.nhom===FIL.nhom)&&(!FIL.k||(FIL.k==='CHO'?(x.buoc===2||x.buoc===3):x.sla===FIL.k))})\n .sort(function(a,b){var pa=a.buoc===4?1:PR[a.sla],pb=b.buoc===4?1:PR[b.sla];return pa-pb||(a.d===''?999:a.d)-(b.d===''?999:b.d)})}\nfunction drawToday(){\n var k=D.kpi;$('k1').textContent=k.quaHan;$('k2').textContent=k.sapHan;$('k3').textContent=k.cho;\n $('mini').innerHTML='<span>Bị trả về <b>'+k.traVe+'</b></span><span>Chưa có phiếu <b>'+k.chuaPhieu+'</b></span><span>Đã xong <b>'+k.xong+'</b></span><span>Tổng <b>'+k.tong+'</b></span>'+(FIL.k?'<span><a href=\"#\" onclick=\"FIL.k=\\'\\';drawToday();return false\">× bỏ lọc ô số</a></span>':'');\n $('chips').innerHTML=['Tất cả'].concat(D.nhomTen.slice(1)).map(function(t,i){return '<span class=\"chip '+(FIL.nhom===i?'on':'')+'\" onclick=\"FIL.nhom='+i+';drawToday()\">'+(i?'<i style=\"background:'+GC[i]+'\"></i>':'')+t+(i?' '+D.nhom[i]:'')+'</span>'}).join(' ');\n var L=todoList();$('todo').innerHTML=L.length?HD+L.slice(0,120).map(rowH).join(''):'<div class=\"empty\">Không có hồ sơ nào cần xử lý trong nhóm này.</div>';selN()}\nfunction drawAll(){var q=($('q').value||'').toLowerCase(),b=$('fb').value;\n var L=D.rows.filter(function(x){return(!b||String(x.buoc)===b)&&(!q||(x.so+' '+x.nd+' '+x.nguoi+' '+x.bic).toLowerCase().indexOf(q)>=0)});\n $('cnt').textContent=L.length+' / '+D.rows.length+' hồ sơ';$('all').innerHTML=L.length?HD+L.map(rowH).join(''):'<div class=\"empty\">Không có hồ sơ khớp.</div>';\n var T={};D.rows.forEach(function(x){var t=T[x.loai]=T[x.loai]||[0,0,0,0,0,0];t[0]++;if(!x.phieu)t[1]++;else t[2]++;if(x.gn)t[3]++;if(x.buoc>=2)t[4]++;if(x.buoc===5)t[5]++});\n $('stat').innerHTML='<tr><th>Thống kê theo loại hồ sơ</th><th>Tổng</th><th>Chưa có phiếu</th><th>Đã có phiếu</th><th>Đã giao nhận</th><th>Đã trình</th><th>Đã có kết quả</th></tr>'+Object.keys(T).sort().map(function(k){return '<tr><td>'+esc(k)+'</td>'+T[k].map(function(v){return '<td>'+v+'</td>'}).join('')+'</tr>'}).join('')}\nfunction drawGN(){$('gnl').innerHTML=D.gn.length?'<div class=\"row hd\" style=\"grid-template-columns:200px 120px 70px minmax(0,1fr) 90px\"><span>Số phiếu giao nhận</span><span>Lúc</span><span>Số HS</span><span>Danh sách</span><span>PDF</span></div>'+D.gn.map(function(g){return '<div class=\"row\" style=\"grid-template-columns:200px 120px 70px minmax(0,1fr) 90px\"><span class=\"ma\" style=\"color:var(--ink)\">'+esc(g.so)+'</span><span>'+esc(g.luc)+'</span><span>'+g.n+'</span><span class=\"ma\">'+esc(String(g.ds).replace(/\\n/g,' · '))+'</span><span>'+(g.pdf?'<a href=\"'+esc(g.pdf)+'\" target=\"_blank\">Mở PDF</a>':'—')+'</span></div>'}).join(''):'<div class=\"empty\">Chưa có phiếu giao nhận. Chọn hồ sơ ở tab Hôm nay rồi bấm \"3 Giao nhận\".</div>'}\nfunction drawVB(){$('vbl').innerHTML=D.docs.length?'<div class=\"row hd\" style=\"grid-template-columns:90px minmax(0,1fr) 110px 120px\"><span>Loại</span><span>Văn bản</span><span>Phát hành</span><span>Tình trạng</span></div>'+D.docs.map(function(x){return '<div class=\"row\" style=\"grid-template-columns:90px minmax(0,1fr) 110px 120px\"><span><i style=\"display:inline-block;width:8px;height:8px;border-radius:50%;background:'+GC[x.nhom]+';margin-right:6px\"></i>'+esc(x.loai)+'</span><div><div class=\"nd\" title=\"'+esc(x.nd)+'\">'+esc(x.nd)+'</div><div class=\"ma\">'+esc(x.ma)+' · '+esc(x.tab)+'!'+x.row+(x.nguoi?' · '+esc(x.nguoi):'')+'</div></div><span>'+esc(x.ngay||'—')+'</span>'+(x.tt==='Đã phản hồi'?'<span class=\"tag t-grn\">Đã phản hồi</span>':'<span class=\"tag '+(x.cho>7?'t-red':'t-amb')+'\">Chờ '+(x.cho!==''?x.cho+' ngày':'')+'</span>')+'</div>'}).join(''):'<div class=\"empty\">Sổ CV đi / CV đến / Chỉ thị / NCR chưa có dòng nào.</div>'}\nfunction drawCfg(){var L={phieu_ben_gui:'Ban hành bởi – Công ty',phieu_gui_toi:'Gửi tới – Công ty',phieu_nguoi_lap:'Người lập',phieu_nguoi_nhan:'Người nhận mặc định',phieu_so_hop_dong:'Số hợp đồng',phieu_so_ngay_ph:'Số ngày chờ phản hồi',phieu_du_an:'Dự án',phieu_dia_chi:'Địa chỉ',phieu_thu_muc:'Thư mục Drive lưu PDF',code_TML:'Mẫu số phiếu giao nhận'};\n $('cfg').innerHTML=Object.keys(L).map(function(k){return '<tr><td>'+L[k]+'</td><td>'+(esc(D.cfg[k])||'<span class=\"note\">(trống)</span>')+'</td></tr>'}).join('')+'<tr><td>Trợ lý AI</td><td>'+(D.aiOn?'<span class=\"tag t-grn\">Đang bật</span>':'<span class=\"tag t-gry\">Chưa bật</span> <span class=\"note\">'+esc(D.aiWhy)+'</span>')+'</td></tr>';\n $('log').innerHTML=D.log.length?D.log.map(function(a){return '<tr><td style=\"width:90px\">'+esc(a[0])+'</td><td style=\"width:140px\">'+esc(a[2])+'</td><td>'+esc(a[3])+'</td></tr>'}).join(''):'<tr><td class=\"note\">Chưa có nhật ký.</td></tr>'}\nfunction legend(){var s=[['t-red','Quá hạn'],['t-org','Bị trả về / Không duyệt'],['t-amb','Sắp hạn'],['t-gry','Còn hạn / chưa trình'],['t-grn','Đã xong']],p=[['nt','Nhà thầu'],['tv','TVGS'],['tk','TVTK'],['cdt','CĐT']];\n $('legend').innerHTML='<span>Hạn:</span>'+s.map(function(x){return '<span class=\"tag '+x[0]+'\">'+x[1]+'</span>'}).join('')+'<span style=\"margin-left:10px\">Bên đang giữ việc:</span>'+p.map(function(x){return '<span class=\"p p-'+x[0]+'\">'+x[1]+'</span>'}).join('')}\nfunction render(d){D=d;$('at').textContent='Cập nhật '+d.at+' · đọc từ Register, không sửa dữ liệu gốc';nav();drawToday();drawAll();drawGN();drawVB();drawCfg();legend()}\nfunction load(msg,full){call('tqnManagerData',[!full],render,msg)}\nfunction chosen(){var s=Object.keys(SEL);if(!s.length){toast('Đánh dấu ô vuông ở đầu dòng hồ sơ trước.');return null}return s}\nfunction paper(t,html,sub){openM(t,'<iframe id=\"pf\" style=\"width:100%;height:68vh;border:1px solid var(--line);border-radius:8px;background:#f1f5f9\"></iframe>'+(sub?'<p class=\"note\">'+sub+'</p>':''),'<button class=\"sm pri\" onclick=\"var f=document.getElementById(\\'pf\\');f.contentWindow.focus();f.contentWindow.print()\">In / Lưu PDF</button>');\n var f=$('pf');f.srcdoc='<html><head><meta charset=\"utf-8\"></head><body style=\"margin:12px;background:#f1f5f9\">'+html+'</body></html>'}\nfunction one(kind,so){SEL={};SEL[so]=1;if(kind==='TAO_PHIEU')act('phieu');else if(kind==='GIAO_NHAN')act('gn');else if(kind==='NHAC')act('nhac',so);else if(kind==='TRINH_LAI')act('rev')}\nfunction act(a,arg){$('menu').classList.remove('show');var s;\n if(a==='nhap'){load('Đang đọc Register và kiểm tra…',true);setTimeout(function(){toast('Đã nhập và kiểm. Register gốc không bị sửa.')},400);return}\n if(a==='lammoi'){load('Đang làm mới…',true);return}\n if(a==='bochon'){SEL={};drawToday();drawAll();return}\n if(a==='phieu'){if(!(s=chosen()))return;call('tqnMakeForms',[s],function(r){SEL={};paper('Phiếu vừa tạo ('+r.n+')',r.html,'Đã ghi vào TQN_PHIEU'+(r.pdf?' và lưu '+r.pdf+' PDF vào thư mục Drive \"'+esc(D.cfg.phieu_thu_muc)+'\"':' (chưa lưu PDF vào Drive – bấm \"In / Lưu PDF\" để lưu)')+'.');load()},'Đang tạo phiếu…');return}\n if(a==='gn'){if(!(s=chosen()))return;call('tqnMakeTransmittal',[s],function(r){SEL={};paper('Phiếu giao nhận '+r.so,r.html,'Đã cấp số '+esc(r.so)+' cho '+r.n+' hồ sơ, ghi vào TQN_GIAO_NHAN.');load()},'Đang lập phiếu giao nhận…');return}\n if(a==='nhan'){if(!(s=chosen()))return;call('tqnSpineLabels',[s],function(h){paper('Nhãn gáy',h)});return}\n if(a==='inds'){var L=TAB==='hs'?D.rows.filter(function(x){var q=($('q').value||'').toLowerCase(),b=$('fb').value;return(!b||String(x.buoc)===b)&&(!q||(x.so+' '+x.nd).toLowerCase().indexOf(q)>=0)}):todoList();\n  paper('Danh sách ('+L.length+')','<table style=\"width:100%;border-collapse:collapse;font:12px Arial;background:#fff\"><tr style=\"background:#EEF1F5\"><th style=\"border:1px solid #999;padding:4px\">STT</th><th style=\"border:1px solid #999;padding:4px\">Số hiệu</th><th style=\"border:1px solid #999;padding:4px\">Nội dung</th><th style=\"border:1px solid #999;padding:4px\">Đang chờ</th><th style=\"border:1px solid #999;padding:4px\">Hạn</th><th style=\"border:1px solid #999;padding:4px\">Bước</th></tr>'+L.map(function(x,i){return '<tr><td style=\"border:1px solid #999;padding:4px;text-align:center\">'+(i+1)+'</td><td style=\"border:1px solid #999;padding:4px;font-family:monospace\">'+esc(x.so)+'</td><td style=\"border:1px solid #999;padding:4px\">'+esc(x.nd)+'</td><td style=\"border:1px solid #999;padding:4px\">'+esc(x.bic)+'</td><td style=\"border:1px solid #999;padding:4px\">'+esc(x.han)+'</td><td style=\"border:1px solid #999;padding:4px\">'+esc(x.buocTen)+'</td></tr>'}).join('')+'</table>');return}\n if(a==='thumuc'){call('tqnFolderUrl',[],function(u){if(u)window.open(u,'_blank');else toast('Chưa có thư mục phiếu (tạo phiếu đầu tiên sẽ tự tạo).')});return}\n if(a==='nhac'){call('tqnReminderText',[arg],function(t){openM('Nháp văn bản nhắc phản hồi','<p class=\"note\">Tool chỉ soạn nháp, không tự gửi. Sửa rồi chép vào email / công văn.</p><textarea id=\"rt\">'+esc(t)+'</textarea>','<button class=\"sm pri\" onclick=\"var e=document.getElementById(\\'rt\\');e.select();document.execCommand(\\'copy\\');toast(\\'Đã chép.\\')\">Chép</button>')});return}\n if(a==='mamoi'||a==='rev'){google.script.run.tqnMenuNewDoc();return}\n if(a==='airaa'){call('tqnAiBatch',[],function(r){toast(typeof r==='string'?r:'AI đã rà xong, xem tab TQN_AI.',6000)},'AI đang rà…');return}\n if(a==='docphieu'){docPhieu();return}}\n/* ---- Đọc phiếu bằng AI ---- */\nfunction docPhieu(){openM('Đọc phiếu bằng AI','<p class=\"note\">Chụp hoặc scan phiếu đã có ý kiến TVGS / CĐT (ảnh JPG/PNG hoặc PDF, tối đa ~5 MB). AI đọc số hiệu, kết quả, ngày, ý kiến rồi <b>đề xuất</b> ghi vào sổ; anh xem lại rồi mới bấm Ghi.</p>'+(D&&!D.aiOn?'<p class=\"tag t-amb\" style=\"display:block;text-align:left\">AI chưa bật: '+esc(D.aiWhy)+'</p>':'')+'<label class=\"drop\" for=\"fi\">Bấm để chọn ảnh / PDF phiếu<input type=\"file\" id=\"fi\" accept=\"image/*,application/pdf\" style=\"display:none\" onchange=\"sendForm(this.files[0])\"></label><div id=\"fr\"></div>')}\nfunction sendForm(f){if(!f)return;if(f.size>5*1024*1024){toast('File quá 5 MB, chụp lại nhỏ hơn.');return}\n var rd=new FileReader();rd.onload=function(){var b64=String(rd.result).split(',')[1];$('fr').innerHTML='<p><span class=\"spin\"></span> AI đang đọc phiếu…</p>';\n  google.script.run.withSuccessHandler(showForm).withFailureHandler(function(e){$('fr').innerHTML='<p class=\"tag t-red\" style=\"display:block;text-align:left\">'+esc(e.message||e)+'</p>'}).tqnAiReadForm(b64,f.type||'image/jpeg')};rd.readAsDataURL(f)}\nvar FORM=null;\nfunction showForm(r){FORM=r;var a=r.ai,tc={CAO:'t-grn',TRUNG_BINH:'t-amb',THAP:'t-red'}[a.do_tin_cay]||'t-gry';\n var h='<table class=\"cmp\"><tr><th style=\"width:30%\">AI đọc được</th><th></th></tr><tr><td>Số hiệu</td><td class=\"ma\" style=\"color:var(--ink)\">'+esc(a.so_hieu)+(a.rev?' · rev '+esc(a.rev):'')+'</td></tr><tr><td>Bên phản hồi</td><td>'+party(a.ben_phan_hoi)+'</td></tr><tr><td>Kết quả</td><td><b>'+esc(a.ket_qua)+'</b></td></tr><tr><td>Ngày</td><td>'+esc(a.ngay_phan_hoi||'—')+'</td></tr><tr><td>Người ký</td><td>'+esc(a.nguoi_ky||'—')+'</td></tr><tr><td>Ý kiến</td><td>'+esc(a.y_kien||'—')+'</td></tr><tr><td>Độ tin cậy</td><td><span class=\"tag '+tc+'\">'+esc(a.do_tin_cay)+'</span> <span class=\"note\">'+esc(a.ghi_chu)+'</span></td></tr></table>';\n if(!r.found){h+='<p class=\"tag t-amb\" style=\"display:block;text-align:left;margin-top:10px\">'+esc(r.msg)+'</p>';$('fr').innerHTML=h;return}\n h+='<p style=\"margin:14px 0 4px\"><b>Đề xuất ghi vào '+esc(r.at)+'</b> <span class=\"note\">'+esc(r.nd)+'</span></p><table class=\"cmp\"><tr><th></th><th>Ô</th><th>Mục</th><th>Sổ đang ghi</th><th>AI đề xuất</th></tr>'+r.de_xuat.map(function(p,i){return '<tr><td><input type=\"checkbox\" id=\"fp'+i+'\" '+(p.ghi?'checked':'disabled')+'></td><td class=\"ma\">'+esc(p.o)+'</td><td>'+esc(p.ten)+'</td><td>'+(p.cur?esc(p.cur):'<span class=\"note\">(trống)</span>')+'</td><td>'+esc(p.moi)+(p.ghi?'':' <span class=\"note\">– ô đã có, không đè</span>')+'</td></tr>'}).join('')+'</table>'\n  +'<div style=\"margin-top:12px;display:flex;gap:8px\"><button class=\"pri\" onclick=\"applyForm()\">Ghi các ô đã chọn</button><button onclick=\"closeM()\">Bỏ</button></div>';\n $('fr').innerHTML=h}\nfunction applyForm(){var it=FORM.de_xuat.filter(function(p,i){var c=$('fp'+i);return c&&c.checked});if(!it.length){toast('Chưa chọn ô nào.');return}\n call('tqnAiApplyForm',[FORM.sheet,FORM.row,it],function(r){closeM();toast('Đã ghi '+r.ghi.length+' ô'+(r.bo.length?', bỏ '+r.bo.length+' ô đã có dữ liệu':'')+'. Có nhật ký trong TQN_AUDIT.',6000);load()},'Đang ghi…')}\ndocument.querySelectorAll('.kpi').forEach(function(k){k.onclick=function(){FIL.k=FIL.k===k.dataset.f?'':k.dataset.f;go('hn');drawToday()}});\ndocument.addEventListener('click',function(){$('menu').classList.remove('show')});\n$('modal').onclick=function(e){if(e.target.id==='modal')closeM()};\nnav();load();\n</script></body></html>\n"};
