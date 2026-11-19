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
    var dangMo = null; try { dangMo = ss.getActiveSheet(); } catch (e) { }
    sh = ss.insertSheet(name);
    try { if (dangMo && dangMo.getName() !== name) ss.setActiveSheet(dangMo); } catch (e) { }   // tạo tab phụ không làm nhảy khỏi tab đang làm
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
  ['weather_hours', '6 11 13 18', 'Các giờ cập nhật trong ngày (sáng/trưa/chiều/tối, giống bản tin). Mỗi lần cập nhật lại đúng dòng hôm nay'],
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

function tqnWeatherUrl_(cfg, khoang) {
  var base = cfg.weather_api_key ? 'https://customer-api.open-meteo.com/v1/forecast' : 'https://api.open-meteo.com/v1/forecast';
  var q = {
    latitude: cfg.lat, longitude: cfg.lon, timezone: cfg.timezone, wind_speed_unit: 'kmh',
    daily: 'weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max,wind_gusts_10m_max,sunrise,sunset',
    hourly: 'relative_humidity_2m,temperature_2m,apparent_temperature,precipitation_probability,precipitation,weather_code,wind_gusts_10m,wind_speed_10m'
  };
  if (khoang) { q.start_date = khoang.tu; q.end_date = khoang.den; } else q.forecast_days = 2;
  if (cfg.weather_api_key) q.apikey = cfg.weather_api_key;
  return base + '?' + Object.keys(q).map(function (k) { return k + '=' + encodeURIComponent(q[k]); }).join('&');
}

/** Đọc số liệu của ngày thứ k trong phản hồi (k=0 là ngày đầu). */
function tqnWeatherDay_(j, k) {
  var d = j.daily || {}, h = j.hourly || {}, at = function (a, i) { return a && a[i] !== undefined ? a[i] : null; };
  var hh = {}; Object.keys(h).forEach(function (key) { hh[key] = (h[key] || []).slice(k * 24, k * 24 + 24); });
  var hum = (hh.relative_humidity_2m || []).filter(function (x) { return typeof x === 'number'; });
  var feel = (hh.apparent_temperature || []).filter(function (x) { return typeof x === 'number'; });
  return {
    date: d.time[k], tmin: at(d.temperature_2m_min, k), tmax: at(d.temperature_2m_max, k), code: at(d.weather_code, k),
    rain: at(d.precipitation_sum, k), pop: at(d.precipitation_probability_max, k),
    wind: at(d.wind_speed_10m_max, k), gust: at(d.wind_gusts_10m_max, k),
    hum: hum.length ? Math.round(hum.reduce(function (a, b) { return a + b; }, 0) / hum.length) : null,
    feel: feel.length ? Math.max.apply(null, feel) : null,
    buoi: tqnWeatherBuoi_(hh), gio: hh,
    mai: (d.time.length > k + 1 && at(d.temperature_2m_max, k + 1) !== null) ? { date: d.time[k + 1], tmin: d.temperature_2m_min[k + 1], tmax: d.temperature_2m_max[k + 1], code: d.weather_code[k + 1], rain: at(d.precipitation_sum, k + 1) } : null,
    sunrise: d.sunrise ? String(d.sunrise[k] || '').slice(11, 16) : '', sunset: d.sunset ? String(d.sunset[k] || '').slice(11, 16) : ''
  };
}

/** Gọi API có thử lại; trả về {ok, data (ngày đầu), j, raw, fetchedAt}. khoang = {tu,den} 'yyyy-MM-dd' để lấy lại ngày đã qua. */
function tqnWeatherFetch_(cfg, khoang) {
  var url = tqnWeatherUrl_(cfg, khoang), last = '';
  for (var i = 0; i < 3; i++) {
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = res.getResponseCode(), raw = res.getContentText();
      if (code === 200) {
        var j = JSON.parse(raw), d = j.daily || {};
        var need = ['time', 'temperature_2m_min', 'temperature_2m_max', 'weather_code'];
        for (var n = 0; n < need.length; n++) if (!d[need[n]] || d[need[n]][0] === undefined || d[need[n]][0] === null) throw new Error('Thiếu trường ' + need[n]);
        return { ok: true, raw: raw, j: j, fetchedAt: new Date(), data: tqnWeatherDay_(j, 0) };
      }
      last = 'HTTP ' + code;
      if (code >= 400 && code < 500 && code !== 429) break;   // lỗi tham số: thử lại vô ích
    } catch (e) { last = String(e.message || e); }
    Utilities.sleep(Math.pow(2, i) * 2000);
  }
  return { ok: false, err: last };
}

/** Chia 24 giờ hôm nay thành 3 buổi giống bản tin: Sáng 06–11, Chiều 12–17, Tối 18–23. */
var TQN_BUOI = [['Sáng', 6, 10], ['Trưa', 11, 13], ['Chiều', 14, 17], ['Tối', 18, 23]];
function tqnWeatherBuoi_(h) {
  var arr = function (k) { return (h[k] || []).slice(0, 24); };   // h đã là 24 giờ của đúng ngày
  var t = arr('temperature_2m'), pp = arr('precipitation_probability'), pr = arr('precipitation'), wc = arr('weather_code'), g = arr('wind_gusts_10m');
  if (!t.length) return [];
  var num = function (x) { return typeof x === 'number'; };
  return TQN_BUOI.map(function (b) {
    var r = function (a) { return a.slice(b[1], b[2] + 1).filter(num); };
    var tt = r(t), p = r(pp), m = r(pr), c = r(wc), gg = r(g);
    if (!tt.length) return null;
    var dem = {}, chinh = c[0];
    c.forEach(function (x) { dem[x] = (dem[x] || 0) + 1; if (dem[x] > (dem[chinh] || 0) || (dem[x] === dem[chinh] && x > chinh)) chinh = x; });
    return { ten: b[0], tu: b[1], den: b[2], tmin: Math.min.apply(null, tt), tmax: Math.max.apply(null, tt),
      pop: p.length ? Math.max.apply(null, p) : null, mua: m.length ? Math.round(m.reduce(function (a, x) { return a + x; }, 0) * 10) / 10 : null,
      code: chinh, gust: gg.length ? Math.max.apply(null, gg) : null };
  }).filter(Boolean);
}
function tqnSoVi_(x) { return String(Math.round(x * 10) / 10).replace('.', ','); }
/** Ghi chú dễ đọc, mỗi buổi một dòng. Dòng cuối ghi "Tự động" để lần sau nhận ra dòng do tool ghi. */
function tqnWeatherNote_(cfg, w) {
  var d = w.data, L = [];
  (d.buoi || []).forEach(function (b) {
    L.push(b.ten + ' ' + Math.round(b.tmin) + '–' + Math.round(b.tmax) + '°: ' + (TQN_WMO_VI[b.code] || 'mã ' + b.code).toLowerCase() +
      (b.pop !== null ? ', khả năng mưa ' + b.pop + '%' : '') + (b.mua ? ' (' + tqnSoVi_(b.mua) + ' mm)' : ''));
  });
  var x = [];
  if (d.gust !== null && d.gust !== undefined) x.push('Gió giật ' + Math.round(d.gust) + ' km/h');
  if (d.sunrise) x.push('mặt trời mọc ' + d.sunrise + ', lặn ' + d.sunset);
  if (x.length) L.push(x.join(' · '));
  if (d.mai) { var p = d.mai.date.split('-'); L.push('Ngày mai ' + p[2] + '/' + p[1] + ': ' + (TQN_WMO_VI[d.mai.code] || '').toLowerCase() + ' ' + Math.round(d.mai.tmin) + '–' + Math.round(d.mai.tmax) + '°' + (d.mai.rain ? ', mưa ' + tqnSoVi_(d.mai.rain) + ' mm' : '')); }
  if (!(d.buoi || []).length && d.pop !== null) L.unshift('Khả năng mưa ' + d.pop + '%');
  return L.join('\n');
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
        if (!/^\[TQN\]|Tự động – |^📷 /.test(String(notes[i][0]))) return 'BO_QUA_NHAP_TAY';
        row = i + 2; break;
      }
    }
  }
  var action = row ? 'CAP_NHAT' : 'THEM';
  if (!row) {
    row = Math.max(2, last + 1);
    if (last >= 2) {   // chèn đúng thứ tự ngày (lấy lại ngày cũ không bị dồn xuống cuối)
      for (var q = 0; q < dates.length; q++) if (tqnIsDate_(dates[q][0]) && dates[q][0] > day) { sh.insertRowBefore(q + 2); row = q + 2; break; }
    }
  }
  var r1 = function (x) { return (x === null || x === undefined) ? '' : Math.round(x * 10) / 10; };
  var d = w.data;
  var note = tqnWeatherNote_(cfg, w);
  var set = function (c, v) { if (c) sh.getRange(row, c).setValue(v); };
  // Ghi 1 lần cho cả dòng (nhanh hơn ~10 lần gọi): giữ nguyên giá trị các cột không thuộc tool
  var cMax = Math.max(m.date, m.thu, m.tmin, m.tmax, m.wx, m.rain, m.hum, m.wind, 1);
  var dong = sh.getRange(row, 1, 1, cMax).getValues()[0];
  var dat = function (c, v) { if (c) dong[c - 1] = v; };
  dat(m.date, day); dat(m.thu, TQN_THU[day.getDay()]); dat(m.tmin, r1(d.tmin)); dat(m.tmax, r1(d.tmax));
  dat(m.wx, TQN_WMO_VI[d.code] || ('Mã WMO ' + d.code)); dat(m.rain, r1(d.rain)); dat(m.hum, d.hum === null ? '' : d.hum); dat(m.wind, r1(d.wind));
  sh.getRange(row, 1, 1, cMax).setValues([dong]);
  set(m.note, note);
  var tom = (d.buoi || []).map(function (b) { return b.ten + ' ' + (b.pop === null ? '–' : b.pop + '%'); }).join(' · ');
  if (m.note && w.banTin) {
    try {
      var nhan = '📷 Xem bản tin hình (bấm mở)', full = nhan + '\n' + note;
      sh.getRange(row, m.note).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(full).setLinkUrl(0, nhan.length, w.banTin).build());
    } catch (e) { set(m.note, '📷 Bản tin: ' + w.banTin + '\n' + note); }
  } else if (m.note) set(m.note, '📷 Bản tin hình chưa tạo được\n' + note);
  try { if (m.note) sh.getRange(row, m.note).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); sh.getRange(row, 1, 1, m.note).setVerticalAlignment('middle'); if (m.note) sh.getRange(row, m.note).setWrap(true).setFontSize(9); sh.getRange(row, 1, 1, m.note).setVerticalAlignment('middle');
    if (sh.getColumnWidth(m.note) < 380) sh.setColumnWidth(m.note, 380);
    tqnWeatherIconCells_(sh, row, m.note + 1, d);
    sh.setRowHeightsForced(row, 1, 110); } catch (e) { tqnAudit_('THOI_TIET', 'Định dạng/hình lỗi: ' + e.message, {}); }
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
    var gio = parseInt(Utilities.formatDate(new Date(), cfg.timezone, 'HH'), 10), slot = gio < 11 ? '06' : gio < 13 ? '11' : gio < 18 ? '13' : '18';
    var key = cfg.project_id + ':' + Utilities.formatDate(new Date(), cfg.timezone, 'yyyy-MM-dd') + ':weather' + slot;
    var props = PropertiesService.getDocumentProperties();
    if (manual !== true && props.getProperty('lastWeather') === key) { tqnAudit_('THOI_TIET', 'Bỏ qua: hôm nay đã chạy ' + key, {}); return 'DA_CHAY'; }
    var w = tqnWeatherFetch_(cfg);
    if (!w.ok) { tqnWeatherRaw_(cfg, w, 'FAILED'); tqnAudit_('THOI_TIET', 'FAILED: ' + w.err, {}); return 'FAILED: ' + w.err; }
    w.banTin = tqnWeatherBanTin_(cfg, w);
    var act = tqnWeatherWrite_(cfg, w);
    try { tqnDinhDangThoiTiet(); } catch (e) { }
    tqnWeatherRaw_(cfg, w, act);
    props.setProperty('lastWeather', key);
    tqnAudit_('THOI_TIET', act + ' ' + w.data.date, w.data);
    return act;
  } finally { lock.releaseLock(); }
}

function tqnInstallWeather() {
  var cfg = tqnWeatherCfg_();
  tqnRemoveWeather(true);
  var raw = String(cfg.weather_hours || '').trim();
  var gios = /^\d{1,2}([\s,;]+\d{1,2})*$/.test(raw) ? raw.split(/[\s,;]+/).map(function (x) { return parseInt(x, 10); }).filter(function (x) { return x >= 0 && x < 24; }) : [];
  if (!gios.length) gios = [6, 11, 13, 18];
  gios.forEach(function (h) { ScriptApp.newTrigger('tqnWeatherJob').timeBased().atHour(h).nearMinute(0).everyDays(1).inTimezone(cfg.timezone).create(); });
  tqnAudit_('BAT_THOI_TIET', 'Hằng ngày khoảng ' + gios.join('h, ') + 'h (' + cfg.timezone + ')', {});
  return gios.join('h, ') + 'h';
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
  SpreadsheetApp.getUi().alert('Đã bật cập nhật thời tiết hằng ngày lúc ' + tqnInstallWeather() + ' (Google có thể lệch ±15 phút).');
}
function tqnMenuWeatherOff() { SpreadsheetApp.getUi().alert('Đã tắt ' + tqnRemoveWeather(false) + ' lịch thời tiết.'); }


/* ===== Bản tin hình (minh chứng) – giống mẫu bản tin TQN anh Nho gửi 19/11 ===== */
function tqnWxIcon_(c, dem) {
  if (c === null || c === undefined) return '';
  if (c >= 95) return '⛈'; if (c >= 51) return '🌧'; if (c >= 45) return '🌫'; if (c >= 2) return '☁'; return dem ? '☾' : '☀';
}
function tqnWeatherBanTinHtml_(cfg, w) {
  var d = w.data, p = d.date.split('-'), day = new Date(+p[0], +p[1] - 1, +p[2]), e = tqnEsc_;
  var g = d.gio || {}, T = g.temperature_2m || [], PP = g.precipitation_probability || [], WC = g.weather_code || [];
  var ic = function (c, dem, px) { return (c === null || c === undefined) ? '' : '<img src="' + TQN_WX_ICON[tqnWxIconKey_(c, dem)] + '" width="' + px + '" height="' + px + '">'; };
  var so = function (x) { return x === null || x === undefined ? '–' : Math.round(x); };
  // Câu tóm tắt như app điện thoại
  var bMax = null; (d.buoi || []).forEach(function (b) { if (b.pop !== null && (!bMax || b.pop > bMax.pop)) bMax = b; });
  var tom = bMax && bMax.pop >= 40 ? (TQN_WMO_VI[bMax.code] || 'Có mưa') + ' vào buổi ' + bMax.ten.toLowerCase() + '.' : 'Ít khả năng mưa, ' + (TQN_WMO_VI[d.code] || '').toLowerCase() + '.';
  tom += ' Cao ' + so(d.tmax) + '° và thấp ' + so(d.tmin) + '°C.' + (d.gust ? ' Gió giật đến ' + so(d.gust) + ' km/h.' : '');
  var khoi = function (ten, a, b) {
    var h1 = '', h2 = '', h3 = '', h4 = '';
    for (var i = a; i <= b; i++) {
      h1 += '<td class="hh">' + (i < 10 ? '0' : '') + i + ':00</td>';
      h2 += '<td class="ic">' + ic(WC[i], i >= 18 || i < 6, 34) + '</td>';
      h3 += '<td class="tt">' + so(T[i]) + '°</td>';
      h4 += '<td class="pp">💧' + (PP[i] === null || PP[i] === undefined ? '–' : PP[i] + '%') + '</td>';
    }
    return '<table class="card" width="100%" cellspacing="0" bgcolor="#6E9FB0" style="background:#6E9FB0"><tr><td colspan="' + (b - a + 1) + '" class="ct">' + ten + '</td></tr><tr>' + h1 + '</tr><tr>' + h2 + '</tr><tr>' + h3 + '</tr><tr>' + h4 + '</tr></table>';
  };
  var loai = w.quaKhu ? 'SỐ LIỆU NGÀY ĐÃ QUA' : 'DỰ BÁO';
  var mua = bMax ? '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td><div class="nho">💧 Mưa sắp tới</div><div class="vua">Khả năng mưa cao nhất: buổi ' + bMax.ten.toLowerCase() + ' (' + bMax.tu + ':00–' + bMax.den + ':00)</div></td><td align="right" class="lon">' + bMax.pop + '%</td></tr></table>' : '';
  var mai = d.mai ? '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td class="vua">Ngày mai ' + d.mai.date.slice(8, 10) + '/' + d.mai.date.slice(5, 7) + ' &nbsp;' + ic(d.mai.code, false, 26) + ' ' + e(TQN_WMO_VI[d.mai.code] || '') + (d.mai.rain ? ' · ' + tqnSoVi_(d.mai.rain) + ' mm' : '') + '</td><td align="right" class="vua"><b>' + so(d.mai.tmax) + '° ' + so(d.mai.tmin) + '°</b></td></tr></table>' : '';
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#0F172A;margin:0;padding:16px;background:#DCEBFA}' +
    '.top td{font-size:10px;font-weight:bold;color:#1E3A8A}.tag{background:#0F6E66;color:#fff;padding:2px 7px;border-radius:8px}.tag2{background:#B45309;color:#fff;padding:2px 7px;border-radius:8px}' +
    '.hero{color:#fff}.big{font-size:54px;font-weight:300;line-height:1}.cond{font-size:20px}.sub{font-size:12px;line-height:1.6}' +
    '.sum{color:#fff}' +
    '.card{background:#6E9FB0;color:#fff;border-radius:12px;margin-top:10px}.card td{text-align:center;padding:3px 0}.ct{text-align:left !important;padding:8px 12px 2px !important;font-weight:bold;font-size:12px}' +
    '.hh{font-size:11px;color:#E2E8F0}.tt{font-size:17px;font-weight:bold}.pp{font-size:10px;color:#E0F2FE}' +
    '.card2{background:#4F7FC0;color:#fff;border-radius:12px;margin-top:10px}.card2 td{padding:10px 14px}.nho{font-size:11px}.vua{font-size:13px}.lon{font-size:30px;font-weight:bold}' +
    '.ft{font-size:8.5px;color:#334155;margin-top:10px;line-height:1.5}.warn{background:#1E3A8A;color:#DBEAFE;padding:1px 5px;border-radius:3px}' +
    '</style></head><body bgcolor="#DCEBFA"><table width="100%" cellspacing="0" cellpadding="12" bgcolor="#DCEBFA" style="background:#DCEBFA"><tr><td>' +
    '<table class="top" width="100%"><tr><td>TQN · BẢN TIN THỜI TIẾT CÔNG TRƯỜNG</td><td align="right"><span class="tag">' + loai + '</span> <span class="tag2">TỰ ĐỘNG</span></td></tr></table>' +
    '<table width="100%" cellspacing="0" cellpadding="14" bgcolor="#5B9BE0" style="background:#5B9BE0;margin-top:8px"><tr><td class="hero" style="color:#fff"><div style="font-size:15px;font-weight:bold">' + e(cfg.project_name || cfg.project_id || '') + '</div>' +
    '<table width="100%"><tr><td><div class="big">' + so(d.tmax) + '°</div><div class="cond">' + e(TQN_WMO_VI[d.code] || '') + '</div>' +
    '<div class="sub">↑' + so(d.tmax) + '° / ↓' + so(d.tmin) + '°' + (d.feel !== null && d.feel !== undefined ? ' · Cảm giác như ' + so(d.feel) + '°' : '') + '<br>' + TQN_THU[day.getDay()] + ', ' + p[2] + '/' + p[1] + '/' + p[0] + (d.sunrise ? ' · Mặt trời mọc ' + d.sunrise + ', lặn ' + d.sunset : '') + '</div></td>' +
    '<td align="right" valign="top">' + ic(d.code, false, 90) + '</td></tr></table></td></tr></table>' +
    '<table width="100%" cellspacing="0" cellpadding="10" bgcolor="#4F86C6" style="background:#4F86C6;margin-top:10px"><tr><td class="sum" style="color:#fff;font-size:13px">' + e(tom) + '</td></tr></table>' +
    khoi('Sáng 06:00–11:00', 6, 11) + khoi('Chiều 12:00–17:00', 12, 17) + khoi('Tối 18:00–23:00', 18, 23) +
    mua + '<table class="card2" width="100%" cellspacing="0" bgcolor="#4F7FC0" style="background:#4F7FC0"><tr><td class="vua">Lượng mưa cả ngày <b>' + (d.rain === null ? '–' : tqnSoVi_(d.rain) + ' mm') + '</b> · Độ ẩm <b>' + (d.hum === null ? '–' : d.hum + '%') + '</b> · Gió <b>' + so(d.wind) + ' km/h</b>' + (d.gust ? ', giật ' + so(d.gust) + ' km/h' : '') + '</td></tr></table>' + mai +
    '<div class="ft">TQN Register · ' + e(cfg.project_id || '') + ' · GPS ' + cfg.lat + ', ' + cfg.lon + ' · Nguồn Open-Meteo (open-meteo.com, CC BY 4.0) · ' + (w.quaKhu ? 'SỐ LIỆU MÔ HÌNH' : 'DỰ BÁO') + ', không phải số đo tại công trường · tạo lúc ' + Utilities.formatDate(w.fetchedAt, cfg.timezone, 'HH:mm dd/MM/yyyy') +
    '<br><span class="warn">Bản nội bộ – chưa xác minh giấy phép thương mại</span></div></td></tr></table></body></html>';
}
function tqnEsc_(x) { return String(x === null || x === undefined ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
/** Lưu bản tin PDF vào Drive (thư mục riêng), trả link. Lỗi thì trả '' – không chặn việc ghi Register. */
function tqnWeatherBanTin_(cfg, w) {
  try {
    if (typeof DriveApp === 'undefined') return '';
    var ten = cfg.weather_folder || 'TQN Register – Bản tin thời tiết';
    var it = DriveApp.getFoldersByName(ten), f = it.hasNext() ? it.next() : DriveApp.createFolder(ten);
    var name = 'ThoiTiet_' + w.data.date + (w.quaKhu ? '_ca-ngay' : '_' + Utilities.formatDate(w.fetchedAt, cfg.timezone, 'HH') + 'h');
    var cu = f.getFilesByName(name + '.pdf'); while (cu.hasNext()) cu.next().setTrashed(true);   // cập nhật lại thì thay bản cũ
    var blob = Utilities.newBlob(tqnWeatherBanTinHtml_(cfg, w), 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
    return f.createFile(blob).getUrl();
  } catch (e) { tqnAudit_('THOI_TIET', 'Không tạo được bản tin: ' + e.message, {}); return ''; }
}

/** Cập nhật thời tiết cho khoảng ngày (vd từ 01/11 đến hôm nay). tu/den: 'yyyy-MM-dd' hoặc 'dd/MM/yyyy'. */
function tqnWeatherKhoang(tu, den) {
  var cfg = tqnWeatherCfg_(), why = tqnWeatherBlocked_(cfg);
  if (why) return 'CHAN: ' + why;
  var chuan = function (x) { x = String(x || '').trim(); var m = x.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); if (/^\d{4}-\d{2}-\d{2}$/.test(x)) return x; throw new Error('Ngày không đúng dạng dd/mm/yyyy: ' + x); };
  tu = chuan(tu); den = chuan(den);
  var homNay = Utilities.formatDate(new Date(), cfg.timezone, 'yyyy-MM-dd');
  if (den > homNay) den = homNay;
  if (tu > den) throw new Error('Ngày bắt đầu sau ngày kết thúc.');
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return 'BAN';
  try {
    var w0 = tqnWeatherFetch_(cfg, { tu: tu, den: den });
    if (!w0.ok) { tqnWeatherRaw_(cfg, w0, 'FAILED'); return 'FAILED: ' + w0.err; }
    var n = w0.j.daily.time.length, kq = { THEM: 0, CAP_NHAT: 0, BO_QUA_NHAP_TAY: 0 }, banTin = 0;
    for (var k = 0; k < n; k++) {
      var w = { ok: true, fetchedAt: w0.fetchedAt, data: tqnWeatherDay_(w0.j, k), quaKhu: w0.j.daily.time[k] < homNay };
      if (w.data.tmax === null) continue;
      w.banTin = tqnWeatherBanTin_(cfg, w); if (w.banTin) banTin++;
      var act = tqnWeatherWrite_(cfg, w); kq[act] = (kq[act] || 0) + 1;
    }
    try { tqnDinhDangThoiTiet(); } catch (e) { }
    tqnWeatherRaw_(cfg, w0, 'KHOANG ' + tu + '→' + den);
    tqnAudit_('THOI_TIET', 'Khoảng ' + tu + '→' + den, kq);
    return 'Đã thêm ' + kq.THEM + ' ngày, cập nhật ' + kq.CAP_NHAT + ' ngày' + (kq.BO_QUA_NHAP_TAY ? ', giữ nguyên ' + kq.BO_QUA_NHAP_TAY + ' ngày nhập tay' : '') + ' · ' + banTin + ' bản tin hình trong Drive';
  } finally { lock.releaseLock(); }
}
function tqnMenuWeatherKhoang() {
  var ui = SpreadsheetApp.getUi(), tz = tqnWeatherCfg_().timezone, now = new Date();
  var macDinh = '01/' + Utilities.formatDate(now, tz, 'MM/yyyy') + ' - ' + Utilities.formatDate(now, tz, 'dd/MM/yyyy');
  var r = ui.prompt('Cập nhật thời tiết theo khoảng ngày', 'Nhập: từ ngày - đến ngày (dd/mm/yyyy - dd/mm/yyyy)\nVí dụ: ' + macDinh + '\nNgày đã qua dùng số liệu mô hình Open-Meteo, ghi rõ trong Ghi chú. Dòng nhập tay không bị ghi đè.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = (r.getResponseText() || macDinh).split(/\s*[-–]\s*|\s+đến\s+/);
  try { ui.alert('Thời tiết', tqnWeatherKhoang(t[0], t[1] || t[0]), ui.ButtonSet.OK); } catch (e) { ui.alert('Lỗi: ' + e.message); }
}

/** Hình nhỏ đặt cạnh ô Ghi chú: biểu đồ Sáng/Trưa/Chiều/Tối (khả năng mưa % và nhiệt độ cao °C). Trả data URL PNG, lỗi trả ''. */
function tqnWeatherAnh_(w) {
  try {
    if (typeof Charts === 'undefined') return '';
    var b = w.data.buoi || []; if (!b.length) return '';
    var dt = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Buổi').addColumn(Charts.ColumnType.NUMBER, 'Mưa %').addColumn(Charts.ColumnType.NUMBER, 'Nhiệt °C');
    b.forEach(function (x) { dt.addRow([x.ten + ' ' + (TQN_WMO_VI[x.code] || '').split(' ')[0].toLowerCase(), x.pop === null ? 0 : x.pop, Math.round(x.tmax)]); });
    var p = w.data.date.split('-');
    var c = Charts.newColumnChart().setDataTable(dt.build()).setColors(['#3B82F6', '#F59E0B'])
      .setTitle(p[2] + '/' + p[1] + ' · ' + (TQN_WMO_VI[w.data.code] || '') + ' · ' + Math.round(w.data.tmin) + '–' + Math.round(w.data.tmax) + '°C')
      .setRange(0, 100).setLegendPosition(Charts.Position.BOTTOM).setDimensions(320, 160).build();
    return 'data:image/png;base64,' + Utilities.base64Encode(c.getAs('image/png').getBytes());
  } catch (e) { return ''; }
}

/* ===== Biểu tượng thời tiết từng buổi (đặt ngay cạnh Ghi chú: Sáng | Trưa | Chiều | Tối) ===== */
var TQN_WX_ICON = {"nang": "data:image/png;base64,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", "trang": "data:image/png;base64,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", "itmay": "data:image/png;base64,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", "itmay_dem": "data:image/png;base64,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", "may": "data:image/png;base64,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", "suong": "data:image/png;base64,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", "mua_phun": "data:image/png;base64,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", "mua": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAJqElEQVR42u2cXYxdVRXH/2vvfT9npkNBMJgmBarU3OFBQ4IESCZKTSE8+HSuDyBoP8YYJJb5sKIm557E8NGZUnmQMC2QYpQ09/BAjBEVCJnEDwSfKveSQKgQoEWMLdPOzP04e6/lw70DbQ39wLkfZzi/t8ncde85e+211n/vvc4BEhISEhISEhISEhISEhISEhISEhISEhISEhISEjoGxeMyhTwvVIVChaojI1KoVGT5P9WRESpUKlStjkgYegyQJG5dITzP0165rM/LplzWnufpJAL+D3zfV6VSSYhas3nL1ANDBvoaEb4WUBtI6fXsLCttlLB7C+A3iNSLFu6lx6d3ngAAEaFSqURBEHDigPOa9WUdhkUHANt/OHMdifo2FG5WSq/TxgAsYOH2pQsUKUARnLVgdu+A8YwQ79+3a/Ivp39f4oBzmPlBEPC28ZkrdUr9nJS6WSkDa5tw1jEAJgAgoZPKg7STvtJGK2PSYLYQ5mdcxDsefXDyteXvTRxwxiJbVGEYurGpmW3KmBml9HCjXmeCiACKiM56rSIiBLCAKJPNKmY3z9ZO7p2efNTzPB2G5b4r0qofBt/3SxSGods2NfNIKpPd56wdbtRrjggKRPpcBh8AiIhan4dq1GvOWTucymT3bZuaeSQMQ+f7JQKEkgg4TbWExaLbPjG9NzMwsL2+uGgFcs6DfvaIIJcdGDCNxcV9+3ZPjS3/XhIBrZxvwmLRbZ14oJQbHNpeW1qIQDArMfgfRQRMbWkhyg0Obd868UApLBad7/vmUx8By+pky/jMjZmMeTaKrANEd+iaBCCXShndaNivP/7g5PP9oo5Ur/J+oVCRsZ2zw0pjlgUkwqqDE4JEWLGAlMbs2M7Z4UKhIv1QD1RvUk9JB0HA7OZ35HL5DTZqWiLq6LUQkbJR0+Zy+Q3s5ncEQcC+X9KfwhQkJEIoju/ODiv3ijapy521AurCZBCwNoacjf45z/qqwprj0ZEjn6NLLz3serVO6LYD6K67HkofvPCo27gwuNlksr9tNOpM1M1IFJfO5HSj2bzlsemJ350kmcgrhqrbdYG6LTeX/94yvmt/Lp+/vV6rOSJ0TZWIiGhtBCTvCNPzWtNBIfXH2ft+UP1oNV6Sbi3YqAt3TO0kLJ7npy/acMFNIvJVdrwVhEFpbSR0NxJFoIyBMWkQEaJmw4LwAomant2149mTt0Vi7YBTNtamZraQonFj0iNEhGajDpHe7QqIgFvbFgARTCqdAQCwcweONj+4M9wTHPV9UUFAHEsHLM+gb91570W5weyvU6nMZmebsFHkBCRE0H20GSiAMATIDQzpRr328oJqbH7y/nuOdToSqDOD35o5t999/zX5TOZXSqe+0KgvWYBUdwvuJ4kMjrK5gVTUbLxcU42bf3nfPUepnULj4QDfVwgCvm1i+pK81v/QxlzSrNctKTKICSIS5fKDqdrS4nPrBhc2V6tVCsPQxWAhJuQD8Hw/nQWejuPgtxdtqVptIcrl85veXVxzWxiG7nyPRnsSActSc+vkrp/k80M/W1o8YYniNfgnRQGbVIpcZF+bZ/Xl8oPjdWqfBvVnBIhQWCzyd+556GIFGm/Ua0yARkxpbV1EnMlmN65V/DUiEs8LVd+mIL9U0gBER9EtmWz+QuesYIW2lXvoBhFAnMg3W7K6v2sA+76vIOI5dhL3oW8LHxJmAmEDRKhyUj9Sf9UAEQKR3HqXvyaXHjyktLqInYt/BIiI0prY8b8ju3DlEw8FHyzfa19FgFds5cZ8enCjNmaYmVdB+jkFE+Wy1PcylBStUVobEay29kCdy6QpBusAYoisps5Z1tpAINV1OP6B7/tqpVfEK6vRFTsR8KpxgEA5dlCk1h3B0IV7g8n/9GUNCMseA0Ckmq+6qLmoNamebnWu4GLAWeuyufw6tyC3AJC23O6zFEQkvu+rQ+nmMQd5VekUQKslEkgcOxFgEwCMjFSlX2uAmgsCq0BPGWMgsjr69ImERIQIsh4AKpVCfzogQGvGW8VPNxv1RrvLYTVV5Fp/q6AgYM8r6/3TO1+31h3IZLNKRGKfhkRItNZCwN86oRxX9MvCQkVEhJzCj6JG/ajShuJejIlALoqISZ4BgGp1ZEXvZ2V3K+fmpDoyop+8e+zEl67dlMrlB25stpquYrkrKiw2m8ubZqP+7GMzU/f7gHr44e8z+rkzLix67HmePj60OL20uPBiPj+QEuFI4hUJIiJRKpM2UbN5TFHme+jQzgp1rBWFCFvv3r1Wp/CbdDZ3fbNeh3OWAWIiECAKp96ViEBaOat3CoogRErpdCYHFzWP1G3tG/t3//jlTh3OU0f7gYjk6rGx1NVrN+4k0HdJ6XVaazjn4JwFO+blJScRKa0NtDEgpbreLiHtwWDHsC46QcBT8wvzPz3wi+BwJzsjqONNWac96ciQmxSpUWa50hgzTEQAEWzUdAAdZua/AzjYA70vShkQqYPQ7qXZeyfeBTrfoNWNiUZeuaxOfyplzJ/5jF2QglFKa2MAx+9lI35zz56JGtAfT2uWyx4TdTYdUrefdn+/UKG5IHAft0jzvLIuFCo9O0tYfhK/W93S1EOZQcXiR4fchUJFgiCQ1bV6TkhISEhISEhISEhI6Is1vmj4omJrv3oQGh19wXzy1wX02h4x3IrwRH9+6P0bOZ0/fOiRoVdiZx9fBwgBJOvvOHZB2tj3SJk0i7ykyfw1Irvn0N61by/34PSnPWL+sg6v9TspFd2gzKBhjlinBr4isDtcY7551hvvtX3cHTBaaEUaKdpEOq0JiCDihO3Tbz1x+Xutovjxg9Br+9g7YC6AK3ivpAW8WVwNAihSKS2EOQAYPct19No+5g5oza7F/GevIJW5gm2dCcq46EQziuj37QHi/rWPuQNG/Xb+TclNOjWUFrBVJquEG4cGlv51qF0guV/tY++AOYABISWySdiCBEI6B4L6QzW8qjnqn7k5rNf2MXeAEALi9Xe8OSxQ14tbghBpcQ0RlucAYK56piPIXtvH3QEfyr+BG5QZGGa2Vilj2C7OR7z4JwBAeIb822v7uDvgVPmXIRJxpPMg8J/feuKy+fOTj923j70D/kf+EYiUARM9B5Cct3zssn3MHdCaXbXsBZcpylz+SeVj7+xj7oBl+WfSaU/nhjMQjkhnWFzzjfORj72yj70D5rA8uyjFzfrbIEqnBgeNAj1/7vKxd/bxT0Htl929/tjFpePvVr4olL4uqi3ca4kPnDpAfWqPVXTwEm/71eIEr6xHfTEfvkc0VvYJCQkJCQkJq47/ApuYkLPQLnbPAAAAAElFTkSuQmCC", "dong": "data:image/png;base64,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"};
function tqnWxIconKey_(c, dem) {
  if (c >= 95) return 'dong'; if (c >= 61 && c <= 67 || c >= 80) return 'mua'; if (c >= 51) return 'mua_phun';
  if (c >= 45) return 'suong'; if (c === 3) return 'may'; if (c >= 1) return dem ? 'itmay_dem' : 'itmay'; return dem ? 'trang' : 'nang';
}
function tqnWeatherIconCells_(sh, row, c0, d) {
  var b = d.buoi || []; if (!b.length || typeof SpreadsheetApp.newCellImage !== 'function') return;
  b.forEach(function (x, i) {
    var c = c0 + i, tieuDe = x.ten + ' ' + (x.tu < 10 ? '0' : '') + x.tu + '–' + x.den + 'h';
    var h = sh.getRange(1, c), cu = String(h.getValue());
    if (!cu || cu === 'Hình thời tiết (tự động)' || /–\d+h$/.test(cu)) h.setValue(tieuDe).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    var mo = x.ten + ': ' + (TQN_WMO_VI[x.code] || '') + ', ' + Math.round(x.tmin) + '–' + Math.round(x.tmax) + '°C, mưa ' + (x.pop === null ? '–' : x.pop + '%');
    sh.getRange(row, c).setValue(SpreadsheetApp.newCellImage().setSourceUrl(TQN_WX_ICON[tqnWxIconKey_(x.code, x.ten === 'Tối')]).setAltTextTitle(mo).setAltTextDescription(mo).build());
    if (sh.getColumnWidth(c) !== 78) sh.setColumnWidth(c, 78);
  });
}

/** Giao diện tab Theo dõi thời tiết cho dễ nhìn (chỉ định dạng, không đổi dữ liệu). */
function tqnDinhDangThoiTiet() {
  var cfg = tqnWeatherCfg_(), t = tqnWeatherSheet_(cfg), sh = t.sh, m = t.m;
  var last = Math.max(2, sh.getLastRow()), nc = m.note + 4;
  var hdr = sh.getRange(1, 1, 1, nc);
  hdr.setBackground('#1B3F7A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 40);
  [['#F59E0B', '#1F2937'], ['#EA580C', '#FFFFFF'], ['#2563EB', '#FFFFFF'], ['#4338CA', '#FFFFFF']].forEach(function (c, i) {
    sh.getRange(1, m.note + 1 + i).setBackground(c[0]).setFontColor(c[1]);
  });
  sh.getRange(1, m.note).setBackground('#0F6E66');
  var body = sh.getRange(2, 1, last - 1, nc);
  body.setFontSize(10).setVerticalAlignment('middle').setFontColor('#1E293B')
    .setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  var nen = []; for (var r = 2; r <= last; r++) { var x = []; for (var c = 0; c < nc; c++) x.push(r % 2 ? '#F1F5F9' : '#FFFFFF'); nen.push(x); }
  body.setBackgrounds(nen);   // 1 lần gọi cho cả bảng
  sh.getRange(2, 1, last - 1, 2).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 3, last - 1, m.note - 3).setHorizontalAlignment('center');
  if (m.date) sh.getRange(2, m.date, last - 1, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, m.note, last - 1, 1).setHorizontalAlignment('left').setFontSize(9);
  sh.getRange(2, m.note + 1, last - 1, 4).setHorizontalAlignment('center').setBackground('#EEF4FB');
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  tqnAudit_('THOI_TIET', 'Định dạng giao diện tab thời tiết', { dong: last - 1 });
  return 'Đã định dạng ' + (last - 1) + ' dòng';
}

/** Thử 1 ngày để xem bản tin (không đụng các ngày khác). */
function tqnThuBanTin1Ngay() { Logger.log(tqnWeatherKhoang('18/11/2026', '18/11/2026')); }

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
function tqnWorkflowData() { return tqnCached_('tqn_wf', tqnWorkflowDataGoc_, 180); }
function tqnWorkflowDataGoc_() {
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
function tqnAiCall_(cfg, prompt, maxTokens) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'), last = '';
  for (var i = 0; i < 3; i++) {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: cfg.ai_model, max_tokens: maxTokens || 1200, messages: [{ role: 'user', content: prompt }] })   // prompt: chữ, hoặc mảng [ảnh/PDF, chữ] khi đọc phiếu
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

/**
 * TrangChu.gs – tab "TQN Trang chủ": thanh nút bấm được giống hệt thanh nút của bản xem thử.
 * Google Sheets không cho chèn nút vào thanh công cụ, nên dùng hình nút đặt trên trang + gắn hàm (assignScript).
 * Chỉ tạo/sửa tab "TQN Trang chủ", không đụng tab khác.
 */
var TQN_TRANG_CHU = 'TQN Trang chủ';
// [nhóm, hàm, nhãn, rộng px, cao px, PNG base64 (ảnh 2x)]
var TQN_NUT = [["Hằng ngày", "", "Hằng ngày", 65, 35, "iVBORw0KGgoAAAANSUhEUgAAAIIAAABGCAIAAACPGNsIAAAJw0lEQVR4AexY+1PU1xXfXR6CAvKQx8pLEERL1IgJIAgiUQGTDIpGGyNNpjOJdSadsdPOtH9COv2hM03aTmfaxtQ2g4/WEuOjFPGFiuEtC0hAEZUuyEvkDQr9LGf35svd/a7r8s3ma+YyH6/ne+7j3Hs+955z7+pmxJ8KPKDTiD8VeEDQoAISNBpBg6BBFR5QxSTEaRA0qMIDqpiEOA2CBlV4QBWTEKdB0DAvD3yvOovToAo6BQ2CBlV4QBWTEKdB0KAKD6hiEuI0CBpU4QFVTEKcBkGDKjygikmI0+B6GmxYVIaG8YnJv3x+6nd/Pj42PmHDiFA9ywMK0DA0PGrs7ntj68awkMCbTW0Dg0Pt94wjo+PPMv2i1mOxv/nDPz76+MjhotNKrUEBGuD3o8WlJRdvvLJ21ZJA/xNfXsDn5OSUUlNU2zj60KCdeZu0Wm1XT39VXbMi01OAhqjw0Ky0pJHRsWNfnD9+qkwzo8ndnBrg76vI/NQ5SFxMRF52KuZ26Xrd8MgYhHmCp6HrYR9CCjAxdzs/fToNJdBp7LE2mbo+cd+ubXvztwBv79z68kvx1m3u3jeiO8BVIXxBCfT1D3JV+MQ07nR0XqtsqDO0Ym7QPBNIVBgNQPTgGvcNDEIPYNNwVfgcfDx8q63jamVDY8sdRFpo7GBZlB6LLXg9i2vzsHcA4wMIElxVT98j6IGBR0NcFU/D+StVCCkA55SJyUkogS//e5Ubgj5LLn6FWuDitRrScOXR4vNHi0uB6vpb0irQAyVQUdMo1U89eXLyzKXf/qno2Bdllyvqzl2oOHz0DC4C1s6V9oLc3dOP0YC/HT/b2/8IGoavapuhB+ALpoQAehDo//jZyX+fvXylou5UydXff/pPRNfRMdkMJ7de0IDxAXTHyAzYx5//qwR6YHScH5angXV7LmF8fBJ7lrpgQz158pRkmyV8ivY2q6TK4nNXWm7fI81i30U6nWmq2FBHTpwFc6S3X87MzJwuvW6/DWonJqeOFZ9HoIcMsHDa1v7gcNEZmzEH85db70srYwMWmwIy9jFrg2Hrm1rpGhkbtTQ8LBgaKUxrk347Jxta7kxPz1BfcNDcepdkmyWWXXql0mYVU9Y3tcEL+AwM8HunYNvB9woOvb8nI2UtNDBUVl4NwREYu3tvNt223xKXi8GhEbRBxD/wox0HCnfAXHxMJDSPh0euVxsgcODWi50nbZCVnkSfFdXfnO/KWnMyz0xbR7XSUiEamucstWHup9QeyYZbdzq7bOQYqkVZfqMeJfD6lrTI8FAInp4e6clrloYtgYyDPzQ8CsERlJVXgXi5lgjTjS3tqPVa4Jmfm0EbGYfvzW3p7m5u0NNugCCFYXaBPou8F3h6QM+tN2F5VGhwAPT3OrvpkIEnShUrYiPDggNRxUEBGmCAjK1fk5CYEAMDMC/nJp1O6+HujjbnyipQ2kRv/yB19/VZyJ1fsILECNAgNrszpZeXJ2Rk7Esy6Qq1LG7ERodLxwTrbxdshSHc+tBMCrZeuBsHCFUdD7powpAJm9IsB6LKdJhu1DSRfnP6ehK4UpaGmoaWS9dqGa5VNnA92WfNzRaSE+KiV8ZHk4yoQgJXuunc0pJXQ4koXylz6caNBQ2AQH8/lCZY/gUFLI6J0gPkYot6zv+40dO3PmQJAjHkWkNr79xcDSVhcDYcQQ4K4G1hB8AQgFopHFkv7NLBRXprbu0wdvdihMSE2ACZe7wsDYgbCIsMVXOvNxiUoelrUyZY6O2FB0TcsggPD9Nmv9nYxhpwQmpSIp19RB6bVxEWQ7AlqS+cyDYECf0Dj6nKfpmzOQW53U6unpwyPzMpvGA0RCEywUpkI+gZHFxv9kbTxofpUyXl6IvNkZFqym2QrSFLQ3CQP7YDA56O1p2huX23k+7gK5abchqMxcdEQP94eOSB8SEEa6BN3msboIe7cT+GwEGn05IGayChp2+QbQgS+h85RMNiPx+8aTAI9qPNA6rTmj1gvmBoNLc7OskEK6enpzECwfH1RuhDcCbQi7qvTYzz9/PBp02YJ2Fdl5e9ofCtXIa33sy2bgMNy06+PotwGQf8fM3GGmxdUWbwyNZocG5WxpnCF9Ije5Rpzc7X4GBhZAD3QpTAQu8FtCG8vRbg0z4YeZpZW2mvrkYuRZcL5dVTlr2PTwJGJmFszPyjJE4q2WK7gRpQ+VzrzdjwMvXSarWZqWaZNFwpSwPXzuYnfjhqbb9PVXj14GECVFhueLge4M1CtdbllsxXPWZzNQt3M5YNGWq5S/T0DZBPoyPCaENAsB7KvgYXHkqzyNUI01zj4CWmKw2UbDckr/sB2fL28oJeiuddb5AltyHisb0lHZDJ86IBe9mOoxFzwASzxAnYoemzuZocLa3FpPUhQdBghLrGVgjzRFxMRGz0UgxibWtZpN7NzeSE9vvGZwY6Q0u70+uFdTswzcBOtf2qhlvm58Ke/Nd+9dNChh15mdSRHWH65MoUS67m9PhMTkpECZRersRlAe7DaxZ7GT9UQOkEcrJMudq6IyhfsyoOepgoOllq7O6DDCu1hq/p0YtPBsP81svGsRacpwHX5/91me5hWElMpF46dHxMJG0x/OqAX+6kVVIZEZNytVRJ8qr46KTVKyBj9+FXrF9/8vdP/nqi+NxlGIUSJwk3CAiOw5SrLdRyvbIz1tOTCteKz46d+ejjI58Wnf7PhRuUWpFpaS0wPc/1cnaln87TwK6k8bGRcKh0UMwbN1fS1DfZiyosV1NjabktKwWPtbCQIJYq3d3dIpaG4Ef1g+8WwK3Sxo7IeK+AP+uWSFH7d+duTFkrfabgXbJ8WXh+bub+3Tm0OkXWa22dNDwN7+zKodhCrw9qhBIZhvT41QWfwKa0daR5Y2s6Pjns3L6JatNeMT3WUPvLD/dD84uD+yBLgQgGPQCnS/WQV69a/t7e7T//yb5392z/oDAffffvysEFFDSjVg5I4xgNwBtY2ga5+sMf74YewA9wc6rc3TYmr4GJnx34YeHuXDQ79P5eXA5xKFmz510vOuLdA1vAoQ/24tMOeBrsNP2uquB0vFqkW/XbmwkCbLg+2OahcdKoY91eABocW8iL3UrQoAr+BA2CBlV4QBWTEKdB0KAKD6hiEuI0CBpU4QFVTEKcBkGDKjygikmo8TSowjGunYSgwbX+lrEmaJBxjGvVggbX+lvGmqBBxjGuVQsaXOtvGWuCBhnHuFYtaHCtv2WsCRpkHONataDBtf6WsSZokHGMU2qnOwkanHadkh0FDUp60+mxBA1Ou07JjoIGJb3p9FiCBqddp2RHQYOS3nR6LEGD065TsuP/AQAA//+kJ9ZBAAAABklEQVQDAAL6A2G1oFfuAAAAAElFTkSuQmCC"], ["Hằng ngày", "tqnMenuManager", "Quản lý hồ sơ", 128, 36, "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"], ["Hằng ngày", "tqnMenuRun", "Nhập & kiểm tra", 142, 36, "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"], ["Hằng ngày", "tqnMenuWorkflow", "Kiểm soát quy trình + AI", 186, 36, "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"], ["Hằng ngày", "tqnMenuNewDoc", "Tạo hồ sơ / cấp mã", 158, 36, "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"], ["Hằng ngày", "tqnShowSidebar", "Bảng điều khiển", 142, 36, "iVBORw0KGgoAAAANSUhEUgAAAR4AAABICAIAAACvCw3fAAAQAElEQVR4AeydB3xUxRbGs7vpFRJC8+FTBEUFFCkCgoA0AUE6CIiAYOgt9C6hSQlVJEIEQhOk+UBREETECAktBFSk+R4dEtJ7e//NJDebTbJpm00gk9/J5Nxzzzkz99z57sycubtRp+T2Ex0bHxwadT8o/O7D0Fv3QyTJCJS2CNy+H0LnBwIAATjkhpi082qzHH6SU1JCI2LuPAwV7hISk5KSU3LQlWIZgac5AvR7Oj8QAFfAAVCERcSkIDV40dlDKzI67t6jcErsrSzNnextyjs7VCzn+EyFMv+SJCNQyiJAt6fzAwGAYGVhDigitAAJAyAGwKUPLcwehUQyXjGq2dlYVnZ1ci1r72BnZWmhMdeoVQY8yVMyAk9pBOj2dH4gABBcne0BBdAQ07qgkEggk+11Z4JWYlLy/eDwuPhEc3M1MC3raKuWaMo2bFJYhBEo6a4BBdAAIMAkNj7xQXA4i6Wsjc6AFihkvELJxsqigosjMM2qLSUyAjICIgIApIKzo7WVBQPSw5BI4CPkSpkBLYY2cIWqSxk7RkBFQzIyAjIC2UZApTIrV8aOoQjgAB89nTRohUfFxickWZhrwJWehjyUEZARMBAB5zJ2zAyBT0RUnK6aFlrJySkRkbFIwZUcr4iDJBmBvEcAyJRzskc/PCpGd1qohVZYZAw5ensbK6aPaEh6UiMg211MEWDUsrOxJFUYnjpEiVaoOY6KiWfW6OhgLUSylBGQEchvBNjyYvjS3elSs8bCi421pRp4wUmSEZARyH8EyMgDIuxi4xIoIbXg2GPmQJKMgIxAgSNgZWmObYwCLTIbHJMbpJQkIyAjUOAIsOLCNiExiRJSJycn86eI3mEiOxKbYBYalRIUkfIwLOVBIQhznIRGp+AQt7TZiCRdyQgUPgIatTYpmJT+FjvQ0nZUjUYrLbx3xQP+I2K0cAqLTolLNEtKNtNWo5zOP4M5ThhscQjMcE4V+XcjLWQEiioCGrUWRGKsog41XZY/xqW4hJSg8JToeON6zeQN51QRm1AUzc9UkTyQEchjBEQekJS70NfiTHDGKiNjU0KjCztG5aUxoCos2izzDnhe7KSOjIApImBkaMXEp5i4r4NkKjVFqGQdMgL5iUAhoZWpqvjElPCYTBLTHFApU1DT1CVrkRHIYwSMCa3oTG8n5rEBxlGLKcp1nXGaKL2UsggYDVrkHMkEFlf0qJoGFFftsl4ZgawRMBq04hNVWb2bUlLsDTDlxcq6Sn4EjAatYs8lmCYRf+HiZd9T/rfv3C35t7a0tlB73Q8ePho5fupGnx1/X72uPS6OX6NBK/31juK4iNQ6ExJT/xRxMXbSzG59Pq7f9N2+A4cHPw4p4tqM7P7w0eOVX3itUtXaJ31PG9l1CXOnVqtDQsKmzVn4TvvuxfUcNBq02GXKNbzJySm37j0+E3jzTOANA+R/8fq3P5094feXjs5NDJMNvn+Rlwbk2sJcFZYumN2/Tw/Ujv1yctOWr2FKDvGc7vnhJ9CJk6d0W5WSkvL3tRuUbVo2n+I+ilNTZ81PKOJH0ZlzAbQEGug2hhpzouDgx+gIgs9JTZH/899bcfG556xcy7ls27j2rUYNkpKSLv1xRTE3JWM0aOXaaADz3pClnYd6us3wdpvxlUHaOHf1vrHzturoeGOI+ZnAm7lWlKvCaf9z4l7qlkNGuNPhVq3dQJ8w4KHq8/+uXesVoaDRaARTQspzFwJ//e0U9CgoWGkSfLM2naHmbbswzI4ePrhLp/bXbvyzYdM2dP746+83GrdmHGvaqtN//3cbibEoNDSMlkB+Z84b8BkXn4COIHgDmpwaNHRsoxYd6jRs+dvvfhzqUlx8/PQ5C7mh/QYN/+vK1aioaC4q8NKfPbp0bNa0ka6myXgTQev+o7Cx87c+CArL+4WJ10Z09TEfM88HV7rCAvAh6Xdd3FFRHjx0ZNPWnQuXrurY/cOVn6/Pye2IsVMmTP3UuWyZ6ZPGug3un5OayeSJiYn0IUE3//mfqJfhC8me/d9xCHiuXtc+jxi4Lqc+vz0/+3TPdu9GDeoy8HbuNYBeOHPKuNt37r3XvR8PHUxKJsXHJxw6fIy2cft++90fRqG79x507T1o87ZdTRo14KnRqedHv/qenjx+5Kxp41ctm29jXTyf8TURtM5cuhEbq/2IWJ+Ojdd5DPSaN8gAdXynjoiaooMJhghxgiuYIqVFy1Z/5bMj2yr69Oo6cdyIGZPHjRw6qLjuWbYNy0n4Zv03enV7nwnSBz27NHqzHmrWVlaNG9YHcv0Hj7Kztd2/c9PwTwbu3PJlQnxi7/5u3/94FJ0SSJaWFjwCKlWs0Lxp49493ldaePb8xXff733t+s2dPl6Myd/t3fZitaqMb0nJyX17dVPUTM+YCFoPHqWNV8P7ta5f+4V6taoaoMoVyopAKDqYYCiEiitxWMjyrwsn7924CJ3//addW76s+UoN4dBz1Toek4LXLTu2bzN+lBvdVFdYjLy5uTmNF/T8c8+KlqzxXIikW+cOHAKkFUs8Lvr97LnoUwsL7Wf1EEI8I25fPc9Vv1yjOocN6tUhFDf/9G/ftiWHJZN4BJzzPbJj87pnq/xLaWHdOrW5uisBv73VqAFCF+eyB/dsvXs9wH30UA6LkUwErcQk7afCuE4bawvKApBiqLgqgBMDJhUrlG/6VsPZ09yFDsuS6zf+EbxSRkREHv/Vd+36TT8c+fmf/95S5IJhVeN7yh+68vc1JMz+T/mdXbfBZ9eeby//mf1KWjj08t7CpDQyMgor/7MX8AAlp36ODklOlJSUFBD4B6Prdz/8ZCABEBsXhzdB8HremA3+fvoM666933537cY/pDp0FUhhC8PzAYG6ciZgQs5WhK48LzxtFraUWVMpXNSFgEvrN27dvnMvodBrDwMUVtDDR0F6deHq4qU/tu3cs2XHblql91jkXmAFiYUfbhm0mf/7bP/md7+zVKrnzSiHJoKWUdpqAiflXcsptZjrPOPPXbjoPnXOy280/eCjoR4LPcl6sZ4mC3/v/gNFn7wcEmje4hX01LpvtenSe+CnC5aOmTizVYcedBdFUzDzF694pe7bOJwzfwnr71r1m9MtevYbggcoPkE7fxaaWUtA26h5ByZCrN0HDx9fs37zqbMXZFVD8uhREN4EwSMRRF+cPMOjRp0mXT8YNHPuZyPGTSWTQbZDNy9/7PhJYTh+yhxhJcrvf/xJyKfOmickeSwPfH+YNgvbn0/8pjuK4sHvzLkmrTq169Jnlsdiot2pR//RE6YjV2iU+zRhe/zEb4oQhv2r1xu+07ZTb5bBk6bPbdf5gzqNWuoGPDg4RBj27u8Gtlu069ayffeps+YTga69B7qNmogTo5OEVqaQ+p+7II6rPFO5avr8ijGqQ9d+PEf1Hm88BTv3GsiqWpgopZ//eXpqcPBjRQKzYPHKm+lpBg537Nq3Zt1XJCHgBTGk0C0oxaGB8nFI6PCxU25l3rZmJ0DXvwFzcWrKrHk8s3UbgPzq9Zt9Bg7TG6OQG4UYjsZMmCFcMVmdNlE/KT/KfbreXGD3voPfHvxBmORUfrPvwLQ5C4mJrgKH4HNDahZUVx4TG8vDS0wrFDnDPhUph8ZiSju0/M+kTcB4Wi9btW7uAk8iywJm4dzpSm59+WovhFCzJo2OHNx1LfDU1q/WOjo4IPnfrdv7//M9jC6FR0TUea3W+s+XsQBgd8XFxZmzYGbrjt0wEAoMWTAQC6RNX65kCUFXK1dOq4nQMJHAVEbLSeNHsNJYvWzBqy+/ZNhK9yzpUJ4USFiJsSY5c/LwN1vXN2xQFwmjGVkcGOPSnbv3B7iNoWfjlon38s88VCr9N+MqV6q4dOFslnyH9u9g7Ycm5L15O2VOdPvO3WnpwzV5du7Oz4f2KCvhzzzXMOPVs7W1teHmXj57QreWjVuMv0VZ2qH14eCRYqrQo9+QpSvW0ulJqdHPWrZoKm4J8/KyZZzoDdC0SWPIc9jZ2bZs3qRLp3ZCIfCPvwSjlC7OZclWvdeuNfB7p1mTPj27iFPM7wXzzd4DrOXgK1eqcGD3lratWpD4GjXsY9IPCA0TEN2+a6/QWbJg1riRbmT/und578BuH7YEhDzXkoeI0JkwdtiEscOfqVyxSeM3vb9Y7uCg/RZYZrYKdIVaAUqVKg05KpUKRH00ZBQrN/yQNfFe66k3FUQOrf98KTk9J0fH12u/OnTwR0ig8wGXuAUw2dIX6zeLNSr3a9Wy+dydGi9VJ2EjHhOcOnz0uJ7hFys/G9CvF7GiljEjhoizrNMM1CJ08luWdmhljRdbxguWrLyauhfEWXrG1z5eJA+h2jXTNouRgx9KKCjzxA+JI5BK7aPwUMUK5Smh8IhISui0/zlKqFnTxoofDp0cHSkNE3tT9Bih075tK8FQ2tjYODnlbo7mrdt32FSFgXQ9pPa2mgghgkBZGFJ6KsyYCTNEIqfKv57Z6r1WADir83IuLorw2SrPCJ4pa9aRR5yiZMFGCXXQCQWHPCkooawXwnMEuSDl1hiuRSjntyzt0GL6QZ4auvGH30avla+/VpMFFbmpvgOGhYaFK9EEaSQMOvca0LB5++dq1KtUtbZn+iyRrqOoZcso352anJz2NVri+Y3yi9VeoMwXKckxMi6AIV+2QvnW7Yx3iyfPnEcGRSEFcszfhHLhS5Y9ZC+En769ujJQC95wqVFn9Myk9LjpmXCnlLXZ5m07lauA+fbAIaF89959wWRbKreGsznVwqmCUcYFFMz+qbFi//fd1i327vhK3HuSBLv3HRBXt//AoXfadSPNzWhz+869F6u/wOTwuX9XEWcLUManvwXH3DK/5oWxFXVFRmmz/IIn6a9L7B8IeXRMjGCMW7JqVdBbeM/R0THKc428n+6F8CgU/qOii+RChHPDZfFCy3DbiuEsAHslPR/w15VrtIBlvcciTyYM8CxpAk4fO3xgJ5PDrp3aIykY2drYCMOIyLQpojjMS1kYW+HfITUBI/iJ40ZkSw3rvyEUjFJaWVqysYsr9pdGjp8WH29oXwG1PJK9vR0JJ6Hcu0fnbC+ke+q+udAxcSmhpR/w4OC0j4rY29ty7uz5gLv3tJtXrLwXzZ2huzTibMGoSvpCogCfJlIWIUFBj5lrFaABL1XPmIWSaxk/yi0rNW5YH88a87T3jxMNbrKhaYBItHqtWUKOpGwZJ9RYdC1Z8TlM4Yll8Ksvvyj8kJLNehVIenbLeCVKaJqslNDKFOpjv5xUdnVeTX3pieyW0LBM/RF8cnLy2fMXBV+AUrySg+HBQ0du6mx2BQReRmiYqlerWqG8q9Dx1tm3Ya/m/v2HQm64ZIVG3lLozF2wNCYm05Rp9RfeCmLLpW4boMk0WFftzytXEeaRnBwdSIHS5tnTJggTNvSYWgu+kKWShlm3YfO1zG/PsGtP3q+Q/gtjXtqhNWjoOFa9glq/17PvwOEimvQ/n3PK+gAACIRJREFUkXSq9crLzGcQRkVFfzLSnQn9Lyd/Hzx8PCVCKCW3N5LQ0aP2bVrS1RDis2P3D0lIsmvJfvHUWfMRGia1Wt2re9qTmFTKkBHu4HOx5+dcgvIUMOyBs5PdR1JCR4+fxPDI0V+gTxcsfbv1+zSme9+PWcZwllkcM2QYMv5uoyexObt5264Bn4wRe2LI80U0u8XbbwmTsZNmiirEYYHLwQP6srmPOU+o7n0+3rX3PydOnmLfr0vvgeza9+g7JPDyn5wtFirt0PI95Q9aBF1K36EiCf7lmqXsLXJL2MYdnb778cORn+mIvfu7HTp8zCE9vX7vfp7GClwpZGdn67NhNehFwgYXAwVY3bJjt7IoR26A3EcPa/XO20IBXA0Z4b58jReZw7xPVtlFmDBmmPBAerr/kFHQug0+YvXfpmVzce3EYcTQQUIN7I12nz5l5rwff/rZJX00E6fyXi5bNIeNCfTJ7IFkmEIS7VzoMUM8+8i7kuXv1f8TtrxP+Z3F82u1X32pejWY/JDRdEs7tPQCydbtkIH9ThzZr8zZUGDKPnvaBGsrK3hBbzdpuHpZ2gjD7EgsxsSpPJZ0brKRNV7SvnWOCQs5nugrl2a8kqdRp61zOKtHlpYW3muX9+2V8YkJ9mG9Vi9p1yYfL627jxnms351syaN7O3tFP9kR1cs9hCfRBbCcSM/cR89VJP+oU9AxUDhuSjTK4VCMy8l4Z06cbTQ9Nn+zfFffQVfmJLt+0P7d3Tr3IHdCMUPgy3bwZvXryJWitDEjImgZa5Jqygm9VNbBbhIxVBxVQAnwuTd1i3YyMqWzvkemTtzku5NEiZDB/c/c/LHPdu9Ib8Th3b6fMn6QXi4ez2AHokad1dIfI8d5FChAR/2FnJ6gCKEeenFake/+2b/zk0Hdm+5EuC7fdMXytqGrUzAhk5ORI9ZunD2rz/9h/YE+h0/9v2eTh3aLlkwS1RESxRDdmmFkBJekcO0btnsax+vvwN8aTB+aN7pXw4xbeOUQmq1esLY4QGnjqJw4si3l/yPe8yazLCGNwgTRVOPYVxFAbp89oTuqQH9eiEU1LxpY0IneEp4RZOHDhJBDJ5CTjuFRC85wZNljedCkrcXTh3d9/VGQhrof5wHBAAThngWhpS6Qci2FmFS+DKtxxfekWEPFVy12SF01m494n/x+hmD341x90Fajk5RwwRDzCHFFbwpiQd244b1Id17U+AGkNNf++VGclxv1n+j3huv0QlYz7BIEA5r1XxZMIbLalWfoz1MWQ2rGT5LG55/7ln8vF77VSWXrWcirr36C8/ryUvaISvYhg3qElKm3MXeNhNBq17NqlZW2k9qbT/gO3TmRp0vvfgqK3/g2HkRF+UUJhgixAmuYLJS2itrWU+USMnkmR4ei5a/1bLjLI/FrF5IigwfM9kv/WskehVfyrhERuuJbJTRoKU22LUrujqtnPlheZc8veQmAqn8MxVxSIk5TnAFn5XURruUrL6NL7GytMQpea31G7eScxOpESQQi4QO72a8HIjkySDZyswRMFp/tMhx1Z1WYf1aVb/bMHH/uvFe8z72MvjdGF7zBs4a1WXFjH46ah9jiDlO0txl+WNutEvJ4roIBAvnTt+w1vP119Jeh6UGwNa+bcu9X29kkcChpCc9Akbrj5bmBoet1Dip1aoqlZzr1Xpe+dKLbJn6tV94v1XdtxvU0Dn7PIaYp7rJvrCyyL0B2VsWk5Sh6dC+7SQhyBD8cnj/1cBT3l8sb5T6oaliapGs1pgRMBq0rLQrKWO2LL++LDO+UiW/psWpTxKCFMKL1apa6HxfQHE2SNZtpAgYDVoatVkxdm4rczMaYKSYSDcyAkaIgNGgRVsciuerFKnZzCHtVXIt/8T9ygY/lREwJrTMNSrH4ujiVKoxnKB8Km+dvKiSHQFjQosrtbFU2WW8D4SgyInqqLTIq5EVyAjkMwJGhha121urnGzNVHBFTFRRxtaM6oq4HuleRqAgEVCr6KFmZln3ZwviLN3G2kLl4lC0WQ1SJlRh9aQl3NMjJP8+hREQIBKA4vLU/PAnOf8fOsLKAGnUqrJ2Kmd7FaMKMNCoCzuO8QTACa6YAeIW5xo1MgNNkKdkBEwagaRUEKnValGrWqPWdtCk5BRxbNzSQmMGEoBBOQdVeSdVhUIQ5jjBFVjFrXHbKb3JCBQ+Akmp0BKAwpvawlz7hlJ8gkn+XSkVSpIReEojEJ+g/TI8y/Stf7W1lQVXGhNnnC/ZwZUkGYHSGYGYWO1/ghWAIgJqa0sttOLiExMT0/5PD1JJMgIyAvmKQGJScuqoZWZNPiDVUk1Cw9bGEj4sMoZSkoyAjEABIhAWoYWPXSqUhLk2m+Fkr32HgjmhHLhEUGQpI5CvCCQkJgEfTASUYCAttMhp2Ntq36EIDosqkkQh9UiSEXhKIwBkgkO13/XtYGelVmvz7eJCtdCCA22kCgGfUEIiSUZARiAvEQAyLLSAj2Pq7E8xSYMWKy7XsvYMX7FxCUGhUWJfWVGSjIxAkUfgCayA8QqwABmNWu3qbJ8xYKVeSxq04NVqlauzg0ajRvXh43CAiFCSjICMQLYRACAPgsMBi7lGiys1o1NmvQxoIUepgouDpYUmITH5flB4SHh0ctG8pUFdkmQEntAIAAqgAUASE5OtLMzLuzgAnKzXkglanAZ85Z0dyjjawETFxN99FPYoJDIiOo6cPTBFQZKMQGmLAOsjOj8QiIiKAw6AAmgAkLKOtq7O9jDZBkQfWkLJ3saqkqsjaUNGOXaTydk/fBwBTG8/CJUkI1DaInDnYSidHwiw9wscAAXQqOTqpLuLJYCjW2YPLTRUKlUZB5tnypcpV8bOxtqCBAhJDuSSSloEZHuKOgLkJ+j8QMDW2gI4AAqgoUJqsOL/AwAA//8h17o4AAAABklEQVQDAPQdSVTmJbscAAAAAElFTkSuQmCC"], ["Hằng ngày", "tqnMenuRemindPreview", "Nhắc hạn hôm nay", 156, 36, "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"], ["Tự động", "", "Tự động", 51, 35, "iVBORw0KGgoAAAANSUhEUgAAAGYAAABGCAIAAACxEXYDAAAIYElEQVR4AexY+U9UVxSeGdYBBBSRRUFxQYGCiopWREndsYpi3ahWq63V1qaNXf8Em/7QJhoTm9iauFQbl6B13zdEAQUBZZN9dWAAlU1Z+j3PcPN4wwxeGFyaO/lycu5559x73/fOvffc0bSLHycDGpX4cTIgKOMkTKUSlAnKuBngDhBZJijjZoA7QGSZoIybAe4AkWWCMm4GuANElgnKuBngDnhtWcY9s7c2QFDG/WkEZYIybga4A0SWCcq4GeAOEFkmKONmgDtAZNn/nbKE5Ixt2/f+unN/flEZ97taKOBdyrLLN5OvxN/1cO9vY2P9z/FLGVn5FiKBr5t3hrIH2QUlZbqQwJGfLIvasGqhn4/X3bSs8soqvte1hHcnyqr1dflF5eZBgz6rbyS36po6spB8+qyB7PraJ7CcvXwb6+hRQSl0BQpLKsgTsqi0svn5C4WDohnoP2xVzOyQgBHJ9zNzC0pmTB2/eulcL4+BCreCYsP8Ffb6hiYMBOAdFY/QxOh5haXxiWkp6TkVj6thMYNOlCXczTgUd8E8qC8MQG537j0kC8mc/GKyJ6dmwULEuTg7QVfgyMkr5Al54Oi533Yd3Lnn6JlLCQ2NTQpPal6Jv/f7rkP7jpy9dCMZX+Kvgye37z5svDYPxV1Eh0ByaiYFkgSVMAJ4R7KQfNHScuzUVYyOlX4tIeXM5YQ9h07tPnCivNIkcZ0oo14sJZFHgKODduAAl1fp88nT+pSMnP1HzjU1PVf437xzPyE5vaW1VW4Huf+ev4nEkRuZjvc37oc9ZUrcmetZj4qo6dLPUaORCNFV1+49fBosk10hJQ9mWjBr6s9fryHERM0gO1YEWUiSsVuJ5fP3sfNWVpqZERPMOKvVanS7ddOq+TPf9xrkBk+s9ONnr0NhqNDpb9y5j6bW3m7x/OnffLZ809olMyMmwtLe3g7WFFTCDmCtXbieCMUMUh/k5uaXwGFAf+ePY+ZsXhfz7efLIyaPhaWtrR3pDMUYnSgzftxjC94fY29eGxPo79dtJ7Y21mMDR65ZNt/dzRXOeUVlmbmFUAhJKQ9BDfTI8NAxI4dqtXauzk6TxgUEB4xQqVT1DY1pDx5BMUZ6Zl5phc7Yziw3bqeSjlzxGewB3dbWJjwsxNtT2iIfV9Vga4ZRgb6izNvTHWM7OWoV45lpajTqieMCyAFvSwokJQLyMdB/GJoMQaOHk44NlBQm0ZWNtTWa2Bwhu0SVvo4Y6efkMNjTXe4DBldEzwKoE/kj6H1FGbruAXy9pU+NwCp9LSSAk66pWdranPs5Kl7A3c2wRVbpO53aiLLSWE0NC4aCXSkxpdMBBSOh7skzUga4OpPCpFt/Fz9fL8De3pYZmfIGKEM9tW373ucvqwqsOOhAfFIa5sSysrGxGU2gqdmgYPGiKYe9nR01m5oMPtQkOSU0qL9LP+hYfTgooCiAzY4sWIykVOlrr8bfk0NfI5VK9JTJN0CZk6O0ELDQaBJYFICzkyOabe3tkACdXJKiNsyw4wlsBoBu0lhX1CQJI44U6KDm4vUkKApoNGqysH501XW3ktPloCKJ3Jg0TIi1X4PiP9xnzbJ5uPRgLLwYdOC9MdLGRJsL7A5ae0jAwcGgsHSDkVDfUcFp7Q3pRnbIdpVEve9gDxwXaKJ8YwWq2kCUig3BahEHrR0+HmDcITph6CFl1tZW1EVbWxspJFtbDU3mQPZXlMWlleTp5eFGip2tDa1W3DcAMpJkLAx0k85ZMirkrOmTaAdM6qhsWbZ6uA8gZ111DSXa0CGe+HgAFHrUpewhZa7O0jaBHtm8oQPlHZc+V5cuKn44mEFLa2vSfUPJTgUEOfv5epOCWo8UkvfSskkZPtTgQE25BN3hL88BIkX+CB8DlRAsWLkooaG8InpImbfnQNTKGENXXYudovVlcqVm5ODyDCO2CVoR0M0Db4LyHUAgril0AQQF8u8cNj7Q2kpK6ltJ6eevJmJLLil7fPTkFarOsQkGj5EKNFMDTe44B4wdwkKDyHjhWiLKGkwGifwwp7BSpyd7l7KHlKGveR9MUauljQGnOP7Awql3+lIC7EDk1FDz2wF8GHDvA1Dx19Q+hREl5aK5EVAYUN8uj55JJyYS7Y99cfuOnM3OK4YDzsTYpXNoW0SzS6jVajoHjJ8GjBoaGuwPOz45bhG/7Ni348/DcWeu1dRJM0GGYmg8VaDnlGG9xC6ZjauGvEdHB/uFc6YhL+TGbnVkJSYXNNoP9yFcXOztlNUQNvL1sQuH+XixkxQMBvr7rVu5wLWrO79iRISbyvo5kZNRuHoOcsMcKAq78BDvQfjquLp0+YdCB2XkLpP+I3xx+wMU31zmokJGbFwdvXZFFAplAG+7Zf1HeHO5jyl96xcr0Tnhx69Wb4hdCK5NvRg6ATUrF8/64cvYT1cu2LgmGtfSRXOnYT/CIzl+2rIafX6/OVZuhI6PATsAgtCUA/vmuhVR322KXbs8Cj0jFn8rTZkQhAuy3I3pJiljHt0q2ETxj8WJczcgsQq69e+Nw4PsfNz28WcR/mLrTT/GsSAIx7TxTcDY0wKUoY7HPzOosPF/Q1ubVBAZD2MpS0m5DvcnbNLlpv/PstRYpvqxAGW4cETPixjl57N0QSTbEUyN10v7tLCQ8LCQiCnjaNvuZW89C7cAZRg4YNSwpR9GjhruA71P4eigxX9K4ZOCzZ+SfToHy1DWp1N82zoXlHF/EUGZoIybAe4AkWWCMm4GuANElgnKuBngDnjtWcY9w7cuQFDG/UkEZYIybga4A0SWCcq4GeAOEFkmKONmgDtAZJmgjJsB7gCRZYIybga4A0SWCcoYA32miCzjplZQJijjZoA74D8AAAD//+/NJicAAAAGSURBVAMAEWKCl3JKuPQAAAAASUVORK5CYII="], ["Tự động", "tqnMenuSetupAll", "Bật tự động tất cả", 152, 36, "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"], ["Tự động", "tqnMenuAutoStatus", "Trạng thái", 109, 36, "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"], ["Thời tiết", "", "Thời tiết", 55, 35, "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"], ["Thời tiết", "tqnMenuWeatherNow", "Cập nhật ngay", 132, 36, "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"], ["Thời tiết", "tqnMenuWeatherKhoang", "Theo khoảng ngày", 155, 36, "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"], ["AI", "", "AI", 15, 35, "iVBORw0KGgoAAAANSUhEUgAAAB4AAABGCAIAAABUj7NvAAACAUlEQVR4AeyWS0sCYRSGa7Lotii6UtGF7osI2rTpB/QDatGiTbQJqk2bflAEQRFBLYqgnxA5ecnMsTRvGFJpmab22gF1TqPOUG7ik5fD937fnHfGxzOglKnYR6qq2EdEM7QCiADCCDArJkQAYQSYFRMigDACzP6/CQmFI8qDH/pMpdiXJev1hyB/8Iksq0WB4E/V7uHZ3tE5ZL1xszayO/un0MHJBVlWi0Y7FW88nqCrZfsdLQzVotGyLR/neQy+RGOGcnGxdjSe1+n24rijrQUVMlvzd4LVI+1oy40rnc6gf252urmpAQuzxYlqSNrRst2FlJoaaWSwb2J0EGsA8fpCWOiXRnTk+TUQys4TcpE+OTpAcVdWYw+uEX0pOyhrfCQb2tvdQUwcdw+pVJqO9FSNaKtDoc66WhPeF6izvRU7H4mk3XmPhU7xaNe9Lxp7p+b94wu8LxA2aUe2GWDCo80lgbo9gdhbnG5TtqqiE4nkreJBjyRVb60tbW8s5zQzNYZ96Mpyi6pHqmiLQ6Efaqi/p9ZkKuynEcSOXPCWwpaQKvralh1nXD3xPRtY5NTf29XYUA+L0fQFwliUlSp6eXGeCExNDv/s3FxdpNOe7nY6Jbu+skCWVVU0O/ulFdEMoMT8H1rBmsEUQAQQRoBZMSECCCPArJgQAYQRYFZMCAPyBQAA//8f+g/kAAAABklEQVQDAFUCDkJPZE2iAAAAAElFTkSuQmCC"], ["AI", "tqnMenuAiBatch", "AI rà hồ sơ", 112, 36, "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"], ["Giao diện", "", "Giao diện", 56, 35, "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"], ["Giao diện", "tqnMenuTheme", "Dễ nhìn", 94, 36, "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"], ["Giao diện", "tqnMenuRestoreTheme", "Trả lại gốc", 110, 36, "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"]];

function tqnTaoTrangChu() {
  var ss = tqnSs_(), sh = ss.getSheetByName(TQN_TRANG_CHU);
  if (!sh) sh = ss.insertSheet(TQN_TRANG_CHU, 0);
  sh.getImages().forEach(function (im) { try { im.remove(); } catch (e) { } });
  var COLW = 25, NCOL = 72;
  if (sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
  sh.setColumnWidths(1, NCOL, COLW);
  sh.setHiddenGridlines(true);
  try { sh.setTabColor('#1B3F7A'); } catch (e) { }
  // Tiêu đề
  sh.getRange(1, 1, 1, NCOL).merge().setValue('TQN Register – ' + (tqnConfig_().project_name || 'Register') + ' · bấm nút để chạy')
    .setBackground('#1B3F7A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13).setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(1, 40);
  // Hai hàng nút: hàng 1 = Hằng ngày, hàng 2 = các nhóm còn lại
  var hang = [TQN_NUT.filter(function (x) { return x[0] === 'Hằng ngày'; }), TQN_NUT.filter(function (x) { return x[0] !== 'Hằng ngày'; })];
  [3, 5].forEach(function (r) { sh.setRowHeight(r, 44); });
  sh.setRowHeight(2, 8); sh.setRowHeight(4, 6);
  var n = 0;
  hang.forEach(function (ds, k) {
    var row = k ? 5 : 3, x = 10, nhomTruoc = '';
    ds.forEach(function (b) {
      if (b[1] === '' && nhomTruoc) x += 14;   // khoảng cách giữa các nhóm
      var w = Math.round(b[3]), h = Math.round(b[4]);
      var blob = Utilities.newBlob(Utilities.base64Decode(b[5]), 'image/png', 'nut.png');
      var col = Math.floor(x / COLW) + 1, ox = x % COLW, oy = Math.max(0, Math.round((44 - h) / 2));
      var im = sh.insertImage(blob, col, row, ox, oy);
      im.setWidth(w).setHeight(h);
      if (b[1]) { try { im.assignScript(b[1]); } catch (e) { } try { im.setAltTextTitle(b[2]); } catch (e) { } n++; }
      x += w + 8; nhomTruoc = b[0];
    });
  });
  // Tóm tắt nhanh
  var kq = '';
  try { var d = tqnManagerData(true), k2 = d.kpi; kq = 'Quá hạn ' + k2.quaHan + ' · Sắp hạn ' + k2.sapHan + ' · Đang chờ ' + k2.cho + ' · Bị trả về ' + k2.traVe + ' · Tổng ' + k2.tong; } catch (e) { kq = ''; }
  sh.getRange(7, 1, 1, NCOL).merge().setValue(kq ? 'Tình hình hồ sơ (cập nhật ' + Utilities.formatDate(new Date(), tqnConfig_().timezone, 'HH:mm dd/MM') + '): ' + kq : '')
    .setFontSize(11).setFontColor('#1E293B').setFontWeight('bold');
  sh.getRange(8, 1, 1, NCOL).merge().setValue('Mẹo: các nút này cũng có trong menu TQN Register. Nút bấm được trên máy tính (ứng dụng điện thoại không chạy nút hình).')
    .setFontSize(9).setFontColor('#64748B');
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
  tqnAudit_('TRANG_CHU', 'Tạo tab Trang chủ + ' + n + ' nút', {});
  return n;
}
function tqnMenuTrangChu() { var n = tqnTaoTrangChu(); SpreadsheetApp.getActive().toast('Đã tạo tab "' + TQN_TRANG_CHU + '" với ' + n + ' nút.', 'TQN Register', 5); }

/**
 * Gắn hàng nút "Hằng ngày" (y như bản xem thử) lên dòng 2 của từng sổ hồ sơ – nằm trong vùng cố định,
 * làm việc ở sổ nào cũng thấy. Chỉ chèn hình (không ghi ô), bấm lại thì làm mới, không nhân đôi.
 */
function tqnGanHangNutVaoSo() {
  var cfg = tqnConfig_(), ss = tqnSs_(), ds = TQN_NUT.filter(function (x) { return x[0] === 'Hằng ngày'; }), out = [];
  cfg.detailSheets.forEach(function (name) {
    var sh = ss.getSheetByName(name); if (!sh) return;
    sh.getImages().forEach(function (im) { try { if (String(im.getAltTextTitle() || '').indexOf('TQN_NUT:') === 0) im.remove(); } catch (e) { } });
    try { sh.setRowHeight(2, 42); } catch (e) { }
    var x = 6, c = 1, x0 = 0, n = 0, maxC = sh.getMaxColumns();
    ds.forEach(function (b) {
      var w = Math.round(b[3]), h = Math.round(b[4]);
      while (c < maxC && x0 + sh.getColumnWidth(c) <= x) { x0 += sh.getColumnWidth(c); c++; }
      var im = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(b[5]), 'image/png', 'nut.png'), c, 2, x - x0, Math.max(0, Math.round((42 - h) / 2)));
      im.setWidth(w).setHeight(h);
      try { im.setAltTextTitle('TQN_NUT:' + b[2]); } catch (e) { }
      if (b[1]) { try { im.assignScript(b[1]); } catch (e) { } n++; }
      x += w + 6;
    });
    out.push(name + ': ' + n + ' nút');
  });
  tqnAudit_('HANG_NUT_SO', out.join(' | '), {});
  return out;
}
function tqnMenuHangNut() { var ui = SpreadsheetApp.getUi(); ui.alert('TQN – hàng nút trên từng sổ', tqnGanHangNutVaoSo().join('\n') + '\n\nNút nằm ở dòng 2 mỗi sổ. Bấm được trên máy tính (ứng dụng điện thoại không chạy nút hình).', ui.ButtonSet.OK); }

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
      im.setWidth(441).setHeight(42); im.setAltTextTitle('TQN_BADGE_MA'); im.assignScript('tqnMenuTrangThaiMa');
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
    sh.getRange(hr, start).setValue(TQN_QL.NHAN + ' (tự động – xem tab ' + TQN_QL.SH_MAU + ')');
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
  return k;
}

/** Ghi lại dropdown + màu cho các cột cấp duyệt (gọi khi cài / đổi mẫu phiếu). */
function tqnQlDinhDang_(sh, cfg, p, k) {
  var n = Math.max(sh.getMaxRows() - cfg.firstRow + 1, 1), rules = sh.getConditionalFormatRules().filter(function (r) {
    var rg = r.getRanges()[0]; return !(rg && rg.getColumn() >= k.start);
  });
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
    [['✔', '#B7E4C7', '#0B5F57'], ['↩', '#F8C9C4', '#8A1C12'], ['⏳', '#FFE8A3', '#6B4E00'], ['✎', '#ECEFF3', '#475467']].forEach(function (x) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(x[0]).setBackground(x[1]).setFontColor(x[2]).setRanges([g]).build());
    });
  }
  if (k.kt > 0) {
    var t = sh.getRange(cfg.firstRow, k.kt, n, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setBackground('#FDE2E1').setFontColor('#8A1C12').setRanges([t]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✔').setFontColor('#0B5F57').setRanges([t]).build());
  }
  sh.setConditionalFormatRules(rules);
  // 2 cột tự tính: cảnh báo khi gõ tay (không khóa cứng)
  [k.giu, k.kt].forEach(function (c) {
    if (c < 0) return;
    var pr = sh.getRange(cfg.firstRow, c, n, 1).protect().setDescription('TQN – cột tự tính, không cần gõ');
    pr.setWarningOnly(true);
  });
}

/** Đọc lại các dòng r1..r2 (số dòng sheet) và ghi cột Ai đang giữ + Kiểm tra. Trả số dòng có cảnh báo. */
function tqnQlTinhDong_(sh, cfg, p, k, r1, r2) {
  r1 = Math.max(r1 || cfg.firstRow, cfg.firstRow); r2 = Math.min(r2 || sh.getLastRow(), sh.getLastRow());
  if (r2 < r1) return 0;
  var n = r2 - r1 + 1, lastC = Math.max(sh.getLastColumn(), k.kt), v = sh.getRange(r1, 1, n, lastC).getValues();
  var lk = k.link > 0 ? sh.getRange(r1, k.link, n, 1).getRichTextValues() : null, m = k.map, out = [], loi = 0;
  var ngay = [m.nt && m.nt.plan, m.nt && m.nt.actual].filter(function (c) { return c != null && c >= 0; });
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
    var ds = tqnQlCotTuVan_(tqnDetailMap_(hdr));
    ds.forEach(function (k) { if (an) sh.hideColumns(k.from, k.n); else sh.showColumns(k.from, k.n); });
    if (ds.length) out.push(name + ': ' + ds.map(function (k) { return k.n + ' cột'; }).join(' + '));
  });
  PropertiesService.getDocumentProperties().setProperty('tqnAnTuVan', an ? 'on' : 'off');
  tqnAudit_(an ? 'AN_TU_VAN' : 'HIEN_TU_VAN', out.join(' | '), {});
  return out;
}
function tqnMenuAnTuVan() { var ui = SpreadsheetApp.getUi(); ui.alert('Đã ẩn khối Tư vấn (dữ liệu vẫn giữ, chỉ ẩn cột)', tqnAnHienTuVan(true).join('\n') || 'Không thấy khối tư vấn', ui.ButtonSet.OK); }
function tqnMenuHienTuVan() { var ui = SpreadsheetApp.getUi(); ui.alert('Đã hiện lại khối Tư vấn', tqnAnHienTuVan(false).join('\n'), ui.ButtonSet.OK); }

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

/**
 * Auto.gs – bật/tắt toàn bộ tự động bằng 1 nút. Chạy trên máy chủ Google: máy anh tắt vẫn chạy.
 *  - tqnHourlyJob : mỗi giờ nhập & kiểm lại (chỉ đọc tab gốc, ghi tab TQN_*), không gửi mail.
 *  - tqnDailyJob  : nhắc hạn ~07:00 (Remind.gs).
 *  - tqnWeatherJob: thời tiết ~06:00 (Weather.gs), chỉ khi đủ điều kiện GPS + điều khoản nguồn.
 */
var TQN_AUTO_HANDLERS = ['tqnHourlyJob', 'tqnDailyJob', 'tqnWeatherJob', 'tqnMoBangKhiMo', 'tqnTuDongMaKhiSua', 'tqnPhanPhoiFile', 'tqnQuanLyKhiSua', 'tqnEmailSang', 'tqnEmailSep'];
/** Mở bảng Quản lý hồ sơ mỗi khi mở file (trình kích hoạt onOpen cài đặt – trình kích hoạt đơn giản không mở được hộp thoại). */
function tqnMoBangKhiMo() { try { tqnMenuManager(); } catch (e) { } }
function tqnBatMoBangKhiMo() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnMoBangKhiMo') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tqnMoBangKhiMo').forSpreadsheet(tqnSs_()).onOpen().create();
  return 'ok';
}

function tqnHourlyJob() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { tqnAudit_('TU_DONG_GIO', 'Bỏ qua: đang có lần chạy khác', {}); return 'BAN'; }
  try {
    tqnCacheClear_();
    tqnEnsureConfig_();
    var cfg = tqnConfig_(), imp = tqnImport_(cfg), qc = tqnQC_(cfg, imp), rem = tqnRemindList_(cfg, imp, new Date());
    tqnWriteRemind_(cfg, rem);
    var wf = tqnBuildWorkflow_(cfg, imp, new Date());
    try { if (ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'tqnQuanLyKhiSua'; })) tqnQlTinhLaiTatCa(); } catch (e) { }
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
  else out.push('✔ Thời tiết: hằng ngày lúc ' + tqnInstallWeather());
  try { tqnBatMoBangKhiMo(); out.push('✔ Mở file là hiện bảng Quản lý hồ sơ'); } catch (e) { }
  try { out.push('✔ Tự cấp mã khi gõ Nội dung: ' + tqnBatTuDongMa().join(', ')); } catch (e) { out.push('✖ Tự cấp mã: ' + e.message); }
  try { tqnBatThaFile(); out.push('✔ Chỗ thả file: tự phân phối 15 phút/lần (Drive › TQN Register – Hồ sơ › 0. THẢ FILE VÀO ĐÂY)'); } catch (e) { out.push('✖ Thả file: ' + e.message); }
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
    (on.tqnWeatherJob ? '✔' : '✖') + ' Thời tiết (6h · 11h · 13h · 18h)' + (p.getProperty('lastWeather') ? ' – ' + p.getProperty('lastWeather') : ''),
    (on.tqnTuDongMaKhiSua ? '✔' : '✖') + ' Tự cấp mã SD · RFI khi gõ Nội dung',
    (on.tqnPhanPhoiFile ? '✔' : '✖') + ' Thả file tự phân phối 15 phút/lần' + (p.getProperty('lastThaFile') ? ' – ' + p.getProperty('lastThaFile') : '')
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
    .addItem('▣ Quản lý hồ sơ (bảng chính)', 'tqnMenuManager')
    .addItem('⌂ Tạo / làm mới tab Trang chủ (thanh nút)', 'tqnMenuTrangChu')
    .addItem('⌂ Gắn hàng nút lên đầu từng sổ (SD, RFA, RFI…)', 'tqnMenuHangNut')
    .addSeparator()
    .addItem('★ Bật tự động tất cả (bấm 1 lần)', 'tqnMenuSetupAll')
    .addItem('Trạng thái tự động', 'tqnMenuAutoStatus')
    .addItem('Tắt tự động tất cả', 'tqnMenuRemoveAll')
    .addSeparator()
    .addItem('Bảng điều khiển', 'tqnShowSidebar')
    .addItem('Quy trình hồ sơ (ai đang giữ việc)', 'tqnMenuWorkflow')
    .addItem('Tạo hồ sơ mới / Trình lại – cấp mã', 'tqnMenuNewDoc')
    .addItem('Tự cấp mã SD · RFI khi gõ Nội dung: BẬT', 'tqnMenuBatTuDongMa')
    .addItem('Cấp mã cho các dòng còn trống mã', 'tqnMenuCapMaTrong')
    .addItem('Kiểm & sửa mã nhập tay (trùng số, STT, khoảng trắng)', 'tqnMenuKiemMa')
    .addItem('📂 Thả file tự phân phối (mở thư mục / phân phối ngay)', 'tqnMenuThaFile')
    .addItem('⛔ Tiến độ ↔ hồ sơ chặn thi công (tạo / làm mới tab)', 'tqnMenuTienDo')
    .addItem('🛡 Quản lý theo mẫu phiếu A/B/C/D: cài / cập nhật', 'tqnMenuQuanLy')
    .addItem('🛡 Tính lại Ai đang giữ + Kiểm tra', 'tqnMenuTinhLaiQuanLy')
    .addItem('🙈 Ẩn khối Tư vấn (dự án không có tư vấn)', 'tqnMenuAnTuVan')
    .addItem('👁 Hiện lại khối Tư vấn', 'tqnMenuHienTuVan')
    .addItem('↻ Trình lại rev mới (đứng ở dòng bị trả về)', 'tqnMenuTrinhLai')
    .addItem('🔎 Tạo bộ lọc sẵn: Việc của từng người / Đang chờ / Cần sửa', 'tqnMenuBoLoc')
    .addItem('👥 Phân quyền theo email (cột kết quả phiếu)', 'tqnMenuPhanQuyen')
    .addItem('✉ Email nhắc (bị trả về · 7h sáng · thứ Hai cho sếp): bật / tắt', 'tqnMenuEmail')
    .addItem('🤖 AI đọc mẫu phiếu dự án mới → đề xuất', 'tqnMenuAiMauPhieu')
    .addItem('🌐 Trang web cập nhật hồ sơ (lấy link gửi đội)', 'tqnMenuMoWeb')
    .addItem('Tự cấp mã: tắt', 'tqnMenuTatTuDongMa')
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
    .addItem('Thời tiết: cập nhật từ ngày… đến ngày…', 'tqnMenuWeatherKhoang')
    .addItem('Thời tiết: bật tự động (6h · 11h · 13h · 18h)', 'tqnMenuWeatherOn')
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
    tqnCacheClear_();
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
/** Nhớ kết quả 3 phút cho các bảng mở nhiều lần (không phải đọc lại 32 sổ mỗi lần bấm). Chạy nhập/kiểm mới thì xóa nhớ. */
var TQN_CACHE_KEYS = ['tqn_sb', 'tqn_wf'];
function tqnCached_(key, fn, sec) {
  var c = null; try { c = CacheService.getDocumentCache(); } catch (e) { }
  if (c) { var v = c.get(key); if (v) { try { return JSON.parse(v); } catch (e) { } } }
  var kq = fn();
  if (c) { try { var s = JSON.stringify(kq); if (s.length < 95000) c.put(key, s, sec || 180); } catch (e) { } }
  return kq;
}
function tqnCacheClear_() { try { CacheService.getDocumentCache().removeAll(TQN_CACHE_KEYS); } catch (e) { } }
function tqnXoaNho() { tqnCacheClear_(); return true; }
function tqnSidebarData() { return tqnCached_('tqn_sb', tqnSidebarDataGoc_, 180); }
function tqnSidebarDataGoc_() {
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
/** Phiếu gửi văn bản cho Thông báo / Công văn / Lấy ý kiến / Biên bản (không phải hồ sơ trình duyệt). */
function tqnPhieuVanBanHtml_(x, c, ngay, han) {
  return '<div class="paper">' + tqnPaperHead_(c) + tqnPaperInfo_(c, 'PHIẾU GỬI VĂN BẢN – ' + tqnEsc_(String(x.loai).toUpperCase()), '#DDEBF7', x.so, ngay, han, 'Đề nghị phản hồi trước')
    + '<div class="vv">Về việc: ' + tqnEsc_(x.nd) + '</div><div class="bh">NỘI DUNG GỬI</div>'
    + '<table><tr><td style="height:80px;vertical-align:top">' + c.phieu_ben_gui + ' gửi ' + tqnEsc_(x.loai).toLowerCase() + ' số ' + tqnEsc_(x.so) + ': ' + tqnEsc_(x.nd) + '.<br>Kính đề nghị Quý đơn vị xem xét và phản hồi.</td></tr>'
    + '<tr><td style="height:34px;vertical-align:top">Đính kèm (nếu có):</td></tr></table>'
    + '<div class="bh">XÁC NHẬN TIẾP NHẬN / PHẢN HỒI</div><table class="sm"><tr><td class="gray" style="width:50%">BÊN NHẬN</td><td class="gray">Ý KIẾN PHẢN HỒI</td></tr>'
    + '<tr><td style="height:80px;text-align:center;vertical-align:top">Ký, ghi rõ họ tên<br><br><br>Ngày: …/…/20…</td><td style="vertical-align:top"></td></tr></table></div>';
}
function tqnPhieuHtml_(x, c, ngay, han) {
  if (/thông báo|công văn|lấy ý kiến|biên bản|^cv/i.test(String(x.loai))) return tqnPhieuVanBanHtml_(x, c, ngay, han);
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

