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
