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
