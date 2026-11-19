/**
 * ThaFile.gs – Chỗ thả file: thả PDF/bản vẽ vào thư mục Drive "0. THẢ FILE VÀO ĐÂY" → TQN tự:
 *   1. đọc mã hồ sơ trong tên file (vd 261117-COA.NP-NT-SDS-005 … .pdf, 261116-COA-TBC-RFI-003 rev01.pdf)
 *   2. tìm đúng dòng trong sổ (SD, RFI, RFA, …) – so theo loại + gói + số thứ tự, không phụ thuộc ngày/rev
 *   3. chuyển file vào thư mục con theo sổ (SD, RFI, …), gắn link vào ô Mã hồ sơ (chữ giữ nguyên, bấm là mở file)
 *   4. ghi nhật ký TQN_THA_FILE; file không nhận ra mã → chuyển vào "Chưa nhận ra mã" + báo.
 * Không xóa file nào. Chạy 15 phút/lần (Bật tự động) hoặc bấm "Phân phối file ngay".
 */
var TQN_THA_HEADERS = ['Thời điểm', 'Tên file', 'Mã nhận ra', 'Sổ', 'Dòng', 'Kết quả', 'Link file'];
var TQN_THA_GOC = 'TQN Register – Hồ sơ', TQN_THA_VAO = '0. THẢ FILE VÀO ĐÂY', TQN_THA_LOI = 'Chưa nhận ra mã';

function tqnThuMucCon_(cha, ten) { var it = cha.getFoldersByName(ten); return it.hasNext() ? it.next() : cha.createFolder(ten); }
function tqnThaFolders_() {
  var p = PropertiesService.getDocumentProperties(), id = p.getProperty('thaFileRoot'), goc = null;
  if (id) try { goc = DriveApp.getFolderById(id); if (goc.isTrashed()) goc = null; } catch (e) { goc = null; }
  if (!goc) { var it = DriveApp.getFoldersByName(TQN_THA_GOC); goc = it.hasNext() ? it.next() : DriveApp.createFolder(TQN_THA_GOC); p.setProperty('thaFileRoot', goc.getId()); }
  return { goc: goc, vao: tqnThuMucCon_(goc, TQN_THA_VAO), loi: tqnThuMucCon_(goc, TQN_THA_LOI) };
}

/** Lấy mã trong tên file. Trả {code, key, type} hoặc null. */
function tqnMaTuTenFile_(ten) {
  var s = String(ten || '').replace(/\.[a-z0-9]{2,5}$/i, '').replace(/_/g, '-');
  var m = /(\d{6})\s*-\s*(COA[.\-A-Z0-9 ]*?)-\s*(RFI|RFA|RFC|SDS|MSS|MAS|OTS)\s*-\s*(?:([A-Z]{2,4})\s*-\s*)?(\d{3,4})(?:\s*-?\s*rev\s*\d+)?/i.exec(s);
  if (!m) return null;
  var code = tqnCleanCode_(m[0]);
  return { code: code, key: tqnMaKhoa_(code), type: m[3].toUpperCase() };
}

/** Bảng tra: khóa mã → {tab, row, col} ở mọi sổ chi tiết. */
function tqnBangMa_(cfg) {
  var map = {};
  cfg.detailSheets.forEach(function (n) {
    var sh = tqnSs_().getSheetByName(n); if (!sh) return;
    var lastC = Math.min(sh.getLastColumn(), 40), lastR = sh.getLastRow(); if (lastR < cfg.firstRow) return;
    var m = tqnDetailMap_(sh.getRange(cfg.hdrFrom, 1, cfg.hdrTo - cfg.hdrFrom + 1, lastC).getValues());
    sh.getRange(cfg.firstRow, m.code + 1, lastR - cfg.firstRow + 1, 1).getValues().forEach(function (r, i) {
      var c = tqnCleanCode_(r[0]); if (!/-(RFI|RFA|RFC|SDS|MSS|MAS|OTS)-/i.test(c)) return;
      map[tqnMaKhoa_(c)] = { tab: n, row: cfg.firstRow + i, col: m.code + 1, code: c };   // trùng → lấy dòng dưới cùng (mới nhất)
    });
  });
  return map;
}

/** Phân phối mọi file đang nằm trong thư mục thả. Trả danh sách kết quả. */
function tqnPhanPhoiFile() {
  var lock = LockService.getDocumentLock(); if (!lock.tryLock(20000)) return [];
  try {
    var cfg = tqnCodeCfg_(), f = tqnThaFolders_(), files = f.vao.getFiles(), bang = null, out = [], log = [];
    while (files.hasNext()) {
      var file = files.next(), ten = file.getName(), ma = tqnMaTuTenFile_(ten), kq = '', o = null;
      if (!bang) bang = tqnBangMa_(cfg);
      if (ma) o = bang[ma.key];
      if (o) {
        var dich = tqnThuMucCon_(f.goc, o.tab);
        file.moveTo(dich);
        var sh = tqnSs_().getSheetByName(o.tab), cell = sh.getRange(o.row, o.col), txt = String(cell.getDisplayValue() || o.code);
        cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(txt).setLinkUrl(file.getUrl()).build());
        var cu = cell.getNote(); cell.setNote((cu ? cu + '\n' : '') + 'File: ' + ten + ' (TQN gắn lúc ' + tqnGioNgay_(cfg) + ')');
        kq = 'ĐÃ PHÂN PHỐI → ' + o.tab + ' dòng ' + o.row;
        log.push([new Date(), ten, ma.code, o.tab, o.row, 'Đã gắn link + chuyển vào thư mục ' + o.tab, file.getUrl()]);
      } else {
        file.moveTo(f.loi);
        kq = ma ? 'Mã ' + ma.code + ' chưa có trong sổ' : 'Tên file không có mã hồ sơ';
        log.push([new Date(), ten, ma ? ma.code : '', '', '', kq + ' → để ở "' + TQN_THA_LOI + '"', file.getUrl()]);
      }
      out.push(ten + ': ' + kq);
    }
    if (log.length) {
      var lg = tqnOwnSheet_(TQN.PREFIX + 'THA_FILE', TQN_THA_HEADERS);
      lg.getRange(lg.getLastRow() + 1, 1, log.length, log[0].length).setValues(log);
      tqnAudit_('PHAN_PHOI_FILE', out.join(' | ').slice(0, 4000), { so: log.length });
      try { tqnSs_().toast(out.slice(0, 6).join('\n') + (out.length > 6 ? '\n…' : ''), 'TQN – đã phân phối ' + out.length + ' file', 8); } catch (e) { }
      try { tqnCacheClear_(); } catch (e) { }
    }
    PropertiesService.getDocumentProperties().setProperty('lastThaFile', Utilities.formatDate(new Date(), cfg.timezone, 'dd/MM/yyyy HH:mm') + ' · ' + out.length + ' file');
    return out;
  } finally { lock.releaseLock(); }
}

/** Bật: tạo thư mục + lịch 15 phút (không trùng). */
function tqnBatThaFile() {
  var f = tqnThaFolders_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnPhanPhoiFile') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tqnPhanPhoiFile').timeBased().everyMinutes(15).create();
  tqnAudit_('BAT_THA_FILE', f.vao.getUrl(), {});
  return f.vao.getUrl();
}
function tqnTatThaFile() { var n = 0; ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tqnPhanPhoiFile') { ScriptApp.deleteTrigger(t); n++; } }); return n; }

/* ---------- menu ---------- */
function tqnMenuThaFile() {
  var url = tqnBatThaFile(), html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:13px;line-height:1.5">' +
    '<div style="background:linear-gradient(90deg,#0F6E66,#1B3F7A);color:#fff;padding:10px 12px;border-radius:8px;font-weight:bold">⚡ CHỖ THẢ FILE – tự phân phối (15 phút/lần)</div>' +
    '<p>1. Bấm mở thư mục, kéo thả PDF / bản vẽ vào.<br>2. Tên file có mã hồ sơ, vd <b>261117-COA.NP-NT-SDS-005.pdf</b>, <b>261116-COA-TBC-RFI-003 rev01.pdf</b>.<br>' +
    '3. TQN tự chuyển file vào thư mục đúng sổ và gắn link vào ô <b>Mã hồ sơ</b> (bấm mã là mở file).<br>4. File không nhận ra mã → thư mục "' + TQN_THA_LOI + '".</p>' +
    '<p><a href="' + url + '" target="_blank" style="display:inline-block;background:#0F6E66;color:#fff;padding:8px 14px;border-radius:8px;text-decoration:none;font-weight:bold">📂 Mở thư mục thả file</a> ' +
    '<button onclick="this.disabled=true;this.innerText=\'Đang phân phối…\';google.script.run.withSuccessHandler(function(r){document.getElementById(\'kq\').innerText=r.length?r.join(\'\\n\'):\'Không có file mới trong thư mục.\';}).tqnPhanPhoiFile()" style="background:#1B3F7A;color:#fff;border:0;padding:8px 14px;border-radius:8px;font-weight:bold;cursor:pointer">Phân phối ngay</button></p>' +
    '<pre id="kq" style="white-space:pre-wrap;background:#F4F6F9;padding:8px;border-radius:6px;min-height:20px"></pre></div>').setWidth(560).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'TQN – Thả file tự phân phối');
}
