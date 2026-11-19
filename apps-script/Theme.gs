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
/** Màu nền → màu nhạt dễ nhìn (giữ tông màu để vẫn phân biệt được khối). */
function tqnSoftBg_(hex) {
  var c = tqnHexToHsl_(hex); if (!c) return hex;
  if (c.l >= 0.97) return hex;                                   // trắng
  if (c.s < 0.12) return c.l < 0.5 ? '#EEF1F5' : '#F7F8FA';      // xám / đen
  if (c.l < 0.35) return '#EEF1F5';                              // xanh than đậm (cột ngăn) → vạch xám rất nhạt
  if (c.h >= 35 && c.h <= 65) return '#FFF5D6';                  // vàng
  return tqnHslToHex_(c.h, 0.62, 0.93);                          // màu khác → pastel cùng tông
}
function tqnSoftFont_(hex, bg) {
  var c = tqnHexToHsl_(hex), b = tqnHexToHsl_(bg);
  if (!c) return hex;
  if (c.l > 0.85 && b && b.l > 0.8) return '#1B3F7A';           // chữ trắng trên nền đã thành nhạt
  if (c.s > 0.5 && c.h >= 200 && c.h <= 250 && c.l >= 0.35) return '#1F3A68';   // xanh dương gắt
  if (c.s > 0.5 && (c.h <= 10 || c.h >= 350)) return '#C5221F';  // đỏ gắt
  return hex;
}

function tqnThemeSheets_(cfg) { return [cfg.log_sheet].concat(cfg.detailSheets); }

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
    if (hdr[i]) { b1 = '#EAF0FB'; f1 = '#1B3F7A'; w1 = 'bold'; }
    else { b1 = tqnSoftBg_(b0); f1 = tqnSoftFont_(f0, b1); }
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
