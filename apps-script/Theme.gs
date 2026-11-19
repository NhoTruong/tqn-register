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
  try { tqnApplyStatusColors(); } catch (e) { }
  try { tqnApplyOptionColors(); } catch (e) { }
  try { tqnApplyTabColors(); } catch (e) { }
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
    if (!(sh.getName() in goc)) goc[sh.getName()] = sh.getTabColor() || '';
    sh.setTabColor(c); n++;
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
