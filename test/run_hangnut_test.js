// Test hàng nút trên đầu từng sổ (TrangChu.gs › tqnGanHangNutVaoSo)
const fs = require('fs'), path = require('path'), vm = require('vm');
let p = 0, f = 0; const T = (t, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + t + (i ? ' | ' + i : '')); ok ? p++ : f++; };
function sh(name, widths) {
  const s = { name, imgs: [], h: {}, getImages: () => s.imgs.slice(), setRowHeight(r, v) { s.h[r] = v; }, getMaxColumns: () => widths.length, getColumnWidth: c => widths[c - 1],
    insertImage(b, c, r, ox, oy) { const im = { c, r, ox, oy, alt: '', setWidth(w) { im.w = w; return im; }, setHeight(h) { im.hh = h; return im; }, setAltTextTitle(a) { im.alt = a; }, getAltTextTitle: () => im.alt, assignScript(fn) { im.fn = fn; }, remove() { s.imgs = s.imgs.filter(x => x !== im); } }; s.imgs.push(im); return im; } };
  return s;
}
const SD = sh('SD', [40, 230, 330, 100, 110, 40, 20, 90, 90, 80, 80]), RFA = sh('RFA', [40, 240, 330, 100, 110, 40]);
const ctx = { console, Math, JSON, String, Utilities: { newBlob: () => ({}), base64Decode: () => [] } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../apps-script/TrangChu.gs'), 'utf8'), ctx);
Object.assign(ctx, { tqnConfig_: () => ({ detailSheets: ['SD', 'RFA', 'MAT'] }), tqnSs_: () => ({ getSheetByName: n => ({ SD, RFA })[n] || null }), tqnAudit_() {} });
const out = vm.runInContext('tqnGanHangNutVaoSo()', ctx);
T('N01 Gắn hàng nút lên dòng 2 của SD và RFA (MAT không có tab → bỏ qua)', out.length === 2 && SD.imgs.every(i => i.r === 2) && SD.h[2] === 42, out.join(' | '));
T('N02 Đủ 6 nút Hằng ngày có gắn lệnh (Quản lý hồ sơ … Nhắc hạn)', SD.imgs.filter(i => i.fn).length === 6 && SD.imgs.some(i => i.fn === 'tqnMenuNewDoc'));
T('N03 Vị trí nút nằm trong ô neo (lệch ngang nhỏ hơn bề rộng cột)', SD.imgs.every(i => i.ox >= 0 && i.ox < [40, 230, 330, 100, 110, 40, 20, 90, 90, 80, 80][i.c - 1]), SD.imgs.map(i => i.c + '+' + i.ox).join(' '));
vm.runInContext('tqnGanHangNutVaoSo()', ctx);
T('N04 Bấm lại không nhân đôi nút', SD.imgs.length === 7 && RFA.imgs.length === 7, SD.imgs.length + '/' + RFA.imgs.length);
console.log(`\nHàng nút: ${p} PASS, ${f} FAIL`); process.exit(f ? 1 : 0);
