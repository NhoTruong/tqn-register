// Test TrangChu.gs bằng giả lập: tạo đủ nút, gắn đúng hàm, không đụng tab khác.
const fs = require('fs'), path = require('path'), vm = require('vm');
const imgs = [], touched = new Set();
function Rng(sh) { const r = {}; ['merge','setValue','setBackground','setFontColor','setFontWeight','setFontSize','setVerticalAlignment','setHorizontalAlignment'].forEach(m => r[m] = () => { touched.add(sh.name); return r; }); return r; }
const sh = { name: 'TQN Trang chủ', getImages: () => [], getMaxColumns: () => 26, insertColumnsAfter() {}, setColumnWidths() {}, setHiddenGridlines() {}, setTabColor() {}, setRowHeight() {}, getRange() { return Rng(sh); },
  insertImage(b, c, r, ox, oy) { const im = { c, r, ox, oy, fn: '', setWidth(w) { im.w = w; return im; }, setHeight(h) { im.h = h; return im; }, assignScript(f) { im.fn = f; return im; }, setAltTextTitle() { return im; } }; imgs.push(im); return im; } };
const ss = { getSheetByName: () => null, insertSheet: () => sh, setActiveSheet() {}, moveActiveSheet() {} };
const ctx = { Utilities: { newBlob: () => ({}), base64Decode: () => [], formatDate: () => '15:30 19/11' }, SpreadsheetApp: { getActive: () => ({ toast() {} }) },
  tqnSs_: () => ss, tqnConfig_: () => ({ project_name: 'Coastal', timezone: 'Asia/Ho_Chi_Minh' }), tqnAudit_() {}, tqnManagerData: () => ({ kpi: { quaHan: 11, sapHan: 6, cho: 30, traVe: 4, tong: 111 } }), Math, JSON, String };
vm.createContext(ctx); vm.runInContext(fs.readFileSync(path.join(__dirname, '../apps-script/TrangChu.gs'), 'utf8'), ctx);
const n = vm.runInContext('tqnTaoTrangChu()', ctx);
const fns = imgs.filter(i => i.fn).map(i => i.fn);
const menuSrc = fs.readdirSync(path.join(__dirname, '../apps-script')).filter(f => f.endsWith('.gs')).map(f => fs.readFileSync(path.join(__dirname, '../apps-script', f), 'utf8')).join('\n');
let p = 0, f = 0; const T = (t, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + t + (i ? ' | ' + i : '')); ok ? p++ : f++; };
T('T01 Tạo 13 nút bấm được + 5 nhãn nhóm', n === 13 && imgs.length === 18, n + ' nút, ' + imgs.length + ' hình');
T('T02 Mọi nút gắn hàm có thật trong code', fns.every(x => new RegExp('function ' + x + '\\(').test(menuSrc)), fns.join(','));
T('T03 Nút Quản lý hồ sơ đứng đầu hàng 1', imgs.find(i => i.fn)?.fn === 'tqnMenuManager' && imgs[1].r === 3);
T('T04 Hai hàng nút không vượt khung (≤ 72 cột × 25px)', imgs.every(i => i.c <= 72 && i.ox < 25));
T('T05 Chỉ đụng tab Trang chủ', [...touched].every(x => x === 'TQN Trang chủ'), [...touched].join(','));
console.log(`KẾT QUẢ TRANG CHỦ: ${p} PASS / ${f} FAIL`);
