// Test TuDongMa.gs: dữ liệu giống tab SD và RFI thật (19/11/2026).
const fs = require('fs'), path = require('path'), vm = require('vm');
const D = n => path.join(__dirname, '../apps-script', n);
function mkSheet(name, rows) {
  const sh = { name, rows, writes: [], getName: () => name, getLastRow: () => rows.length, getLastColumn: () => 12,
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      const R = { getValues: () => rows.slice(r - 1, r - 1 + nr).map(x => { const a = x.slice(c - 1, c - 1 + nc); while (a.length < nc) a.push(''); return a; }),
        getValue: () => (rows[r - 1] || [])[c - 1] ?? '',
        setValue(v) { sh.writes.push([r, c, v]); rows[r - 1][c - 1] = v; return R; },
        setValues(v) { v.forEach((x, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; x.forEach((y, j) => rows[r - 1 + i][c - 1 + j] = y); }); return R; },
        setNumberFormat() { return R; }, setNote(n) { sh.notes = sh.notes || {}; sh.notes[r + ':' + c] = n; return R; }, setBackground(b) { sh.bg = sh.bg || {}; sh.bg[r + ':' + c] = b; return R; }, getBackground() { return ((sh.bg || {})[r + ':' + c] || '#ffffff').toLowerCase(); } };
      return R;
    }, appendRow(x) { rows.push(x); } };
  return sh;
}
const hdr = [['', '', '', '', '', ''], ['', '', '', '', '', ''], ['STT', 'Mã hồ sơ', 'Nội dung chi tiết', 'Người thực hiện', 'Tình trạng hồ sơ', 'Rev'], ['', '', '', '', '', ''], ['I', 'KẾT CẤU', '', '', '', '']];
const SD = mkSheet('SD', hdr.map(r => r.slice()).concat([
  [1, '261117-COA.NP-NT-SDS-001', 'Bản vẽ thép móng RE4', 'Nho', 'Trả về', '00'],
  [2, '261117-COA.NP-NT-SDS-002', 'Bản vẽ thép móng RE5', 'Nho', '', '00'],
  [6, '261117-COA.NP-NT-SDS-011', 'Bản vẽ sàn tầng 2 RE4', 'Nho', '', ''],
  ['', '', 'Bản vẽ sàn tầng 2 RE5', 'Nho', '', ''],
  ['', '', '', '', '', '']]));
const RFIh = hdr.map(r => r.slice()); RFIh[4] = ['I', 'PHẦN THÂN', '', '', '', ''];
const RFI = mkSheet('RFI', RFIh.concat([
  [1, '261116-COA-TBC-RFI-001 rev01', 'Gạch porcelain', 'Mr. Thịnh', '', '01'],
  [3, '261116- COA-TBC-RFI-003 rev01', 'Cấu tạo nền công viên', 'Mr. Thịnh', '', '01'],
  [4, '', 'Hỏi cao độ hố ga', '', '', ''],
  [5, '', '', '', '', '']]));
const MAT = mkSheet('MAT', hdr.map(r => r.slice()).concat([[1, '', 'Vật tư thép', '', '', '']]));
const capMa = mkSheet('TQN_CAP_MA', [['h']]);
const sheets = { SD, RFI, MAT, TQN_CAP_MA: capMa, TQN_CONFIG: mkSheet('TQN_CONFIG', [['Khóa', 'Giá trị']]) };
let toast = '', toastTitle = '', audit = [], triggers = [];
const ss = { getSheetByName: n => sheets[n] || null, toast: (t, ti) => { toast = t; toastTitle = ti || ''; } };
const ctx = { console, Math, JSON, String, Date, parseInt, RegExp, Object,
  Utilities: { formatDate: () => '20261119' }, Session: { getActiveUser: () => ({ getEmail: () => 'nho' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger(t) { triggers = triggers.filter(x => x !== t); }, newTrigger: fn => ({ forSpreadsheet: () => ({ onEdit: () => ({ create() { triggers.push({ getHandlerFunction: () => fn }); } }) }) }) },
  SpreadsheetApp: { flush() {} }, PropertiesService: {}, CacheService: {} };
vm.createContext(ctx);
let src = ['Config', 'Read', 'Workflow', 'TuDongMa'].map(n => fs.readFileSync(D(n + '.gs'), 'utf8')).join('\n');
vm.runInContext(src, ctx);
Object.assign(ctx, { tqnSs_: () => ss, tqnAudit_: (a, d) => audit.push(a + ':' + d), tqnCacheClear_() {},
  tqnOwnSheet_: () => capMa });
capMa.hideSheet = () => {};
let p = 0, f = 0; const T = (t, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + t + (i ? ' | ' + i : '')); ok ? p++ : f++; };

// 1. Gõ nội dung dòng mới SD (dòng 9 đã có nội dung, chưa mã) → sửa ô C9
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return this.s}, s:null, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 9}, getLastRow:function(){return 9}}})'.replace('s:null', 's:sheets_SD'), Object.assign(ctx, { sheets_SD: SD }));
T('M01 SD: gõ nội dung → mã tiếp theo SDS-012 (nối từ 011, giữ gói NP, ngày hôm nay)', SD.rows[8][1] === '261119-COA.NP-NT-SDS-012', SD.rows[8][1]);
T('M02 SD: STT trống → 7 (dòng trên 6 + 1)', SD.rows[8][0] === 7, SD.rows[8][0]);
T('M03 SD: mẫu không có rev → không tự điền Rev', SD.rows[8][5] === '', JSON.stringify(SD.rows[8][5]));
T('M04 Ghi sổ TQN_CAP_MA trạng thái TỰ ĐỘNG + toast', capMa.rows.some(r => r[3] === '261119-COA.NP-NT-SDS-012' && r[9] === 'TỰ ĐỘNG') && /SDS-012/.test(toast));
// 2. Gõ lại cùng ô → không cấp lần 2
const nw = SD.writes.length;
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_SD}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 9}, getLastRow:function(){return 9}}})', ctx);
T('M05 Sửa lại dòng đã có mã → không đổi, không cấp trùng', SD.writes.length === nw && SD.rows[8][1] === '261119-COA.NP-NT-SDS-012');
// 3. Dòng trống nội dung → không làm gì
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_SD}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 10}, getLastRow:function(){return 10}}})', ctx);
T('M06 Dòng chưa có nội dung → không cấp mã', SD.rows[9][1] === '');
// 4. RFI: số tiếp theo sau 003 (mã có khoảng trắng "261116- COA") → 004 rev01, Rev = 01
ctx.sheets_RFI = RFI;
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_RFI}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 8}, getLastRow:function(){return 8}}})', ctx);
T('M07 RFI: mã 261119-COA-TBC-RFI-004 rev01', RFI.rows[7][1] === '261119-COA-TBC-RFI-004 rev01', RFI.rows[7][1]);
T('M08 RFI: Rev trống → 01; STT có sẵn 4 → giữ nguyên', RFI.rows[7][5] === '01' && RFI.rows[7][0] === 4, RFI.rows[7][5] + '/' + RFI.rows[7][0]);
// 5. Tab không bật (MAT) → không đụng
ctx.sheets_MAT = MAT;
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_MAT}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 6}, getLastRow:function(){return 6}}})', ctx);
T('M09 Tab MAT chưa bật → không cấp mã', MAT.rows[5][1] === '' && MAT.writes.length === 0);
// 6. Sửa dòng tiêu đề → bỏ qua
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_SD}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 3}, getLastRow:function(){return 3}}})', ctx);
T('M10 Sửa vùng tiêu đề → bỏ qua', SD.writes.length === nw);
// 7. Dán nhiều dòng một lúc → số liên tiếp, không trùng
SD.rows.push(['', '', 'Bản vẽ A', '', '', ''], ['', '', 'Bản vẽ B', '', '', '']);
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_SD}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return 11}, getLastRow:function(){return 12}}})', ctx);
T('M11 Dán 2 dòng → SDS-013, SDS-014', SD.rows[10][1] === '261119-COA.NP-NT-SDS-013' && SD.rows[11][1] === '261119-COA.NP-NT-SDS-014', SD.rows[10][1] + ' / ' + SD.rows[11][1]);
// 8. Không bao giờ ghi đè ô có dữ liệu
const ghiDe = SD.writes.concat(RFI.writes).filter(w => w[1] === 2).length;
T('M12 Chỉ ghi vào ô Mã trống (4 lần ghi mã SD+RFI)', ghiDe === 4, ghiDe);
// 9. Bật: không sinh trigger trùng
vm.runInContext('tqnBatTuDongMa(); tqnBatTuDongMa();', ctx);
T('M13 Bật 2 lần → 1 trình kích hoạt', triggers.filter(t => t.getHandlerFunction() === 'tqnTuDongMaKhiSua').length === 1);
T('M14 Tắt tự cấp mã qua auto_code_types trống', (() => { sheets.TQN_CONFIG.rows.push(['auto_code_types', '']); const c = vm.runInContext('tqnCodeCfg_()', ctx); return vm.runInContext('tqnAutoTypeOfTab_', ctx)(c, 'SD') === ''; })());

// ===== Kiểm & sửa mã nhập tay: dữ liệu thật tab SD 19/11 (rút gọn) =====
const SD2 = mkSheet('SD', hdr.map(r => r.slice()).concat([
  [1, '261117-COA.NP-NT-SDS-001', 'Thép móng RE4', 'Nho', 'Trả về', '00'],
  [2, '261117-COA.NP-NT-SDS-002', 'Thép móng RE5', 'Nho', 'Đã trình duyệt', '00'],
  [2, '261117-COA.NP-NT-SDS-003', 'Thép bể ngầm RE4', 'Nho', 'Đã trình duyệt', '00'],
  [13, '261117-COA.NP-NT-SDS-025', 'Thép thang bộ RE4', 'Nho', '', ''],
  [14, '261117-COA.NP-NT-SDS-026', 'Thép thang bộ RE5', 'Nho', '', ''],
  [13, '261117-COA.NP-NT-SDS-025', 'Thép cột tầng 1 đến mái RE4', 'Nho', '', ''],
  [14, '261117-COA.NP-NT-SDS-026', 'Thép cột tầng 1 đến mái RE5', 'Nho', 'Đã trình duyệt', ''],
  ['', 'KẾT CẤU CÔNG VIÊN', '', '', '', ''],
  [1, '261113-COA.HTKT-NT-SDS-001', 'Mặt dựng', 'Nho', 'Đã hoàn thành', '00'],
  [27, '261114-COA.HTKT-NT-SDS-002', 'Định vị bulong', 'Nho', 'Đã hoàn thành', ''],
  [2, '261113- COA.HTKT-NT-SDS-003', 'Đảo cà phê', 'Nho', '', '']]));
sheets.SD = SD2; capMa.rows.length = 1; sheets.TQN_CONFIG.rows.length = 1;
const allK = vm.runInContext('tqnKiemMaTatCa_()', ctx); console.log(allK.map(k=>k.tab)); const kq = allK.find(k => k.tab === 'SD');
T('M15 Phát hiện SDS-025, SDS-026 trùng (dòng 11, 12)', kq.trung.length === 2 && kq.trung[0].row === 11 && kq.trung[0].dongDau === 9, JSON.stringify(kq.trung.map(t => t.row + '<' + t.dongDau)));
T('M16 Trùng chưa trình → số mới NP-027; đã trình → không đổi, chỉ báo', kq.trung[0].moi === '261119-COA.NP-NT-SDS-027' && kq.trung[1].moi === '' && kq.trung[1].doiDuoc === false, kq.trung.map(t => t.moi || 'giữ').join(' / '));
T('M17 Gói HTKT đánh số riêng: SDS-001 HTKT không tính trùng với NP', kq.trung.every(t => /NP/.test(t.cu)));
T('M18 Dư khoảng trắng "261113- COA" → chuẩn hóa', kq.sai.length === 1 && kq.sai[0].moi === '261113-COA.HTKT-NT-SDS-003', JSON.stringify(kq.sai));
T('M19 STT chạy 1..n từng mục: dòng 8 2→3, dòng 15 27→2, dòng 16 2→3', ['8:3', '15:2', '16:3'].every(x => kq.stt.some(t => t.row + ':' + t.moi === x)) && !kq.stt.some(t => t.row === 14), kq.stt.map(t => t.row + ':' + t.cu + '→' + t.moi).join(' '));
const truoc = SD2.rows.map(r => r.slice());
const nSua = vm.runInContext('tqnSuaMa_(tqnKiemMaTatCa_())', ctx);
T('M20 Sửa: dòng 11 → SDS-027, dòng 12 (đã trình) giữ 026, ghi CAP_MA có mã cũ', SD2.rows[10][1] === '261119-COA.NP-NT-SDS-027' && SD2.rows[11][1] === '261117-COA.NP-NT-SDS-026' && capMa.rows.some(r => r[9] === 'SỬA MÃ TRÙNG' && r[7] === '261117-COA.NP-NT-SDS-025'), nSua + ' ô');
T('M21 Không đụng cột Nội dung / Tình trạng / Rev', SD2.rows.every((r, i) => [2, 3, 4, 5].every(c => r[c] === truoc[i][c])));
const lan2 = vm.runInContext('tqnKiemMaTatCa_()', ctx).find(k => k.tab === 'SD');
T('M22 Kiểm lại sau sửa: chỉ còn 1 trùng đã trình (để anh xử lý tay)', lan2.trung.length === 1 && !lan2.trung[0].moi && lan2.sai.length === 0 && lan2.stt.length === 0, JSON.stringify(lan2.trung.map(t => t.row)));

// ===== Đồng bộ ngày trình + Rev vào mã RFI =====
const hdr3 = [['', '', '', '', '', '', ''], ['', '', '', '', '', '', ''], ['STT', 'Mã hồ sơ', 'Nội dung chi tiết', 'Người thực hiện', 'Tình trạng hồ sơ', 'Rev', 'Nhà thầu đệ trình'], ['', '', '', '', '', '', 'Ngày trình\nduyệt'], ['I', 'PHẦN THÂN', '', '', '', '', '']];
const RFI3 = mkSheet('RFI', hdr3.concat([
  [1, '261116-COA-TBC-RFI-001 rev01', 'Gạch porcelain', 'Mr. Thịnh', 'Đã trình duyệt', '01', new Date(2026, 10, 16)],
  [2, '261116-COA-TBC-RFI-002 rev01', 'Gạch mosaic', 'Mr. Thịnh', 'Đã trình duyệt', '01', new Date(2026, 10, 16)],
  [3, '', '', '', '', '', '']]));
RFI3.getLastColumn = () => 7; sheets.RFI = RFI3; ctx.sheets_RFI3 = RFI3;
ctx.Utilities.formatDate = d => d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
const ev = r => 'tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_RFI3}, getColumn:function(){return 3}, getLastColumn:function(){return 3}, getRow:function(){return ' + r + '}, getLastRow:function(){return ' + r + '}}})';
vm.runInContext(ev(6), ctx);
T('M23 Mã đã khớp ngày trình + Rev → không đổi', RFI3.rows[5][1] === '261116-COA-TBC-RFI-001 rev01' && RFI3.writes.length === 0);
RFI3.rows[6][5] = '02'; RFI3.rows[6][6] = new Date(2026, 10, 20); vm.runInContext(ev(7), ctx);
T('M24 Trình lại: sửa Rev 02 + ngày 20/11 → mã thành 261120-COA-TBC-RFI-002 rev02', RFI3.rows[6][1] === '261120-COA-TBC-RFI-002 rev02', RFI3.rows[6][1]);
T('M25 Ghi CAP_MA "ĐỒNG BỘ NGÀY/REV" kèm mã cũ', capMa.rows.some(r => r[9] === 'ĐỒNG BỘ NGÀY/REV' && r[7] === '261116-COA-TBC-RFI-002 rev01'));
RFI3.rows[7][2] = 'Hỏi cao độ hố ga'; RFI3.rows[7][6] = new Date(2026, 10, 21); vm.runInContext(ev(8), ctx);
T('M26 Dòng mới đã có ngày trình 21/11 → mã lấy ngày đó (số 004 vì 003 đã nằm trong sổ cấp mã)', RFI3.rows[7][1] === '261121-COA-TBC-RFI-004 rev01' && RFI3.rows[7][5] === '01', RFI3.rows[7][1]);
T('M27 SD (không có cột Ngày trình duyệt, mẫu không rev) → không đồng bộ', vm.runInContext('tqnDongBoMa_(sheets_SD, tqnCodeCfg_(), "SDS", 6, 20)', ctx).length === 0);

// ===== Gõ tay vào cột Mã =====
T('M28 Ô mã tự cấp có ghi chú "TQN tự cấp lúc …"', /TQN tự cấp lúc/.test(Object.values(SD.notes || {}).join(' ')));
RFI3.rows.push([5, '261116 - COA-TBC-RFI-001 rev01', 'Gõ tay trùng', '', '', '', '']);
const rr = RFI3.rows.length;
vm.runInContext('tqnTuDongMaKhiSua({range:{getSheet:function(){return sheets_RFI3}, getColumn:function(){return 2}, getLastColumn:function(){return 2}, getRow:function(){return ' + rr + '}, getLastRow:function(){return ' + rr + '}}})', ctx);
T('M29 Gõ tay "261116 - COA…" → tự bỏ khoảng trắng', RFI3.rows[rr - 1][1] === '261116-COA-TBC-RFI-001 rev01', RFI3.rows[rr - 1][1]);
T('M30 Gõ tay trùng số với dòng 6 → tô đỏ + ghi chú + thông báo, KHÔNG tự xóa', RFI3.bg[rr + ':2'] === '#FDE2E1' && /trùng số với dòng 6/.test(RFI3.notes[rr + ':2']) && /kiểm mã gõ tay/.test(toastTitle), RFI3.notes[rr + ':2']);
console.log(`KẾT QUẢ TỰ CẤP MÃ: ${p} PASS / ${f} FAIL`);
