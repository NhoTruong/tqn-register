// Test QuanLy.gs: lõi mẫu phiếu + luật, và gắn vào sổ giả lập theo tiêu đề thật RFA/RFI (19/11/2026).
const fs = require('fs'), path = require('path'), vm = require('vm');
const D = n => path.join(__dirname, '../apps-script', n);
let p = 0, f = 0; const T = (t, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + t + (i !== undefined ? ' | ' + i : '')); ok ? p++ : f++; };
const live = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture_live.json'), 'utf8'));
function col(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function mkSheet(name, rows) {
  const sh = { name, rows, links: {}, notes: {}, bg: {}, prot: [], maxC: Math.max(...rows.map(r => r.length)),
    getName: () => name, getLastRow: () => { for (let i = rows.length; i > 0; i--) if ((rows[i - 1] || []).some(x => x !== '' && x != null)) return i; return 0; },
    getLastColumn: () => { let m = 0; rows.forEach(r => r.forEach((x, j) => { if (x !== '' && x != null) m = Math.max(m, j + 1); })); return m; },
    getMaxColumns: () => sh.maxC, getMaxRows: () => 1000, insertColumnsAfter(a, n) { sh.maxC += n; },
    setColumnWidth() {}, getConditionalFormatRules: () => [], setConditionalFormatRules(r) { sh.cf = r; },
    getProtections: () => [], hidden: [], hideColumns(a, n) { sh.hidden.push([a, n]); }, showColumns(a, n) { (sh.shown = sh.shown || []).push([a, n]); },
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      const get = () => { const o = []; for (let i = 0; i < nr; i++) { const a = []; for (let j = 0; j < nc; j++) { const v = (rows[r - 1 + i] || [])[c - 1 + j]; a.push(v == null ? '' : v); } o.push(a); } return o; };
      const put = v => { v.forEach((x, i) => { while (rows.length < r + i) rows.push([]); x.forEach((y, j) => rows[r - 1 + i][c - 1 + j] = y); }); };
      const R = { getValues: get, getDisplayValues: () => get().map(x => x.map(String)), getValue: () => get()[0][0],
        setValue(v) { put([[v]]); return R; }, setValues(v) { put(v); return R; },
        getRichTextValue: () => ({ getLinkUrl: () => sh.links[r + ':' + c] || null }),
        getRichTextValues: () => get().map((x, i) => [{ getLinkUrl: () => sh.links[(r + i) + ':' + c] || null }]),
        setRichTextValue(rt) { put([[rt.text]]); sh.links[r + ':' + c] = rt.url; return R; },
        setNote(n) { sh.notes[r + ':' + c] = n; return R; }, setBackground(b) { sh.bg[r + ':' + c] = b; return R; },
        setFontColor() { return R; }, setFontWeight() { return R; }, setWrap() { return R; }, setDataValidation(dv) { sh.dv = (sh.dv || []).concat([[c, dv]]); return R; },
        protect() { const pr = { setDescription() { return pr; }, setWarningOnly(w) { pr.w = w; return pr; } }; sh.prot.push([c, pr]); return pr; },
        getColumn: () => c, getRow: () => r, getLastRow: () => r + nr - 1, getLastColumn: () => c + nc - 1, getSheet: () => sh,
        getA1Notation: () => col(c) + r };
      return R;
    } };
  return sh;
}
const clone = n => live[n].values.slice(0, 20).map(r => r.map(x => x == null ? '' : (x && x.__d ? new Date(x.__d) : x)));
const RFA = mkSheet('RFA', clone('RFA')), RFI = mkSheet('RFI', clone('RFI'));
const sheets = { RFA, RFI };
let toast = '', triggers = [], audit = [];
const ss = { getSheetByName: n => sheets[n] || null, toast: t => { toast = t; } };
const rt = () => { const o = { setText(t) { o.text = t; return o; }, setLinkUrl(u) { o.url = u; return o; }, build: () => o }; return o; };
const dvb = () => { const o = { requireValueInList(l) { o.list = l; return o; }, setAllowInvalid() { return o; }, setHelpText() { return o; }, build: () => o }; return o; };
const cfb = () => { const o = { getBooleanCondition: () => ({ getCriteriaValues: () => [o.f] }), getRanges: () => [], whenFormulaSatisfied(f) { o.f = f; return o; }, whenTextStartsWith() { return o; }, whenTextEqualTo() { return o; }, whenTextContains() { return o; }, setBold() { return o; }, setBackground() { return o; }, setFontColor() { return o; }, setRanges() { return o; }, build: () => o }; return o; };
const ctx = { console, Math, JSON, String, Date, parseInt, RegExp, Object, isNaN,
  Utilities: { formatDate: (d, tz, f) => f === 'yyMMdd' ? '261119' : '17:55 19/11/2026' }, Session: { getActiveUser: () => ({ getEmail: () => '' }), getScriptTimeZone: () => 'Asia/Ho_Chi_Minh', getEffectiveUser: () => ({ getEmail: () => 'nho@x.com' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger() {}, newTrigger: fn => ({ forSpreadsheet: () => ({ onEdit: () => ({ create() { triggers.push(fn); } }) }) }) },
  SpreadsheetApp: { newRichTextValue: rt, newDataValidation: dvb, newConditionalFormatRule: cfb, ProtectionType: { SHEET: 1, RANGE: 2 } } };
vm.createContext(ctx);
vm.runInContext(['Config', 'Read', 'QuanLy', 'WebApp'].map(n => fs.readFileSync(D(n + '.gs'), 'utf8')).join('\n'), ctx);
const mau = mkSheet('TQN_MAU_PHIEU', [['h']]), log = mkSheet('TQN_NHAT_KY', [['h']]);
Object.assign(ctx, { tqnSs_: () => ss, tqnAudit_: a => audit.push(a),
  tqnConfig_: () => ({ detailSheets: ['RFA', 'RFI', 'SD'], alias: { SD: 'SDS', RFA: 'RFA', RFI: 'RFI' }, hdrFrom: 1, hdrTo: 5, firstRow: 6 }),
  tqnCodeCfg_: () => ({ detailSheets: ['RFA', 'RFI', 'SD'], alias: {}, hdrFrom: 1, hdrTo: 5, firstRow: 6, timezone: 'Asia/Ho_Chi_Minh' }),
  tqnOwnSheet_: (n, h) => { const s = n === 'TQN_NHAT_KY' ? log : mau; s.rows[0] = h.slice(); sheets[n] = s; return s; } });
const run = c => vm.runInContext(c, ctx);

// ---------- 1. Lõi ----------
const p1 = run('tqnQlDocMau_(TQN_MAU_PHIEU_DEFAULT)'); ctx.P1 = p1;
T('Q00a Mặc định dự án KHÔNG tư vấn: RFA/RFC/SDS chỉ 1 cấp CĐT, RFI "CĐT trả lời"', ['RFA', 'RFC', 'SDS'].every(t => p1[t].levels.length === 1 && p1[t].levels[0].name === 'CĐT') && p1.RFI.levels[0].name === 'CĐT trả lời' && !p1._loi.length);
ctx.R = { noiDung: 'x', daTrinh: true, picks: [], tinhTrang: '' };
T('Q00b Không tư vấn: đã trình, chưa kết quả → ⏳ CĐT', run('tqnQlDanhGia_(R,P1.RFA)').giu === '⏳ CĐT');
ctx.R.picks = ['C']; T('Q00c CĐT C → ↩ Nhà thầu sửa & trình lại', /^↩ Nhà thầu – sửa & trình lại \(CĐT C\)/.test(run('tqnQlDanhGia_(R,P1.RFA)').giu));
ctx.R.picks = ['A']; ctx.R.link = 'https://a.sharepoint.com/x'; ctx.R.tinhTrang = 'Đã duyệt'; const d0 = run('tqnQlDanhGia_(R,P1.RFA)');
T('Q00d CĐT A + link → ✔ Đóng, không cảnh báo', d0.giu === '✔ Đóng' && !d0.canhBao.length);
run('TQN_MAU_PHIEU_DEFAULT = TQN_MAU_PHIEU_2CAP'); // các test dưới kiểm mẫu 2 cấp (dự án có tư vấn)
const prof = run('tqnQlDocMau_(TQN_MAU_PHIEU_DEFAULT)');
T('Q01 Mẫu phiếu mặc định: RFA/RFC/SDS 2 cấp TVGS→CĐT, RFI 1 cấp có chuyển tiếp, 0 lỗi',
  ['RFA', 'RFC', 'SDS'].every(t => prof[t].levels.map(l => l.name).join('>') === 'TVGS>CĐT') && prof.RFI.levels.length === 1 && prof.RFI.forward && prof._loi.length === 0);
T('Q02 Nguyên văn B đúng phiếu', prof.RFA.levels[0].options[1].text === 'Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới');
ctx.P = prof;
const dg = (rec, t) => { ctx.R = Object.assign({ noiDung: 'Hồ sơ X', daTrinh: true, tinhTrang: '', picks: [], chuyenToi: '', link: '' }, rec); return run('tqnQlDanhGia_(R, P.' + (t || 'RFA') + ')'); };
const OD = 'https://coastal-my.sharepoint.com/:b:/g/personal/x/abc';
let d;
d = dg({ daTrinh: false }); T('Q03 Chưa trình → ✎ Nhà thầu – đang soạn', d.giu === '✎ Nhà thầu – đang soạn' && !d.canhBao.length, d.giu);
d = dg({}); T('Q04 Đã trình, chưa kết quả → ⏳ TVGS', d.giu === '⏳ TVGS', d.giu);
d = dg({ picks: ['A'] }); T('Q05 TVGS A (gõ tắt "A") → ⏳ CĐT', d.giu === '⏳ CĐT' && !d.canhBao.length, d.giu + ' ' + d.canhBao);
d = dg({ picks: ['B · Chấp thuận yêu cầu với điều kiện hoặc ý kiến phản hồi bên dưới', 'A · Chấp thuận yêu cầu'], tinhTrang: 'Đã duyệt', link: OD });
T('Q06 TVGS B + CĐT A + link + Đã duyệt → ✔ Đóng, không cảnh báo', d.giu === '✔ Đóng' && d.nhom === 'DONG' && !d.canhBao.length, d.canhBao.join(';'));
d = dg({ picks: ['A', 'B'], tinhTrang: 'Đã trình duyệt' });
T('Q07 CĐT B, thiếu link, Tình trạng chưa cập nhật → 2 cảnh báo', d.giu === '✔ Đóng (có ý kiến)' && d.canhBao.length === 2 && /Thiếu link/.test(d.canhBao.join()) && /nên cập nhật/.test(d.canhBao.join()), d.canhBao.join('; '));
d = dg({ picks: ['C'], tinhTrang: 'Đã trình duyệt' });
T('Q08 TVGS C → ↩ Nhà thầu sửa & trình lại, không chờ CĐT', /^↩ Nhà thầu – sửa & trình lại \(TVGS C\)/.test(d.giu), d.giu);
d = dg({ picks: ['D', 'A'] });
T('Q09 TVGS D mà CĐT vẫn A → cảnh báo sai thứ tự', /đã trả về ở cấp trước/.test(d.canhBao.join()), d.canhBao.join('; '));
d = dg({ picks: ['', 'A'] });
T('Q10 CĐT có, TVGS trống → cảnh báo', /CĐT có kết quả nhưng TVGS còn trống/.test(d.canhBao.join()), d.canhBao.join('; '));
d = dg({ picks: ['Duyệt'] });
T('Q11 Gõ tay "Duyệt" không có trong phiếu → cảnh báo', /không có trong mẫu phiếu/.test(d.canhBao.join()), d.canhBao.join('; '));
d = dg({ picks: ['C'], tinhTrang: 'Đã duyệt' });
T('Q12 Phiếu trả về mà Tình trạng "Đã duyệt" → cảnh báo trái nhau', /trái với kết quả phiếu \(bị trả về\)/.test(d.canhBao.join()), d.canhBao.join('; '));
d = dg({ tinhTrang: 'Đã hoàn thành', daTrinh: false });
T('Q13 Dữ liệu cũ chưa có A–D, Tình trạng Đã hoàn thành → ✔ Đóng (theo Tình trạng)', d.giu === '✔ Đóng (theo Tình trạng)', d.giu);
d = dg({ tinhTrang: 'Không duyệt' });
T('Q14 Dữ liệu cũ "Không duyệt" → ↩ Nhà thầu (theo Tình trạng)', d.giu === '↩ Nhà thầu (theo Tình trạng)', d.giu);
d = dg({ picks: ['A'], daTrinh: false });
T('Q15 Có kết quả mà chưa có ngày trình → cảnh báo', /chưa có ngày trình/.test(d.canhBao.join()));
d = dg({ picks: ['A', 'A'], tinhTrang: 'Đã duyệt', link: 'https://drive.google.com/x' });
T('Q16 Link Google Drive (không phải OneDrive) → cảnh báo', /không phải OneDrive/.test(d.canhBao.join()));
d = dg({ picks: ['3 · Chuyển tiếp tới'] }, 'RFI');
T('Q17 RFI chuyển tiếp chưa ghi bên nhận → ⏳ + cảnh báo', /^⏳ bên được chuyển/.test(d.giu) && /Chuyển tới ai/.test(d.canhBao.join()), d.giu);
d = dg({ picks: ['Chuyển tiếp tới'], chuyenToi: 'TVTK' }, 'RFI');
T('Q18 RFI chuyển tiếp TVTK (gõ nguyên văn) → ⏳ TVTK', d.giu === '⏳ TVTK' && !d.canhBao.length, d.giu);
d = dg({ picks: ['2'] }, 'RFI'); T('Q19 RFI "Vui lòng cung cấp thêm thông tin" → ↩ Nhà thầu bổ sung', d.giu === '↩ Nhà thầu – bổ sung thông tin', d.giu);
d = dg({ noiDung: '' }); T('Q20 Dòng nhóm (không nội dung) → để trống', d.giu === '');
// Dự án khác: mẫu phiếu riêng (ví dụ 3 cấp, mã chữ khác) – không sửa code
ctx.M2 = [['RFA', 1, 'PMC', 'NO', 'No exception taken', 'TIEP'], ['RFA', 1, 'PMC', 'R', 'Revise and resubmit', 'TRINH_LAI'],
  ['RFA', 2, 'TVTK', 'OK', 'Đạt', 'TIEP'], ['RFA', 3, 'CĐT', 'AP', 'Approved', 'DONG'], ['RFA', 3, 'CĐT', 'RJ', 'Rejected', 'TU_CHOI'], ['RFA', 3, 'CĐT', 'X', 'Sai', 'BAY']];
const p2 = run('tqnQlDocMau_(M2)'); ctx.P2 = p2;
T('Q21 Dự án khác: mẫu 3 cấp đọc được, dòng kết quả sai bị báo lỗi', p2.RFA.levels.map(l => l.name).join('>') === 'PMC>TVTK>CĐT' && p2._loi.length === 1, p2._loi.join());
ctx.R = { noiDung: 'x', daTrinh: true, picks: ['NO', 'OK'], tinhTrang: '' };
T('Q22 Dự án khác: PMC NO + TVTK OK → ⏳ CĐT', run('tqnQlDanhGia_(R,P2.RFA)').giu === '⏳ CĐT');

// ---------- 2. Gắn vào sổ RFA/RFI giả lập (tiêu đề thật) ----------
const out = run('tqnCaiQuanLy()');
const hdrR = RFA.rows[2], sub = RFA.rows[3];
const st = hdrR.findIndex(x => String(x).indexOf('TQN QUẢN LÝ') === 0) + 1;
T('Q23 RFA: khối TQN QUẢN LÝ thêm bên phải, cách 1 cột, không đè cột cũ', st === 27, 'cột ' + col(st) + ' / ' + out.join(' | '));
T('Q24 RFA: 5 cột TVGS · theo phiếu | CĐT · theo phiếu | Ai đang giữ | Link OneDrive | Kiểm tra',
  sub.slice(st - 1, st + 4).join('|') === 'TVGS · theo phiếu|CĐT · theo phiếu|Ai đang giữ|Link OneDrive|Kiểm tra', sub.slice(st - 1, st + 4).join('|'));
const rs = RFI.rows[2].findIndex(x => String(x).indexOf('TQN QUẢN LÝ') === 0) + 1;
T('Q25 RFI: có cột Trả lời RFI + Chuyển tới', RFI.rows[3].slice(rs - 1, rs + 1).join('|') === 'Trả lời RFI · theo phiếu|Chuyển tới');
T('Q26 Dropdown đúng nguyên văn phiếu', RFA.dv.some(x => x[1].list && x[1].list[2] === 'C · Vui lòng xem ý kiến phản hồi và đệ trình lại'));
T('Q27 Trình kích hoạt "khi sửa" cài 1 lần', triggers.filter(x => x === 'tqnQuanLyKhiSua').length === 1);
// dữ liệu thật: dòng có nội dung đều có Ai đang giữ
const dataRows = RFA.rows.slice(5).filter(r => String(r[2] || '').trim());
T('Q28 RFA thật: mọi dòng có nội dung đều có "Ai đang giữ"', dataRows.length > 0 && dataRows.every(r => String(r[st + 1] || '') !== ''), dataRows.map(r => r[4] + '→' + r[st + 1]).join(' ; '));
T('Q29 Không ghi vào cột gốc A..Z', JSON.stringify(RFA.rows.map(r => r.slice(0, 26))) === JSON.stringify(clone('RFA').map(r => r.slice(0, 26))));

// ---------- 3. Có người sửa ----------
const r0 = 6 + RFA.rows.slice(5).findIndex(r => String(r[2] || '').trim());
RFA.rows[r0 - 1][st - 1] = 'A · Chấp thuận yêu cầu';
ctx.E = { range: RFA.getRange(r0, st), oldValue: undefined, user: { getEmail: () => 'thuky@x.com' } };
run('tqnQuanLyKhiSua(E)');
T('Q30 Nhật ký ghi: người sửa, ô, mã hồ sơ, cột TVGS', log.rows.some(r => r[1] === 'thuky@x.com' && r[3] === col(st) + r0 && /TVGS · theo phiếu/.test(r[5]) && r[7] === 'A · Chấp thuận yêu cầu'), JSON.stringify(log.rows[1]));
T('Q31 Sau khi chọn TVGS A → Ai đang giữ đổi thành ⏳ CĐT (nếu đã trình)', /⏳ CĐT|✎|⤳ Đã thay/.test(RFA.rows[r0 - 1][st + 1]), RFA.rows[r0 - 1][st + 1]);
// dán link OneDrive
RFA.rows[r0 - 1][st + 2] = OD;
ctx.E = { range: RFA.getRange(r0, st + 3), oldValue: '', user: { getEmail: () => 'thuky@x.com' } };
run('tqnQuanLyKhiSua(E)');
T('Q32 Dán link OneDrive → ô hiện "📎 OneDrive" + gắn link vào ô Mã', RFA.rows[r0 - 1][st + 2] === '📎 OneDrive' && RFA.links[r0 + ':' + (st + 3)] === OD && RFA.links[r0 + ':2'] === OD, RFA.rows[r0 - 1][st + 2]);
RFA.rows[r0 - 1][st + 2] = 'https://drive.google.com/abc';
delete RFA.links[r0 + ':' + (st + 3)];
run('tqnQuanLyKhiSua(E)');
T('Q33 Dán link không phải OneDrive → tô đỏ + ghi chú + nhật ký LINK_SAI', RFA.bg[r0 + ':' + (st + 3)] === '#FDE2E1' && log.rows.some(r => r[8] === 'LINK_SAI'));
// email không lấy được
ctx.E = { range: RFA.getRange(r0, 3), oldValue: 'cũ' };
run('tqnQuanLyKhiSua(E)');
T('Q34 Không lấy được email → ghi rõ, không bỏ trống', log.rows.some(r => /không lấy được/.test(r[1]) && r[6] === 'cũ'));
// tab TQN_ bị bỏ qua
const n0 = log.rows.length; ctx.E = { range: mau.getRange(2, 1) }; run('tqnQuanLyKhiSua(E)');
T('Q35 Sửa tab TQN_* → không ghi nhật ký', log.rows.length === n0);
T('Q36 Cột Ai đang giữ / Kiểm tra: khóa kiểu cảnh báo', RFA.prot.some(x => x[0] === st + 2 && x[1].w === true));

// ---------- 4. Phân quyền ----------
ctx.L = [['nho@x.com', 'Chủ', 'Có'], ['ThuKy@x.com', 'Thư ký', 'có'], ['bptc@x.com', 'BPTC', 'Không'], ['sai-email', '', 'Có'], ['', '', '']];
const pq = run('tqnPqDanhSach_(L)');
T('Q37 Phân quyền: chỉ email "Có" được sửa cột kết quả, email sai bị báo', JSON.stringify(pq.ok) === '["nho@x.com","thuky@x.com"]' && pq.loi.length === 1, JSON.stringify(pq));
// ---------- 5. AI đọc mẫu phiếu (câu trả lời giả) ----------
ctx.AIT = 'Đây là kết quả:\n{"cap":[{"ten":"TVGS","thu_tu":1,"lua_chon":[{"ma":"A","nguyen_van":"Chấp thuận yêu cầu","ket_qua":"TIEP"},{"ma":"C","nguyen_van":"Vui lòng xem ý kiến phản hồi và đệ trình lại","ket_qua":"TRINH_LAI"}]},{"ten":"CĐT","thu_tu":2,"lua_chon":[{"ma":"A","nguyen_van":"Chấp thuận yêu cầu","ket_qua":"DONG"}]}],"ghi_chu":"Ô D mờ"}';
const ai = run('tqnAiMauPhieuParse_(AIT,"RFA")');
T('Q38 AI đọc phiếu: JSON → 3 dòng mẫu phiếu đúng thứ tự cấp, giữ ghi chú chưa chắc', ai.rows.length === 3 && ai.rows[2][2] === 'CĐT' && ai.rows[2][5] === 'DONG' && !ai.loi.length && ai.ghiChu === 'Ô D mờ', JSON.stringify(ai.loi));
ctx.AIT = '{"cap":[{"ten":"TVGS","lua_chon":[{"ma":"A","nguyen_van":"OK","ket_qua":"DUYET"}]}]}';
T('Q39 AI trả mã kết quả sai → bị báo lỗi, không dùng được', run('tqnAiMauPhieuParse_(AIT,"RFA")').loi.length === 1);
ctx.AIT = 'Xin lỗi tôi không đọc được';
T('Q40 AI không trả JSON → báo lỗi', /không trả JSON/.test(run('tqnAiMauPhieuParse_(AIT,"RFA")').loi[0]));
// ---------- 6. Trang web ----------
ctx.R = { noiDung: 'x', daTrinh: true, picks: ['A', 'A'], tinhTrang: 'Đã trình duyệt', link: '' };
let kw = run('tqnWebKiemLuu_(R,P.RFA)');
T('Q41 Web: đóng mà thiếu link → CHẶN; Tình trạng chưa cập nhật → chỉ cảnh báo', kw.chan.length === 1 && /Thiếu link/.test(kw.chan[0]) && kw.canh.length === 1, JSON.stringify(kw));
const wd = run('tqnWebDuLieu()');
T('Q42 Web đọc dữ liệu: có RFA + RFI, mỗi dòng có Ai đang giữ', wd.loai.RFA && wd.loai.RFI && wd.rows.length >= 5 && wd.rows.every(r => r.giu), wd.rows.length + ' dòng');
fs.writeFileSync(path.join(__dirname, 'web_preview_data.json'), JSON.stringify(wd));
const rw = wd.rows.find(r => r.tab === 'RFA' && /Đã trình/.test(r.tt));
ctx.O = { tab: 'RFA', row: rw.row, ma: rw.ma, picks: ['A · Chấp thuận yêu cầu', 'A · Chấp thuận yêu cầu'], link: '' };
let er = ''; try { run('tqnWebLuu(O)'); } catch (e) { er = e.message; }
T('Q43 Web Lưu: CĐT A mà không có link → không lưu, báo lỗi tiếng Việt', /Chưa lưu: Thiếu link/.test(er) && !RFA.rows[rw.row - 1][st], er);
ctx.O.link = OD; const nl = log.rows.length;
const wr = run('tqnWebLuu(O)');
T('Q44 Web Lưu đủ link → ✔ Đóng, Tình trạng tự đổi Đã duyệt, ngày tự điền, nhật ký, gắn link', wr.giu === '✔ Đóng' && log.rows.length > nl && log.rows.slice(nl).every(r => /WEB|CẢNH BÁO|TỰ ĐIỀN|TỰ ĐỔI/.test(r[8])) && RFA.rows[rw.row - 1][4] === 'Đã duyệt' && !wr.canh.length && RFA.links[rw.row + ':' + (st + 3)] === OD, JSON.stringify(wr));
ctx.O = { tab: 'RFA', row: rw.row, ma: 'MÃ-KHÁC', picks: [] }; er = ''; try { run('tqnWebLuu(O)'); } catch (e) { er = e.message; }
T('Q45 Web: dòng đã bị người khác đổi mã → không ghi đè, báo tải lại', /tải lại/.test(er));
// ---------- 7. Chạy mạnh trong sheet ----------
const r9 = 9; // RFA dòng 9: COS-002-rev02, Không duyệt, đã trình 16/11
RFA.rows[r9 - 1][st - 1] = 'A · Chấp thuận yêu cầu'; RFA.rows[r9 - 1][st] = 'A · Chấp thuận yêu cầu';
ctx.E = { range: RFA.getRange(r9, st, 1, 2), user: { getEmail: () => 'thuky@x.com' } }; run('tqnQuanLyKhiSua(E)');
T('Q46 Sửa thẳng trên sheet: chọn TVGS A + CĐT A → Tình trạng tự thành "Đã duyệt"', RFA.rows[r9 - 1][4] === 'Đã duyệt', RFA.rows[r9 - 1][4]);
T('Q47 … ô Thực tế của TVGS (L) và CĐT (P) tự điền ngày hôm nay, ghi nhật ký TỰ ĐIỀN/TỰ ĐỔI', RFA.rows[r9 - 1][11] instanceof Date && RFA.rows[r9 - 1][15] instanceof Date && log.rows.filter(r => /TỰ ĐIỀN|TỰ ĐỔI/.test(r[8])).length >= 3);
T('Q48 Ai đang giữ = ✔ Đóng, Kiểm tra báo thiếu link (chưa dán)', /^✔ Đóng/.test(RFA.rows[r9 - 1][st + 1]) && /Thiếu link/.test(RFA.rows[r9 - 1][st + 3]), RFA.rows[r9 - 1][st + 1] + ' | ' + RFA.rows[r9 - 1][st + 3]);
const nTo = (ctx.sheets_RFA = RFA, run('tqnQlToDong_(sheets_RFA, tqnConfig_(), tqnQlKhoi_(sheets_RFA, tqnConfig_(), P.RFA, false))'));
T('Q49 Tô màu cả dòng theo Ai đang giữ: 3 quy tắc, công thức trỏ đúng cột', nTo === 3 && RFA.cf.length === 3 && RFA.cf[0].f === '=LEFT($' + col(st + 2) + '6,1)="↩"', RFA.cf[0].f);
T('Q50 Mã rev mới: đổi ngày + tăng rev, bỏ dấu cách thừa', run('tqnQlMaRevMoi_("261116- COA-TBC-RFI-003 rev01","261119",2)') === '261119-COA-TBC-RFI-003 rev02' && run('tqnQlMaRevMoi_("261111-COA-TBC-RFA-COS-002-rev01","261119",2)') === '261119-COA-TBC-RFA-COS-002-rev02' && run('tqnQlMaRevMoi_("261117-COA.NP-NT-SDS-011","261119",1)') === '261119-COA.NP-NT-SDS-011');
ctx.sheets_RFA = RFA; const tl = run('tqnTrinhLai(sheets_RFA, 7)');
T('Q51 Trình lại dòng 7 → dòng trống kế tiếp, mã rev02 ngày mới, STT nối tiếp, Rev 02, dòng cũ có ghi chú', tl.row === 10 && tl.ma === '261119-COA-TBC-RFA-COS-002-rev02' && RFA.rows[9][5] === '02' && RFA.notes['7:2'] && /Đang thực hiện/.test(RFA.rows[9][4]), JSON.stringify(tl) + ' STT=' + RFA.rows[9][0]);
const req = run('tqnQlBoLocReq_(1, tqnQlKhoi_(sheets_RFA, tqnConfig_(), P.RFA, false), ["Mr. Thịnh"], 6, 1000, 31)');
T('Q52 Bộ lọc sẵn: 4 tình huống + "Việc của Mr. Thịnh", lọc đúng cột', req.length === 5 && req[4].addFilterView.filter.title === 'TQN · Việc của Mr. Thịnh' && req[4].addFilterView.filter.criteria[3].condition.values[0].userEnteredValue === 'Mr. Thịnh');
ctx.PB = [['thinh@x.com', 'Shop', 'Không', '', 'Mr. Thịnh', 'Có'], ['sep@x.com', 'Sếp – Chỉ huy trưởng', '', '', '', ''], ['nghi@x.com', 'QC', '', '', 'Mr. Nghị', 'Không'], ['sai', '', '', '', 'X', '']];
const db = run('tqnEmailDanhBa_(PB)');
T('Q53 Danh bạ email: tên → email, sếp nhận tóm tắt, người "Không" bị bỏ', db.ten['mr. thinh'] === 'thinh@x.com' && db.sep[0] === 'sep@x.com' && !db.ten['mr. nghi'] && db.tatCa.length === 2, JSON.stringify(db));
ctx.DB = db; ctx.RW = [{ nguoi: 'Mr. Thịnh', giu: '↩ Nhà thầu', kt: '✔' }, { nguoi: 'Mr. Thịnh', giu: '⏳ CĐT', kt: '✔' }, { nguoi: 'Mr. Thịnh', giu: '✔ Đóng', kt: '⚠ Thiếu link' }, { nguoi: 'Ai đó', giu: '↩', kt: '' }];
const gm = run('tqnEmailGom_(RW, DB)');
T('Q54 Email 7h: Mr. Thịnh nhận 2 việc (bị trả về + cần sửa), không gửi việc đang chờ bên khác', gm['thinh@x.com'].length === 2 && Object.keys(gm).length === 1);
ctx.PropertiesService = { getDocumentProperties: () => ({ setProperty() {} }) };
const an = run('tqnAnHienTuVan(true)');
T('Q55 Ẩn khối Tư vấn: RFA ẩn TVGS (L–O) + TVTK (U–Z), RFI ẩn TVGS + TVTK, không xóa dữ liệu', JSON.stringify(RFA.hidden) === '[[12,4],[21,6]]' && RFI.hidden.length === 2 && an.length === 2, JSON.stringify(RFA.hidden) + ' ' + JSON.stringify(RFI.hidden));
T('Q56 Ẩn tư vấn thì hiện khối CĐT (sổ gốc ẩn sẵn P–T)', JSON.stringify(RFA.shown || []).indexOf('[16,') >= 0, JSON.stringify(RFA.shown));
ctx.KK = { map: { nt: { actual: 0, plan: 1 }, tvgs: null }, han: { 'CĐT': { ngay: 1 } } };
ctx.PP = { levels: [{ name: 'CĐT' }] };
const hn = new Date(2026, 10, 19);
ctx.HN = hn; ctx.R1 = [new Date(2026, 10, 16), '']; ctx.R2 = [new Date(2026, 10, 18), ''];
const t1 = run("tqnQlTre_('⏳ CĐT', R1, KK, PP, HN)"), t2 = run("tqnQlTre_('⏳ CĐT', R2, KK, PP, HN)"), t3 = run("tqnQlTre_('↩ Nhà thầu', R1, KK, PP, HN)");
T('Q57 Trễ hạn: trình 16/11, CĐT max 1 ngày → hôm 19/11 trễ 2; trình 18/11 → chưa trễ; đang ở nhà thầu → không tính', t1 === 2 && t2 === 0 && t3 === 0, t1 + ',' + t2 + ',' + t3);
const tr = run("JSON.stringify([tqnQlTachRev_('261116-COA-TBC-RFA-COS-001-rev02'), tqnQlTachRev_('261118-COA-TBC-RFI-001 rev00'), tqnQlRevMax_(['261111-COA-TBC-RFA-COS-001-rev01','261116-COA-TBC-RFA-COS-001-rev02'])])");
T('Q58 Nối rev: bỏ ngày đầu + đuôi rev, rev01 có rev02 → rev cũ', tr === '[{"goc":"COA-TBC-RFA-COS-001","rev":2},{"goc":"COA-TBC-RFI-001","rev":0},{"COA-TBC-RFA-COS-001":2}]', tr);
console.log(`\nQuanLy: ${p} PASS, ${f} FAIL`); process.exit(f ? 1 : 0);
