// Test ThaFile.gs: tên file thật → đúng dòng sổ, chuyển thư mục, gắn link, không xóa.
const fs = require('fs'), path = require('path'), vm = require('vm');
const D = n => path.join(__dirname, '../apps-script', n);
function mkSheet(name, rows) {
  const sh = { name, rows, rich: {}, notes: {}, getName: () => name, getLastRow: () => rows.length, getLastColumn: () => 6,
    getRange(r, c, nr, nc) { nr = nr || 1; nc = nc || 1; const R = {
      getValues: () => rows.slice(r - 1, r - 1 + nr).map(x => { const a = x.slice(c - 1, c - 1 + nc); while (a.length < nc) a.push(''); return a; }),
      getDisplayValue: () => String((rows[r - 1] || [])[c - 1] ?? ''),
      setValues(v) { v.forEach((x, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; x.forEach((y, j) => rows[r - 1 + i][c - 1 + j] = y); }); return R; },
      setRichTextValue(v) { sh.rich[r + ':' + c] = v; return R; }, getNote: () => sh.notes[r + ':' + c] || '', setNote(n) { sh.notes[r + ':' + c] = n; return R; } }; return R; },
    appendRow(x) { rows.push(x); } };
  return sh;
}
const hdr = [['', '', '', '', '', ''], ['', '', '', '', '', ''], ['STT', 'Mã hồ sơ', 'Nội dung chi tiết', 'Người thực hiện', 'Tình trạng hồ sơ', 'Rev'], ['', '', '', '', '', ''], ['I', 'KẾT CẤU', '', '', '', '']];
const SD = mkSheet('SD', hdr.map(r => r.slice()).concat([[1, '261117-COA.NP-NT-SDS-005', 'Dầm móng RE4', 'Nho', '', ''], [2, '261113-COA.HTKT-NT-SDS-005', 'Mái che 2', 'Nho', '', '']]));
const RFI = mkSheet('RFI', hdr.map(r => r.slice()).concat([[3, '261116-COA-TBC-RFI-003 rev01', 'Nền công viên', '', '', '01']]));
const LOG = mkSheet('TQN_THA_FILE', [['h']]);
const sheets = { SD, RFI, TQN_THA_FILE: LOG, TQN_CONFIG: mkSheet('TQN_CONFIG', [['Khóa', 'Giá trị']]) };
const mkFile = n => { const f = { n, folder: 'vao', trashed: false, getName: () => n, getUrl: () => 'https://drive/' + encodeURIComponent(n), moveTo(d) { f.folder = d.name; }, setTrashed() { f.trashed = true; } }; return f; };
const files = ['261117-COA.NP-NT-SDS-005.pdf', '261113_COA.HTKT_NT_SDS_005 rev00.pdf', '261120-COA-TBC-RFI-003 rev02.pdf', 'Ban ve tong the.pdf', '261119-COA.NP-NT-SDS-099.pdf'].map(mkFile);
const mkFolder = (name, sub) => ({ name, sub: sub || {}, getId: () => name, getUrl: () => 'https://drive/f/' + name, isTrashed: () => false,
  getFoldersByName(n) { const s = this.sub; let used = false; return { hasNext: () => !used && !!s[n], next: () => { used = true; return s[n]; } }; },
  createFolder(n) { this.sub[n] = mkFolder(n); return this.sub[n]; },
  getFiles() { const l = name === '0. THẢ FILE VÀO ĐÂY' ? files.filter(f => f.folder === 'vao') : []; let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; } });
const vaoF = mkFolder('0. THẢ FILE VÀO ĐÂY'); vaoF.name = 'vao';
vaoF.getFiles = () => { const l = files.filter(f => f.folder === 'vao'); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; };
const goc = mkFolder('TQN Register – Hồ sơ', { '0. THẢ FILE VÀO ĐÂY': vaoF });
let props = {}, toast = '';
const ss = { getSheetByName: n => sheets[n] || null, toast: t => { toast = t; } };
const ctx = { console, Math, JSON, String, Date, parseInt, RegExp, Object,
  Utilities: { formatDate: () => '16:30 19/11/2026' }, Session: { getActiveUser: () => ({ getEmail: () => 'nho' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  DriveApp: { getFolderById: () => goc, getFoldersByName: () => ({ hasNext: () => false }), createFolder: () => goc },
  SpreadsheetApp: { newRichTextValue: () => { const o = { t: '', u: '' }; o.setText = x => { o.t = x; return o; }; o.setLinkUrl = u => { o.u = u; return o; }; o.build = () => o; return o; } },
  ScriptApp: { getProjectTriggers: () => [] } };
vm.createContext(ctx);
vm.runInContext(['Config', 'Read', 'Workflow', 'TuDongMa', 'ThaFile'].map(n => fs.readFileSync(D(n + '.gs'), 'utf8')).join('\n'), ctx);
Object.assign(ctx, { tqnSs_: () => ss, tqnAudit_() {}, tqnCacheClear_() {}, tqnOwnSheet_: () => LOG });
let p = 0, f = 0; const T = (t, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + t + (i ? ' | ' + i : '')); ok ? p++ : f++; };
const ma = n => vm.runInContext('tqnMaTuTenFile_(' + JSON.stringify(n) + ')', ctx);
T('F01 Đọc mã từ tên file có ".pdf"', ma('261117-COA.NP-NT-SDS-005.pdf').code === '261117-COA.NP-NT-SDS-005');
T('F02 Tên dùng gạch dưới "_" vẫn đọc được', ma('261113_COA.HTKT_NT_SDS_005 rev00.pdf') && ma('261113_COA.HTKT_NT_SDS_005 rev00.pdf').key === 'COA.HTKT-NT-SDS-005', JSON.stringify(ma('261113_COA.HTKT_NT_SDS_005 rev00.pdf')));
T('F03 Tên không có mã → null', ma('Ban ve tong the.pdf') === null);
const kq = vm.runInContext('tqnPhanPhoiFile()', ctx);
T('F04 SDS-005 gói NP → SD dòng 6 (không lẫn HTKT-005)', /SDS-005\.pdf: ĐÃ PHÂN PHỐI → SD dòng 6/.test(kq.join('\n')), kq[0]);
T('F05 HTKT-005 → SD dòng 7', /HTKT_NT_SDS_005 rev00\.pdf: ĐÃ PHÂN PHỐI → SD dòng 7/.test(kq.join('\n')));
T('F06 RFI-003 rev02 ngày khác vẫn khớp dòng RFI-003', /RFI-003 rev02\.pdf: ĐÃ PHÂN PHỐI → RFI dòng 6/.test(kq.join('\n')));
T('F07 Ô mã gắn link file, chữ giữ nguyên', SD.rich['6:2'] && SD.rich['6:2'].t === '261117-COA.NP-NT-SDS-005' && /SDS-005\.pdf/.test(decodeURIComponent(SD.rich['6:2'].u)));
T('F08 File chuyển vào thư mục SD / RFI', files[0].folder === 'SD' && files[2].folder === 'RFI', files.map(x => x.folder).join(','));
T('F09 Không có mã / mã chưa có trong sổ → "Chưa nhận ra mã"', files[3].folder === 'Chưa nhận ra mã' && files[4].folder === 'Chưa nhận ra mã');
T('F10 Không xóa file nào', files.every(x => !x.trashed));
T('F11 Nhật ký TQN_THA_FILE đủ 5 dòng', LOG.rows.length === 6, LOG.rows.length - 1);
T('F12 Chạy lại khi thư mục trống → không làm gì', vm.runInContext('tqnPhanPhoiFile()', ctx).length === 0);
console.log(`KẾT QUẢ THẢ FILE: ${p} PASS / ${f} FAIL`);
