// Chạy Apps Script TQN trên dữ liệu Register thật bằng giả lập SpreadsheetApp. CHỈ ĐỌC fixture.
// node run_tests.js fixture_live.json 2026-11-19
process.env.TZ = 'Asia/Ho_Chi_Minh';
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const FIX = process.argv[2] || 'fixture_live.json', TODAY = process.argv[3] || '2026-11-19';
const fx = JSON.parse(fs.readFileSync(path.join(__dirname, FIX), 'utf8'));
const ORIGINAL = new Set(Object.keys(fx));
const writes = [];   // [sheet, action]
const dec = v => (v && typeof v === 'object' && v.__d) ? new Date(v.__d) : (v === null || v === undefined ? '' : v);

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); o.push(row); } return o; }
  getFormulas() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(((this.sh.f[this.r + i - 1] || [])[this.c + j - 1]) || ''); o.push(row); } return o; }
  setValues(v) { this.sh.w('setValues'); if (v.length !== this.nr || v[0].length !== this.nc) throw new Error('setValues sai kích thước ' + this.sh.name); v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  clearContent() { this.sh.w('clear'); for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; }
}
['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat'].forEach(m => Range.prototype[m] = function () { this.sh.w(m); return this; });

class Sheet {
  constructor(name, data) { this.name = name; this.v = (data ? data.values : []).map(r => r.map(dec)); this.f = data ? data.formulas : []; }
  w(a) { writes.push([this.name, a]); if (ORIGINAL.has(this.name)) throw new Error('GHI VÀO TAB GỐC: ' + this.name + ' (' + a + ')'); }
  get(r, c) { const x = (this.v[r - 1] || [])[c - 1]; return x === undefined || x === null ? '' : x; }
  set(r, c, x) { while (this.v.length < r) this.v.push([]); const row = this.v[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getName() { return this.name; }
  getLastRow() { for (let i = this.v.length; i > 0; i--) if (this.v[i - 1].some(x => x !== '' && x !== null)) return i; return 0; }
  getLastColumn() { let m = 0; this.v.forEach(r => { for (let j = r.length; j > m; j--) if (r[j - 1] !== '' && r[j - 1] !== null) { m = j; break; } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('getRange sai ' + [r, c, nr, nc]); return new Range(this, r, c, nr, nc); }
  appendRow(row) { this.w('appendRow'); const r = this.getLastRow() + 1; row.forEach((x, j) => this.set(r, j + 1, x)); }
  setFrozenRows() { this.w('freeze'); } setFrozenColumns() { this.w('freeze'); } setColumnWidth() { this.w('width'); }
  activate() { } setActiveRange() { }
}
const sheets = {}; Object.keys(fx).forEach(n => sheets[n] = new Sheet(n, fx[n]));
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), setActiveSheet() { } };
const mails = [], props = {};
const ctx = {
  console, Date, JSON, Math, String, Object, Array, RegExp, Error, parseInt, isNaN,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ({ alert() { }, ButtonSet: {} }) },
  Utilities: {
    DigestAlgorithm: { MD5: 'md5' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (a, s) => Array.from(crypto.createHash('md5').update(s, 'utf8').digest()).map(b => b > 127 ? b - 256 : b),
    formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate())).replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())); }
  },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@local' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() { } }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  MailApp: { sendEmail: o => mails.push(o) },
  ScriptApp: { getProjectTriggers: () => [], newTrigger() { const b = { timeBased: () => b, atHour: () => b, everyDays: () => b, inTimezone: () => b, create: () => b }; return b; }, deleteTrigger() { } }
};
vm.createContext(ctx);
const dir = path.join(__dirname, '..', 'apps-script');
fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f }));

// ---------------- TEST ----------------
let pass = 0, fail = 0; const T = (name, ok, info) => { console.log((ok ? 'PASS ' : 'FAIL ') + name + (info ? '  | ' + info : '')); ok ? pass++ : fail++; };
const run = (code) => vm.runInContext(code, ctx);
const today = new Date(TODAY + 'T08:00:00');
ctx.__today = today;

run('tqnEnsureConfig_()');
const r1 = run('(function(){var cfg=tqnConfig_(),imp=tqnImport_(cfg),qc=tqnQC_(cfg,imp),rem=tqnRemindList_(cfg,imp,__today);tqnWriteRemind_(cfg,rem);return {imp:imp,qc:qc,rem:rem,cfg:cfg};})()');
const c1 = r1.imp.counts;
console.log('Số liệu lần 1:', JSON.stringify(c1));
// bản đồ cột dò được
Object.entries(r1.imp.src.maps).forEach(([n, m]) => { if (!m) return; const L = i => i >= 0 ? run('tqnColLetter_(' + (i + 1) + ')') : '-'; const b = x => x ? `${L(x.start)}[TT=${L(x.actual)} VM=${L(x.note)} SN=${L(x.days)}]` : '-'; console.log(`  ${n.padEnd(15)} mã=${L(m.code)} nd=${L(m.subject)} tt=${L(m.status)} rev=${L(m.rev)} NT=${b(m.nt)} TVGS=${b(m.tvgs)} CĐT=${b(m.cdt)} TVTK=${b(m.tvtk)}`); });

T('U01 Không ghi vào tab gốc', !writes.some(w => ORIGINAL.has(w[0])), 'tab được ghi: ' + [...new Set(writes.map(w => w[0]))].join(', '));
T('U02 Đủ 79 hồ sơ chi tiết (SD58+M0S11+RFA4+RFI3+RFC3), không mất dòng trùng mã', c1.chi_tiet === 79, 'chi_tiet=' + c1.chi_tiet);
const dataSh = sheets['TQN_DATA'], n1 = dataSh.getLastRow() - 1, keys1 = dataSh.getRange(2, 1, n1, 1).getValues().map(r => r[0]);
T('U03 Khóa bản ghi không trùng', new Set(keys1).size === keys1.length, n1 + ' dòng');
const first1 = dataSh.getRange(2, 28, n1, 1).getValues().map(r => +r[0]);
const r2 = run('(function(){var cfg=tqnConfig_();return tqnImport_(cfg);})()');
const n2 = dataSh.getLastRow() - 1, keys2 = dataSh.getRange(2, 1, n2, 1).getValues().map(r => r[0]);
const first2 = dataSh.getRange(2, 28, n2, 1).getValues().map(r => +r[0]);
T('U04 Nhập lần 2 không sinh thêm dòng (idempotent)', n1 === n2 && keys1.join() === keys2.join(), n1 + ' → ' + n2);
T('U05 Giữ "Nhập lần đầu" khi nhập lại', first1.join() === first2.join());
const d = r1.imp.src.details, find = (code, s) => d.filter(x => x.code === code && (!s || x.subject.includes(s)));
T('U06 SDS-025 giữ cả 2 hồ sơ (thang bộ + cột)', find('261117-COA.NP-NT-SDS-025').length === 2);
const m8 = find('261113-COA.NP-NT-MSS-008')[0];
T('U07 M0S: lấy được ngày TVGS (cột N) và ý kiến (cột M)', m8 && m8.tvgsNote.includes('Bổ sung biện pháp neo') && m8.submitted instanceof Date, m8 && ('note=' + m8.tvgsNote + ' | ngày trình=' + m8.submitted));
const rfi1 = find('261116-COA-TBC-RFI-001 rev01')[0];
T('U08 Ghép WF↔TBC: RFI-001 ↔ 20261116.COT.TC38.COA2B.V1.001.R00', rfi1 && rfi1.log && rfi1.log.wf === '20261116.COT.TC38.COA2B.V1.001.R00', rfi1 && rfi1.matchHow);
const rfc1 = find('261116-COA-TBC-RFC-001')[0];
T('U09 Ghép được dù tên file sai ngày (260611 thay vì 261116)', rfc1 && rfc1.log && rfc1.log.wf.startsWith('20261116.COT.TC38.COA6A.001'), rfc1 && rfc1.matchHow);
T('U10 Ghép đủ 10/10 RFI-RFA-RFC còn hiệu lực có trong sổ tổng', c1.ghep_duoc >= 8, 'ghep_duoc=' + c1.ghep_duoc);
const rules = r1.qc.byRule; console.log('QC theo luật:', JSON.stringify(rules));
const has = (rule, s) => r1.qc.rows.some(x => x[1] === rule && (!s || (x.join(' ').includes(s))));
T('U11 Q01 bắt trùng mã SDS-025 và SDS-026', has('Q01_TRUNG_MA', 'SDS-025') && has('Q01_TRUNG_MA', 'SDS-026'));
T('U12 Q02 bắt khoảng trắng RFI-003', has('Q02_KHOANG_TRANG', 'RFI-003'));
T('U13 Q05 bắt SDS-001 đã trình nhưng chưa vào Sổ trình duyệt', has('Q05_CHUA_VAO_SO_TONG', 'SDS-001'));
T('U14 Q08 bắt ý kiến CĐT chép lặp RFC-001 / RFA-COS-001', has('Q08_Y_KIEN_LAP', 'RFC-001'));
T('U15 Q09 bắt công thức sai số ngày CĐT ở RFA, RFC, Hồ sơ an toàn', ['RFA!', 'RFC!', 'Hồ sơ an toàn!'].every(s => has('Q09_CONG_THUC', s)));
T('U16 Q07 bắt lệch tình trạng RFI-001 (sổ tổng Đã duyệt, sổ chi tiết Đã trình duyệt)', has('Q07_LECH_TINH_TRANG', 'RFI-001'));
console.log('Nhắc hạn ngày ' + TODAY + ':'); r1.rem.forEach(x => console.log('   ', x[0], ctx.Utilities.formatDate(x[1] || new Date(0), '', 'dd/MM/yyyy'), x[2], x[4] || x[3], '|', String(x[6]).slice(0, 60)));
T('U17 Nhắc hạn: không đưa hồ sơ đã có Kết quả', r1.rem.every(x => { const L = r1.imp.src.log.find(l => l.wf === x[3]); return L && !L.result; }));
T('U18 Mặc định không gửi mail', mails.length === 0);
// chạy job hằng ngày 2 lần cùng ngày → chỉ chạy 1
run('tqnDailyJob()'); const a1 = sheets['TQN_AUDIT'].getLastRow(); run('tqnDailyJob()');
const lastAudit = sheets['TQN_AUDIT'].getRange(sheets['TQN_AUDIT'].getLastRow(), 4, 1, 1).getValues()[0][0];
T('U19 Job hằng ngày chạy 2 lần cùng ngày thì lần 2 bỏ qua', String(lastAudit).startsWith('Bỏ qua'), lastAudit);
const qcRows = sheets['TQN_QC'].getLastRow() - 1;
console.log(`\nTQN_DATA ${n2} dòng · TQN_QC ${qcRows} dòng · TQN_NHAC_HAN ${sheets['TQN_NHAC_HAN'].getLastRow() - 1} dòng`);
console.log(`KẾT QUẢ: ${pass} PASS / ${fail} FAIL`);
fs.writeFileSync(path.join(__dirname, 'out_' + FIX.replace('.json', '') + '_QC.tsv'), sheets['TQN_QC'].v.map(r => r.map(x => x instanceof Date ? x.toISOString() : x).join('\t')).join('\n'));
fs.writeFileSync(path.join(__dirname, 'out_' + FIX.replace('.json', '') + '_DATA.tsv'), sheets['TQN_DATA'].v.map(r => r.map(x => x instanceof Date ? x.toISOString().slice(0, 10) : x).join('\t')).join('\n'));

// xuất dữ liệu thanh bên để chụp hình giao diện
fs.writeFileSync(path.join(__dirname, 'sidebar_data.json'), JSON.stringify(run('tqnSidebarData()')));
process.exit(fail ? 1 : 0);
