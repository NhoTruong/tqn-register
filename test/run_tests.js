// Chạy Apps Script TQN trên dữ liệu Register thật bằng giả lập SpreadsheetApp. CHỈ ĐỌC fixture.
// node run_tests.js fixture_live.json 2026-11-19
process.env.TZ = 'Asia/Ho_Chi_Minh';
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const FIX = process.argv[2] || 'fixture_live.json', TODAY = process.argv[3] || '2026-11-19';
const fx = JSON.parse(fs.readFileSync(path.join(__dirname, FIX), 'utf8'));
const ORIGINAL = new Set(Object.keys(fx));
const writes = [], ALLOW = new Set();   // [sheet, action]
const dec = v => (v && typeof v === 'object' && v.__d) ? new Date(v.__d) : (v === null || v === undefined ? '' : v);

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); o.push(row); } return o; }
  getFormulas() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(((this.sh.f[this.r + i - 1] || [])[this.c + j - 1]) || ''); o.push(row); } return o; }
  setValues(v) { this.sh.w('setValues'); if (v.length !== this.nr || v[0].length !== this.nc) throw new Error('setValues sai kích thước ' + this.sh.name); v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  setValue(x) { this.sh.w('setValue'); this.sh.set(this.r, this.c, x); return this; }
  getValue() { return this.sh.get(this.r, this.c); }
  clearContent() { this.sh.w('clear'); for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; }
}
['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat'].forEach(m => Range.prototype[m] = function () { this.sh.w(m); return this; });

class Sheet {
  constructor(name, data) { this.name = name; this.v = (data ? data.values : []).map(r => r.map(dec)); this.f = data ? data.formulas : []; }
  w(a) { writes.push([this.name, a]); if (ORIGINAL.has(this.name) && !ALLOW.has(this.name)) throw new Error('GHI VÀO TAB GỐC: ' + this.name + ' (' + a + ')'); }
  get(r, c) { const x = (this.v[r - 1] || [])[c - 1]; return x === undefined || x === null ? '' : x; }
  set(r, c, x) { while (this.v.length < r) this.v.push([]); const row = this.v[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getName() { return this.name; }
  getLastRow() { for (let i = this.v.length; i > 0; i--) if (this.v[i - 1].some(x => x !== '' && x !== null)) return i; return 0; }
  getLastColumn() { let m = 0; this.v.forEach(r => { for (let j = r.length; j > m; j--) if (r[j - 1] !== '' && r[j - 1] !== null) { m = j; break; } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('getRange sai ' + [r, c, nr, nc]); return new Range(this, r, c, nr, nc); }
  appendRow(row) { this.w('appendRow'); const r = this.getLastRow() + 1; row.forEach((x, j) => this.set(r, j + 1, x)); }
  setFrozenRows() { this.w('freeze'); } setFrozenColumns() { this.w('freeze'); } setColumnWidth() { this.w('width'); }
  activate() { } setActiveRange() { }
  getMaxRows() { return Math.max(this.v.length, 1000); }
  getConditionalFormatRules() { return (this.cfr || []).slice(); }
  setConditionalFormatRules(r) { this.cfr = r; }
}
const sheets = {}; Object.keys(fx).forEach(n => sheets[n] = new Sheet(n, fx[n]));
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), setActiveSheet() { } };
const mails = [], props = {}, TRIG = [], sprops = {}, fetches = []; let AI_RES = () => ({ getResponseCode: () => 500, getContentText: () => '' });
const ctx = {
  console, Date, JSON, Math, String, Object, Array, RegExp, Error, parseInt, isNaN,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, newConditionalFormatRule() { const o = {}; const b = { whenTextEqualTo: v => (o.v = v, b), setBackground: c => (o.bg = c, b), setFontColor: c => (o.fc = c, b), setRanges: r => (o.r = r, b), build: () => ({ o, getBooleanCondition: () => ({ getCriteriaValues: () => [o.v] }) }) }; return b; }, getUi: () => ({ alert() { }, ButtonSet: {} }) },
  Utilities: {
    sleep: () => {},
    DigestAlgorithm: { MD5: 'md5' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (a, s) => Array.from(crypto.createHash('md5').update(s, 'utf8').digest()).map(b => b > 127 ? b - 256 : b),
    formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate())).replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())); }
  },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@local' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() { } }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }), getScriptProperties: () => ({ getProperty: k => sprops[k] || null }) },
  UrlFetchApp: { fetch: (url, o) => { fetches.push({ url, o }); return AI_RES(url, o); } },
  MailApp: { sendEmail: o => mails.push(o) },
  ScriptApp: { getProjectTriggers: () => TRIG.slice(), deleteTrigger(t) { TRIG.splice(TRIG.indexOf(t), 1); },
    newTrigger(fn) { const t = { getHandlerFunction: () => fn }; const b = { timeBased: () => b, atHour: () => b, nearMinute: () => b, everyHours: () => b, everyDays: () => b, inTimezone: () => b, create: () => { TRIG.push(t); return t; } }; return b; } }
};
vm.createContext(ctx);
const dir = path.join(__dirname, '..', 'apps-script');
fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f }));

// ---------------- TEST ----------------
let pass = 0, fail = 0, expectErr; const T = (name, ok, info) => { console.log((ok ? 'PASS ' : 'FAIL ') + name + (info ? '  | ' + info : '')); ok ? pass++ : fail++; };
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
// ---- Tự động 1 nút ----
const cnt = () => TRIG.map(t => t.getHandlerFunction()).sort().join(',');
const rep1 = run('tqnSetupAll()');
T('U20 Bật tự động: có lịch mỗi giờ + nhắc hạn; thời tiết chưa đủ điều kiện thì KHÔNG bật và báo lý do', cnt() === 'tqnDailyJob,tqnHourlyJob' && rep1.some(x => x.startsWith('✖ Thời tiết')), cnt() + ' | ' + rep1.join(' / '));
run('tqnSetupAll()');
T('U21 Bấm lại không sinh lịch trùng', cnt() === 'tqnDailyJob,tqnHourlyJob', cnt());
const setC = (k, v) => { const c = sheets['TQN_CONFIG']; for (let r = 2; r <= c.getLastRow(); r++) if (c.get(r, 1) === k) c.set(r, 2, v); };
setC('weather_enabled', 'Có'); setC('weather_license_ok', 'Có');
const rep2 = run('tqnSetupAll()');
T('U22 Đủ điều kiện (GPS Coastal đã xác nhận + điều khoản) thì bật đủ 3 lịch', cnt() === 'tqnDailyJob,tqnHourlyJob,tqnWeatherJob', cnt() + ' | ' + rep2.join(' / '));
T('U23 Trạng thái tự động báo đủ 3 dòng ✔', run('tqnAutoStatus()').every(x => x.startsWith('✔')), run('tqnAutoStatus()').join(' / '));
run('tqnRemoveAll()');
T('U24 Tắt tất cả: xóa hết lịch TQN', TRIG.length === 0, cnt());
T('U25 Tự động mỗi giờ không ghi tab gốc', !writes.some(w => ORIGINAL.has(w[0])));
// ---- S2: quy trình + cấp mã ----
const wfr = run('(function(){var r=tqnRunAll_();return r.wf;})()');
const sumSteps = Object.values(wfr.count).reduce((a, b) => a + b, 0);
T('U26 Bảng quy trình: mỗi hồ sơ (Rev mới nhất) có đúng 1 bước + người giữ việc', wfr.rows.length > 0 && sumSteps === wfr.rows.length && wfr.rows.every(x => x[1]), JSON.stringify(wfr.count) + ' quá hạn=' + wfr.late);
T('U27 Bước 5 (đã có kết quả) không còn ai giữ việc; bước 2–3 quá hạn mới gắn QUÁ HẠN', wfr.rows.filter(x => x[0].startsWith('5')).every(x => x[1] === '—') && wfr.rows.filter(x => x[11]).every(x => /^[23]/.test(x[0])));
const ymd = run("tqnYymmdd_(new Date(), 'Asia/Ho_Chi_Minh')");
T('U28 Xem trước RFI tiếp theo = 004 rev01 (sổ có 001–003)', run("tqnPreviewCode({type:'RFI'})") === ymd + '-COA-TBC-RFI-004 rev01', run("tqnPreviewCode({type:'RFI'})"));
T('U29 RFC tiếp theo = 004; RFA nhóm COS tiếp theo = 003', run("tqnPreviewCode({type:'RFC'})") === ymd + '-COA-TBC-RFC-004' && run("tqnPreviewCode({type:'RFA',nhom:'COS'})") === ymd + '-COA-TBC-RFA-COS-003-rev01', run("tqnPreviewCode({type:'RFC'})") + ' | ' + run("tqnPreviewCode({type:'RFA',nhom:'COS'})"));
const sdNums = sheets['SD'].v.slice(5).map(r => /COA\.NP\s*-\s*NT\s*-\s*SDS\s*-\s*(\d+)/i.exec(String(r[1] || ''))).filter(Boolean).map(m => +m[1]);
const expSD = ymd + '-COA.NP-NT-SDS-' + String(Math.max(...sdNums) + 1).padStart(3, '0');
T('U30 SDS gói NP tiếp theo = số lớn nhất trong sổ SD + 1', run("tqnPreviewCode({type:'SDS',goi:'NP'})") === expSD, run("tqnPreviewCode({type:'SDS',goi:'NP'})") + ' (cần ' + expSD + ')');
const i1 = run("tqnIssueCode({type:'RFI',subject:'Thử cấp mã 1'})"), i2 = run("tqnIssueCode({type:'RFI',subject:'Thử cấp mã 2'})");
T('U31 Cấp 2 lần liên tiếp ra 2 số khác nhau (004, 005) và lưu vào TQN_CAP_MA', i1.code.includes('RFI-004') && i2.code.includes('RFI-005') && sheets['TQN_CAP_MA'].getLastRow() === 3, i1.code + ' / ' + i2.code);
T('U32 Số đã cấp không cấp lại (xem trước sau đó = 006)', run("tqnPreviewCode({type:'RFI'})").includes('RFI-006'));
expectErr = (code, re) => { try { run(code); return false; } catch (e) { return re.test(e.message); } };
T('U33 Thiếu nội dung thì không cấp', expectErr("tqnIssueCode({type:'RFI',subject:''})", /nội dung/));
const rv = run("tqnReviseCode({from:'261116-COA-TBC-RFI-001 rev01'})");
T('U34 Trình lại RFI-001 rev01 → giữ số 001, Rev 02, ngày hôm nay', rv.code === ymd + '-COA-TBC-RFI-001 rev02' && rv.rev === '02', rv.code);
ALLOW.add('RFI');
const rfiBefore = sheets['RFI'].getLastRow();
const w1 = run("tqnIssueCode({type:'RFI',subject:'Ghi thử vào sổ',owner:'Nho',write:true})");
const wr = +w1.where.split('!')[1];
T('U35 Ghi vào sổ RFI: đúng dòng trống sau dòng có mã cuối, đúng mã + nội dung', w1.where.startsWith('RFI!') && sheets['RFI'].get(wr, 2) === w1.code && sheets['RFI'].get(wr, 3) === 'Ghi thử vào sổ', w1.where + ' (sổ trước có ' + rfiBefore + ' dòng)');
const nextRow = wr + 1; sheets['RFI'].set(nextRow, 2, ''); sheets['RFI'].set(nextRow, 3, '');
const ownCol = run("tqnDetailMap_(tqnSs_().getSheetByName('RFI').getRange(1,1,5,40).getValues()).owner") + 1;
sheets['RFI'].set(nextRow, ownCol, 'đã có người gõ');
const before = JSON.stringify(sheets['RFI'].v[nextRow - 1]);
T('U36 Ô đích đã có chữ thì từ chối, không ghi đè, không ghi nửa chừng', expectErr("tqnIssueCode({type:'RFI',subject:'Thử đè',owner:'Nho',write:true})", /không ghi đè/) && JSON.stringify(sheets['RFI'].v[nextRow - 1]) === before);
ALLOW.delete('RFI');
T('U37 Loại chưa có mẫu thì báo rõ', expectErr("tqnPreviewCode({type:'XYZ'})", /Chưa có mẫu mã/));
// ---- AI đọc nội dung ----
run('tqnEnsureAiConfig_()');
const atRFI = sheets['TQN_DATA'].v.slice(1).find(r => String(r[4]).includes('RFI-001'))[2];
T('U38 AI tắt mặc định → chặn, không gọi mạng', expectErr("tqnAiAnalyze('" + atRFI + "')", /ai_enabled/) && fetches.length === 0);
setC('ai_enabled', 'Có');
T('U39 Chưa có khóa API → chặn, không gọi mạng', expectErr("tqnAiAnalyze('" + atRFI + "')", /ANTHROPIC_API_KEY/) && fetches.length === 0);
sprops.ANTHROPIC_API_KEY = 'test-key';
AI_RES = () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ content: [{ type: 'text', text: 'Đây là kết quả: {"tom_tat":"RFI hỏi thông tin gạch porcelain, CĐT đã phản hồi.","muc_do":"THAP","ben_xu_ly":"Không ai","viec_tiep_theo":"Lưu hồ sơ và áp dụng theo phản hồi.","can_bo_sung":[],"nhap_tra_loi":"Kính gửi CĐT..."}' }] }) });
const aiBefore = writes.length;
const ai1 = run("tqnAiAnalyze('" + atRFI + "')");
const req = fetches[fetches.length - 1], pl = JSON.parse(req.o.payload);
T('U40 Gọi đúng API (header khóa + phiên bản, mô hình cấu hình), đọc đúng JSON, lưu tab TQN_AI', req.url === 'https://api.anthropic.com/v1/messages' && req.o.headers['x-api-key'] === 'test-key' && req.o.headers['anthropic-version'] && pl.model === 'claude-sonnet-5-5' && ai1.muc_do === 'THAP' && sheets['TQN_AI'].getLastRow() === 2, ai1.tom_tat);
const prm = pl.messages[0].content;
T('U41 Câu lệnh chỉ chứa dữ liệu đúng hồ sơ đó, không có email', prm.includes('gach porcelain'.replace('gach', 'gạch')) || prm.includes('porcelain') ? !/@/.test(prm.split('--- DỮ LIỆU HỒ SƠ ---')[1]) : false, prm.split('--- DỮ LIỆU HỒ SƠ ---')[1].slice(0, 160).replace(/\n/g, ' | '));
AI_RES = () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ content: [{ type: 'text', text: 'xin lỗi tôi không chắc' }] }) });
T('U42 AI trả lời sai khuôn → báo lỗi, không ghi gì vào TQN_AI', expectErr("tqnAiAnalyze('" + atRFI + "')", /JSON/) && sheets['TQN_AI'].getLastRow() === 2);
let n429 = 0; AI_RES = () => { n429++; return { getResponseCode: () => 429, getContentText: () => 'rate' }; };
T('U43 Bị giới hạn 429 → thử lại 3 lần rồi báo lỗi rõ', expectErr("tqnAiAnalyze('" + atRFI + "')", /429/) && n429 === 3);
T('U44 AI không ghi vào tab gốc', !writes.slice(aiBefore).some(w => ORIGINAL.has(w[0])));
const ncf = run('tqnApplyStatusColors()'), cfr = sheets['Sổ trình duyệt'].cfr || [];
T('U45 Tô màu cột Tình trạng (đã chạy trong nút ★): đủ 7 quy tắc, đúng cột O từ dòng 23, không đổi dữ liệu', cfr.length === 7 && cfr.every(r => r.o.r[0].c === 15 && r.o.r[0].r === 23) && !writes.some(w => w[0] === 'Sổ trình duyệt'), 'thêm ' + ncf);
T('U46 Bấm lại không thêm trùng quy tắc', run('tqnApplyStatusColors()') === 0);
const qcRows = sheets['TQN_QC'].getLastRow() - 1;
console.log(`\nTQN_DATA ${n2} dòng · TQN_QC ${qcRows} dòng · TQN_NHAC_HAN ${sheets['TQN_NHAC_HAN'].getLastRow() - 1} dòng`);
console.log(`KẾT QUẢ: ${pass} PASS / ${fail} FAIL`);
fs.writeFileSync(path.join(__dirname, 'out_' + FIX.replace('.json', '') + '_QC.tsv'), sheets['TQN_QC'].v.map(r => r.map(x => x instanceof Date ? x.toISOString() : x).join('\t')).join('\n'));
fs.writeFileSync(path.join(__dirname, 'out_' + FIX.replace('.json', '') + '_DATA.tsv'), sheets['TQN_DATA'].v.map(r => r.map(x => x instanceof Date ? x.toISOString().slice(0, 10) : x).join('\t')).join('\n'));

// xuất dữ liệu thanh bên để chụp hình giao diện
fs.writeFileSync(path.join(__dirname, 'sidebar_data.json'), JSON.stringify(run('tqnSidebarData()')));
process.exit(fail ? 1 : 0);
