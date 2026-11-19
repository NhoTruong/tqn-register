// Chạy Apps Script TQN trên dữ liệu Register thật bằng giả lập SpreadsheetApp. CHỈ ĐỌC fixture.
// node run_tests.js fixture_live.json 2026-11-19
process.env.TZ = 'Asia/Ho_Chi_Minh';
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const FIX = process.argv[2] || 'fixture_live.json', TODAY = process.argv[3] || '2026-11-19';
const fx = JSON.parse(fs.readFileSync(path.join(__dirname, FIX), 'utf8'));
const STY = (() => { try { return JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'preview', 'site', 'data.json'), 'utf8')).styles || {}; } catch (e) { return {}; } })();
const ORIGINAL = new Set(Object.keys(fx));
const writes = [], ALLOW = new Set();   // [sheet, action]
const dec = v => (v && typeof v === 'object' && v.__d) ? new Date(v.__d) : (v === null || v === undefined ? '' : v);

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); o.push(row); } return o; }
  getFormulas() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(((this.sh.f[this.r + i - 1] || [])[this.c + j - 1]) || ''); o.push(row); } return o; }
  setValues(v) { this.sh.w('setValues'); if (v.length !== this.nr || v[0].length !== this.nc) throw new Error('setValues sai kích thước ' + this.sh.name); v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  _st(f, d) { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(f(this.sh.sty[(this.r + i) + ',' + (this.c + j)]) || d); o.push(row); } return o; }
  _set(k, v) { v.forEach((row, i) => row.forEach((x, j) => { const key = (this.r + i) + ',' + (this.c + j); const s = this.sh.sty[key] = this.sh.sty[key] || ['', '', 0, '']; if (k === 2) s[2] = x === 'bold' ? 1 : 0; else s[k] = x; })); return this; }
  getBackgrounds() { return this._st(s => s && s[0], '#ffffff'); } getFontColors() { return this._st(s => s && s[1], '#000000'); } getFontWeights() { return this._st(s => s && (s[2] ? 'bold' : 'normal'), 'normal'); }
  setBackgrounds(v) { return this._set(0, v); } setFontColors(v) { return this._set(1, v); } setFontWeights(v) { return this._set(2, v); }
  setValue(x) { this.sh.w('setValue'); this.sh.set(this.r, this.c, x); return this; }
  getValue() { return this.sh.get(this.r, this.c); }
  clearContent() { this.sh.w('clear'); for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; }
}
['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat'].forEach(m => Range.prototype[m] = function () { this.sh.w(m); return this; });

class Sheet {
  constructor(name, data) { this.name = name; this.sty = JSON.parse(JSON.stringify((STY[name] || {}).s || {})); this.v = (data ? data.values : []).map(r => r.map(dec)); this.f = data ? data.formulas : []; }
  w(a) { writes.push([this.name, a]); if (ORIGINAL.has(this.name) && !ALLOW.has(this.name)) throw new Error('GHI VÀO TAB GỐC: ' + this.name + ' (' + a + ')'); }
  get(r, c) { const x = (this.v[r - 1] || [])[c - 1]; return x === undefined || x === null ? '' : x; }
  set(r, c, x) { while (this.v.length < r) this.v.push([]); const row = this.v[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getName() { return this.name; }
  getLastRow() { for (let i = this.v.length; i > 0; i--) if (this.v[i - 1].some(x => x !== '' && x !== null)) return i; return 0; }
  getLastColumn() { let m = 0; this.v.forEach(r => { for (let j = r.length; j > m; j--) if (r[j - 1] !== '' && r[j - 1] !== null) { m = j; break; } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('getRange sai ' + [r, c, nr, nc]); return new Range(this, r, c, nr, nc); }
  appendRow(row) { this.w('appendRow'); const r = this.getLastRow() + 1; row.forEach((x, j) => this.set(r, j + 1, x)); }
  setFrozenRows() { this.w('freeze'); } setFrozenColumns() { this.w('freeze'); } setColumnWidth() { this.w('width'); }
  activate() { } setActiveRange() { } hideSheet() { this.hidden = true; }
  getMaxRows() { return Math.max(this.v.length, 1000); }
  getConditionalFormatRules() { return (this.cfr || []).slice(); }
  setConditionalFormatRules(r) { this.cfr = r; }
}
const sheets = {}; Object.keys(fx).forEach(n => sheets[n] = new Sheet(n, fx[n]));
const ss = { getSheets: () => Object.values(sheets), getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), setActiveSheet() { } };
const mails = [], props = {}, TRIG = [], sprops = {}, fetches = []; let AI_RES = () => ({ getResponseCode: () => 500, getContentText: () => '' });
const ctx = {
  console, Date, JSON, Math, String, Object, Array, RegExp, Error, parseInt, isNaN,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, newConditionalFormatRule() { const o = {}; const b = { whenTextEqualTo: v => (o.v = v, b), setBackground: c => (o.bg = c, b), setFontColor: c => (o.fc = c, b), setRanges: r => (o.r = r, b), build: () => ({ o, getBooleanCondition: () => ({ getCriteriaValues: () => [o.v], getBackground: () => o.bg }) }) }; return b; }, getUi: () => ({ alert() { }, ButtonSet: {} }) },
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


// ===== AUDIT =====

const run=c=>vm.runInContext(c,ctx); ctx.__today=new Date('2026-11-19T08:00:00');
run('tqnEnsureConfig_()');
const out=[];const P=s=>{out.push(s);console.log(s)};
const F=d=>d instanceof Date?ctx.Utilities.formatDate(d,'','dd/MM/yyyy'):(d||'');
const imp=run('(function(){var cfg=tqnConfig_();return tqnImport_(cfg);})()');
const det=imp.src.details, log=imp.src.log;
P('# KIỂM TỪNG TAB – '+new Date().toISOString());
P('\n## 1. Sổ chi tiết đọc được');
const by={};det.forEach(d=>{const b=by[d.sheet]=by[d.sheet]||{n:0,sub:0,tv:0,cd:0,tk:0,st:0,ghep:0};b.n++;if(d.submitted)b.sub++;if(d.tvgsDate||d.tvgsNote)b.tv++;if(d.cdtDate||d.cdtNote)b.cd++;if(d.tvtkDate||d.tvtkNote)b.tk++;if(d.status)b.st++;if(d.log)b.ghep++;});
P('Tab | Dòng | Có ngày trình | Có phản hồi TVGS | CĐT | TVTK | Có tình trạng | Ghép Sổ trình duyệt');
Object.entries(by).forEach(([k,b])=>P(`${k} | ${b.n} | ${b.sub} | ${b.tv} | ${b.cd} | ${b.tk} | ${b.st} | ${b.ghep}`));
P('Sổ trình duyệt: '+log.length+' dòng; có Kết quả: '+log.filter(l=>l.result).length+'; có Ngày phản hồi: '+log.filter(l=>l.reply).length);
P('\n## 2. RFI / RFA / RFC – từng dòng');
det.filter(d=>/^RF/.test(d.sheet)).forEach(d=>P(`${d.sheet}!${d.row} ${d.code} rev${d.rev} | TT: ${d.status||'-'} | trình ${F(d.submitted)||'-'} | TVGS ${F(d.tvgsDate)||'-'} ${d.tvgsNote?'“'+d.tvgsNote.slice(0,50)+'”':''} | CĐT ${F(d.cdtDate)||'-'} ${d.cdtNote?'“'+d.cdtNote.slice(0,50)+'”':''} | Sổ tổng: ${d.log?(d.log.wf+' KQ='+(d.log.result||'-')+' TT='+(d.log.status||'-')):'(chưa ghép)'}`));
const wf=run('(function(){var cfg=tqnConfig_();return tqnBuildWorkflow_(cfg,tqnImport_(cfg),__today);})()');
P('\n## 3. Bước quy trình hiện tại: '+JSON.stringify(wf.count)+' | quá hạn '+wf.late);
const stepOf=code=>{const r=wf.rows.find(x=>String(x[3]).replace(/\s/g,'').includes(code.replace(/\s/g,'')));return r?r[0]+' / '+r[1]:'(không thấy)';};
// 4. Mô phỏng cập nhật phản hồi
P('\n## 4. Thử sửa phản hồi rồi chạy lại (trên bản sao trong bộ nhớ, không đụng file)');
const pick=det.find(d=>d.sheet==='RFI'&&!d.tvgsDate&&!d.cdtDate)||det.find(d=>d.sheet==='RFI');
const m=imp.src.maps['RFI'];
const sh=sheets['RFI'];
const tests=[];
function step(code){const w=run('(function(){var cfg=tqnConfig_();return tqnBuildWorkflow_(cfg,tqnImport_(cfg),__today);})()');const r=w.rows.find(x=>String(x[3]).replace(/\s/g,'')===code.replace(/\s/g,''));return r?r[0]:'?';}
const code=pick.rawCode||pick.code;
P('Hồ sơ thử: RFI!'+pick.row+' '+code+' – trước: '+step(code));
ALLOW.add('RFI');
sh.set(pick.row,m.tvgs.actual+1,new Date('2026-11-18T00:00:00'));
const s1=step(code);P('  + gõ ngày TVGS 18/11 → '+s1);tests.push(['Có ngày TVGS → bước 3 (chờ CĐT)',s1.startsWith('3')||s1.startsWith('5')]);
sh.set(pick.row,m.status+1,'Không duyệt');
const s2=step(code);P('  + Tình trạng "Không duyệt" → '+s2);tests.push(['Không duyệt → bước 4 (trình lại)',s2.startsWith('4')]);
sh.set(pick.row,m.status+1,'Đã duyệt');
const s3=step(code);P('  + Tình trạng "Đã duyệt" → '+s3);tests.push(['Đã duyệt → bước 5',s3.startsWith('5')]);
// RFA: sổ tổng ghi Kết quả
const lg=log.find(l=>!l.result);
if(lg){const shL=sheets['Sổ trình duyệt'];ALLOW.add('Sổ trình duyệt');const H=shL.v.findIndex(r=>r.some(x=>String(x).trim()==='Kết quả'));const cK=shL.v[H].findIndex(x=>String(x).trim()==='Kết quả');
 const before=run('(function(){var cfg=tqnConfig_();var w=tqnBuildWorkflow_(cfg,tqnImport_(cfg),__today);var r=w.rows.filter(function(x){return x[14]==='+JSON.stringify(lg.wf)+'})[0];return r?r[0]:"?";})()');
 shL.set(lg.row,cK+1,'Duyệt có ý kiến');
 const after=run('(function(){var cfg=tqnConfig_();var w=tqnBuildWorkflow_(cfg,tqnImport_(cfg),__today);var r=w.rows.filter(function(x){return x[14]==='+JSON.stringify(lg.wf)+'})[0];return r?r[0]:"?";})()');
 P('Sổ trình duyệt dòng '+lg.row+' ('+lg.wf+'): '+before+' → ghi Kết quả "Duyệt có ý kiến" → '+after);tests.push(['Sổ tổng có Kết quả → bước 5',after.startsWith('5')]);}
P('\n## 5. Các tab có dữ liệu nhưng tool CHƯA đọc');
const read=new Set(Object.keys(by).concat(['Sổ trình duyệt']));
Object.keys(fx).filter(n=>!read.has(n)&&!/^TQN_/.test(n)).forEach(n=>{const s=sheets[n];P(`${n}: ${s.getLastRow()} dòng`)});
P('\n## KẾT QUẢ');tests.forEach(t=>P((t[1]?'PASS ':'FAIL ')+t[0]));
fs.writeFileSync(path.join(__dirname,'out_audit_tabs.md'),out.join('\n'));
