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
  getTabColor() { return this.tab || null; } setTabColor(c) { this.tab = c; return this; }
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


const run=c=>vm.runInContext(c,ctx); ctx.__today=new Date('2026-11-19T08:00:00');
run('tqnEnsureConfig_()'); run('tqnImport_(tqnConfig_())');
const at=['RFA!9','RFI!7','M0S!'+run('(function(){var d=tqnReadAll_(tqnConfig_()).details.filter(function(x){return x.sheet==="M0S"&&x.tvgsNote;})[0];return d.row;})()')];
const P={};at.forEach(a=>{P[a]=run('tqnAiPrompt_(tqnAiRecord_('+JSON.stringify(a)+'))');});
fs.writeFileSync(path.join(__dirname,'ai_prompts.json'),JSON.stringify(P,null,1));
at.forEach(a=>console.log('=== '+a+'\n'+P[a].split('--- DỮ LIỆU HỒ SƠ ---')[1]));
const A=JSON.parse(fs.readFileSync(path.join(__dirname,'ai_answers.json'),'utf8'));
let ok=0;
Object.keys(A).forEach(a=>{ctx.__a=A[a];ctx.__at=a;
 try{const o=run('(function(){var rec=tqnAiRecord_(__at),o=tqnAiParse_(__a);tqnAiSave_(rec,o,"claude (thử)");return o;})()');ok++;
  console.log('PASS đọc kết quả AI '+a+': mức '+o.muc_do+' | '+o.ben_xu_ly+' | '+o.viec_tiep_theo);}catch(e){console.log('FAIL '+a+' '+e.message)}});
try{run('tqnAiParse_("Xin lỗi, tôi không làm được")');console.log('FAIL câu trả lời rác vẫn nhận')}catch(e){console.log('PASS câu trả lời không phải JSON bị chặn: '+e.message)}
console.log('TQN_AI có '+(sheets['TQN_AI'].getLastRow()-1)+' dòng; ghi tab gốc: '+writes.filter(w=>ORIGINAL.has(w[0])).length);
