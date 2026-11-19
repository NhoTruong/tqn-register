// Test Weather.gs bằng giả lập. Phản hồi API là DỮ LIỆU GIẢ LẬP đúng schema Open-Meteo, KHÔNG phải số thời tiết thật.
process.env.TZ = 'Asia/Ho_Chi_Minh';
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const fx = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture_live.json'), 'utf8'));
const ORIGINAL = new Set(Object.keys(fx)); const WEATHER_TAB = 'Theo dõi thời tiết';
const writes = []; const dec = v => (v && typeof v === 'object' && v.__d) ? new Date(v.__d) : (v === null || v === undefined ? '' : v);
class Range { constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); o.push(row); } return o; }
  setValues(v) { this.sh.w('setValues'); v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  setValue(x) { this.sh.w('setValue'); this.sh.set(this.r, this.c, x); return this; }
  setRichTextValue(rt) { this.sh.w('setRich'); this.sh.set(this.r, this.c, rt.text); this.sh.links = (this.sh.links || []).concat([rt.url]); return this; }
  clearContent() { this.sh.w('clear'); for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; } }
['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat', 'setWrap', 'setVerticalAlignment', 'setBackgrounds'].forEach(m => Range.prototype[m] = function () { this.sh.w(m); return this; });
class Sheet { constructor(name, data) { this.name = name; this.v = (data ? data.values : []).map(r => r.map(dec)); }
  w(a) { writes.push([this.name, a]); if (ORIGINAL.has(this.name) && this.name !== WEATHER_TAB) throw new Error('GHI VÀO TAB GỐC: ' + this.name); }
  get(r, c) { const x = (this.v[r - 1] || [])[c - 1]; return x === undefined || x === null ? '' : x; }
  set(r, c, x) { while (this.v.length < r) this.v.push([]); const row = this.v[r - 1]; while (row.length < c) row.push(''); row[c - 1] = x; }
  getName() { return this.name; }
  getLastRow() { for (let i = this.v.length; i > 0; i--) if (this.v[i - 1].some(x => x !== '' && x !== null)) return i; return 0; }
  getLastColumn() { let m = 0; this.v.forEach(r => { for (let j = r.length; j > m; j--) if (r[j - 1] !== '' && r[j - 1] !== null) { m = j; break; } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  appendRow(row) { this.w('appendRow'); const r = this.getLastRow() + 1; row.forEach((x, j) => this.set(r, j + 1, x)); }
  insertRowBefore(r) { this.w('insertRow'); this.v.splice(r - 1, 0, []); }
  setFrozenRows() { } setFrozenColumns() { } setColumnWidth() { } }
const sheets = {}; Object.keys(fx).forEach(n => sheets[n] = new Sheet(n, fx[n]));
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)) };
const props = {}; let apiMode = 'ok', calls = 0, lastUrl = ''; const pdfHtml = [];
const fakeBody = () => JSON.stringify({ latitude: 15.1, longitude: 108.8, timezone: 'Asia/Ho_Chi_Minh',
  daily: { time: ['2026-11-19'], weather_code: [80], temperature_2m_max: [29.4], temperature_2m_min: [23.1], precipitation_sum: [6.2], precipitation_probability_max: [80], wind_speed_10m_max: [11.3], wind_gusts_10m_max: [24.8] },
  hourly: { time: [], relative_humidity_2m: [90, 92, 88, 94] } });
const ctx = { console, Date, JSON, Math, String, Object, Array, RegExp, Error, parseInt, parseFloat, isNaN, encodeURIComponent,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ({ alert() { } }), newRichTextValue: () => { const o = { text: '', url: '' }; const b = { setText: t => (o.text = t, b), setLinkUrl: (a, z, u) => (o.url = u, b), build: () => o }; return b; } },
  DriveApp: { getFoldersByName: () => ({ hasNext: () => false }), createFolder: () => ({ getFilesByName: () => ({ hasNext: () => false }), createFile: b => ({ getUrl: () => 'https://drive/fake/' + b.name }) }) },
  Utilities: { sleep() { }, newBlob: (html, t, n) => { pdfHtml.push(html); const b = { name: n, getAs: () => b, setName: x => (b.name = x, b) }; return b; }, DigestAlgorithm: { MD5: 1 }, Charset: { UTF_8: 1 }, computeDigest: (a, s) => Array.from(crypto.createHash('md5').update(s).digest()),
    formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate())).replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())); } },
  UrlFetchApp: { fetch: (u) => { calls++; lastUrl = u; if (/start_date/.test(u)) return { getResponseCode: () => 200, getContentText: () => fs.readFileSync(path.join(__dirname, 'openmeteo_buoi_19-11.json'), 'utf8').replace(/2026-11-19/g, '2026-11-17').replace(/2026-11-20/g, '2026-11-18') }; if (apiMode === '500') return { getResponseCode: () => 500, getContentText: () => 'err' }; if (apiMode === 'bad') return { getResponseCode: () => 200, getContentText: () => '{"daily":{}}' }; return { getResponseCode: () => 200, getContentText: fakeBody }; } },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@local' }) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() { } }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() { }, newTrigger() { const b = { timeBased: () => b, atHour: () => b, nearMinute: () => b, everyDays: () => b, inTimezone: () => b, create: () => b }; return b; } } };
vm.createContext(ctx);
const dir = path.join(__dirname, '..', 'apps-script');
fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f }));
let pass = 0, fail = 0; const T = (n, ok, i) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (i ? '  | ' + i : '')); ok ? pass++ : fail++; };
const run = c => vm.runInContext(c, ctx);
const setCfg = (k, v) => { const c = sheets['TQN_CONFIG']; for (let r = 2; r <= c.getLastRow(); r++) if (c.get(r, 1) === k) { c.set(r, 2, v); return; } };
const W = sheets[WEATHER_TAB];
run('tqnEnsureConfig_(); tqnEnsureWeatherConfig_();');
let r = run('tqnWeatherJob(true)'); T('W01 Mặc định tắt → chặn, không gọi API', r.startsWith('CHAN') && calls === 0, r);
setCfg('weather_enabled', 'Có'); setCfg('weather_lat', ''); setCfg('weather_gps_confirmed', 'Không'); r = run('tqnWeatherJob(true)'); T('W02 Chưa có GPS → chặn', r.includes('tọa độ') && calls === 0, r);
setCfg('weather_lat', '15.0'); setCfg('weather_lon', '108.8'); r = run('tqnWeatherJob(true)'); T('W03 GPS chưa xác nhận → chặn', r.includes('xác nhận') && calls === 0, r);
setCfg('weather_gps_confirmed', 'Có'); r = run('tqnWeatherJob(true)'); T('W04 Chưa xác nhận điều khoản nguồn → chặn', r.includes('điều khoản') && calls === 0, r);
setCfg('weather_license_ok', 'Có'); apiMode = '500'; r = run('tqnWeatherJob(false)');
T('W05 API lỗi 500 → thử lại 3 lần, FAILED, không ghi vào Register', r.startsWith('FAILED') && calls === 3 && W.getLastRow() === 1, r + ' | gọi ' + calls + ' lần');
apiMode = 'bad'; calls = 0; r = run('tqnWeatherJob(false)'); T('W06 Phản hồi thiếu trường → FAILED, không ghi', r.startsWith('FAILED') && W.getLastRow() === 1, r);
apiMode = 'ok'; calls = 0; r = run('tqnWeatherJob(false)'); const row2 = W.v[1] || [];
T('W07 Thành công: thêm dòng 2 đúng cột Register', r === 'THEM' && W.getLastRow() === 2, r);
console.log('     Dòng ghi:', row2.map(x => x instanceof Date ? x.toLocaleDateString('vi-VN') : x).join(' | '));
T('W08 Cột đúng: A ngày, B thứ, C min, E max (ô gộp), G thời tiết, H mưa, I ẩm, J gió, K ghi chú', row2[1] === 'Thứ Năm' && row2[2] === 23.1 && row2[3] === '' && row2[4] === 29.4 && row2[6] === 'Mưa rào nhẹ' && row2[7] === 6.2 && row2[8] === 91 && row2[9] === 11.3 && /^📷 /.test(String(row2[10])));
T('W09 Ghi chú có khả năng mưa, không còn dòng "Tự động…" (anh Nho yêu cầu 15:04); nguồn ghi trong bản tin', /80%/.test(row2[10]) && !/Tự động/.test(row2[10]));
r = run('tqnWeatherJob(false)'); T('W10 Trigger chạy lần 2 cùng ngày → bỏ qua', r === 'DA_CHAY' && W.getLastRow() === 2, r);
r = run('tqnWeatherJob(true)'); T('W11 Bấm tay cùng ngày → cập nhật lại đúng dòng TQN, không thêm dòng', r === 'CAP_NHAT' && W.getLastRow() === 2, r);
W.set(2, 11, 'Nhập tay bởi BCH'); delete props.lastWeather; r = run('tqnWeatherJob(true)');
T('W12 Dòng cùng ngày do người nhập tay → không ghi đè', r === 'BO_QUA_NHAP_TAY' && W.get(2, 11) === 'Nhập tay bởi BCH', r);
T('W13 Chỉ ghi vào tab thời tiết + tab TQN_*', writes.every(w => w[0] === WEATHER_TAB || w[0].startsWith('TQN_')), [...new Set(writes.map(w => w[0]))].join(', '));
T('W14 URL đúng tham số + timezone dự án', /latitude=15/.test(lastUrl) && /timezone=Asia%2FHo_Chi_Minh/.test(lastUrl) && /precipitation_probability_max/.test(lastUrl) && /wind_speed_unit=kmh/.test(lastUrl));
T('W15 Lưu phản hồi gốc vào TQN_THOI_TIET_RAW (cả lần lỗi)', sheets['TQN_THOI_TIET_RAW'].getLastRow() >= 4);
W.set(2, 11, ''); W.set(2, 11, '📷 Bản tin hình');
r = run("tqnWeatherKhoang('17/11/2026', '18/11/2026')");
const ds = W.v.slice(1).map(x => x[0] instanceof Date ? x[0].getDate() : x[0]);
T('W16 Cập nhật khoảng ngày: thêm 17 và 18/11, chèn đúng thứ tự trước 19/11', /Đã thêm 2 ngày/.test(r) && ds.join(',') === '17,18,19', r + ' | ' + ds.join(','));
T('W17 Mỗi ngày có bản tin hình (PDF) + link trong Ghi chú', /2 bản tin/.test(r) && (W.links || []).length >= 2 && /^📷 Xem bản tin hình/.test(W.v[1][10]), (W.links || []).join(' '));
T('W18 Ngày đã qua ghi rõ "số liệu mô hình", bản tin có 3 buổi + nguồn', /Sáng 06:00–11:00/.test(pdfHtml[pdfHtml.length - 1]) && /Chiều 12:00/.test(pdfHtml[pdfHtml.length - 1]) && /23:00/.test(pdfHtml[pdfHtml.length - 1]) && /Tối/.test(pdfHtml[pdfHtml.length - 1]) && /Open-Meteo/.test(pdfHtml[pdfHtml.length - 1]));
fs.writeFileSync('/tmp/bantin_mau.html', pdfHtml[pdfHtml.length - 1]);
console.log(`KẾT QUẢ THỜI TIẾT: ${pass} PASS / ${fail} FAIL`); process.exit(fail ? 1 : 0);
