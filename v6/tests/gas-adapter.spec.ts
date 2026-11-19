// Chạy ĐÚNG file gas/dist/main.js (bản sẽ push) trong giả lập Apps Script: kiểm SheetsRepo, LockService, quyền, idempotency.
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import vm from 'node:vm';
import { createHash } from 'node:crypto';

function fakeSs(noText = false) {
  const sheets = new Map<string, any[][]>();
  const textFmt = new Set<string>(); if (noText) textFmt.add = (() => textFmt) as any;
  const mk = (name: string) => {
    const g = () => sheets.get(name)!;
    const sh: any = {
      getName: () => name,
      getLastRow: () => g().length,
      getMaxRows: () => 1000,
      setFrozenRows() {},
      appendRow(row: any[]) { g().push(row.map(v => (v === null || v === undefined ? '' : (typeof v === 'string' && /^\d{4}-\d\d-\d\dT/.test(v) && !textFmt.has(name)) ? new Date(v) : v))); },
      getRange(r: number, c: number, nr = 1, nc = 1) {
        return {
          getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (g()[r - 1 + i] || [])[c - 1 + j] ?? '')),
          // giống Sheets thật: chuỗi dạng ngày ISO tự thành Date nếu ô không định dạng văn bản
          setValues(v: any[][]) { v.forEach((row, i) => { const rr = (g()[r - 1 + i] ||= []); row.forEach((x, j) => { rr[c - 1 + j] = x === null || x === undefined ? '' : (typeof x === 'string' && /^\d{4}-\d\d-\d\dT/.test(x) && !textFmt.has(name)) ? new Date(x) : x; }); }); return this; },
          setNumberFormat(f: string) { if (f === '@') textFmt.add(name); return this; },
          setFontWeight() { return this; },
        };
      },
    };
    return sh;
  };
  return {
    sheets,
    getSheetByName: (n: string) => (sheets.has(n) ? mk(n) : null),
    insertSheet: (n: string) => { sheets.set(n, []); return mk(n); },
  };
}

let ctx: any, ss: any, user = 'nt@tbc.vn', locks = 0, busyLock = false;
const props: Record<string, string> = {}, cacheMap = new Map<string, string>();
const drive = { files: new Map<string, { name: string; bytes: number[]; trashed: boolean }>(), n: 0, deny: false };
const call = (fn: string, ...a: any[]) => JSON.parse(JSON.stringify(ctx[fn](...a)));

beforeAll(() => {
  if (!existsSync('gas/dist/main.js')) throw new Error('Chạy npm run build trước');
  ss = fakeSs();
  props.TQN6_MEMBERS = JSON.stringify({ 'nt@tbc.vn': ['CONTRACTOR'], 'cdt@coa.vn': ['REVIEWER'], 'cdt2@coa.vn': ['REVIEWER'] });
  const mkFolder = (id: string) => ({ getId: () => id, createFile: (b: any) => {
    if (drive.deny) throw new Error('You do not have permission to call DriveApp.Folder.createFile');
    const fid = 'D' + ++drive.n; drive.files.set(fid, { name: b.name, bytes: b.bytes, trashed: false });
    return { getId: () => fid, getUrl: () => 'https://drive.google.com/file/d/' + fid, getName: () => b.name };
  } });
  ctx = vm.createContext({
    console, JSON, Date, Math,
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k: string) => props[k] ?? null, setProperty: (k: string, v: string) => { props[k] = v; }, getProperties: () => ({ ...props }), deleteProperty: (k: string) => { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({ get: (k: string) => cacheMap.get(k) ?? null, put: (k: string, v: string) => { cacheMap.set(k, v); } }) },
    DriveApp: {
      getFolderById: (id: string) => { if (!id || id === 'null') throw new Error('no folder'); return mkFolder(id); },
      createFolder: () => mkFolder('FOLDER1'),
      getFileById: (id: string) => { const f = drive.files.get(id); if (!f) throw new Error('not found'); return { isTrashed: () => f.trashed }; },
    },
    Utilities: {
      base64Decode: (b: string) => [...Buffer.from(b, 'base64')].map(x => (x > 127 ? x - 256 : x)),
      newBlob: (bytes: number[], mime: string, name: string) => ({ bytes, mime, name }),
      DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (_a: string, b: number[]) => [...createHash('sha256').update(Buffer.from(b.map(x => (x + 256) % 256))).digest()].map(x => (x > 127 ? x - 256 : x)),
    },
    SpreadsheetApp: { getActive: () => ss, openById: () => ss, flush() {} },
    LockService: { getDocumentLock: () => ({ tryLock: () => { if (busyLock) return false; locks++; return true; }, releaseLock() {} }), getScriptLock: () => null },
    Session: { getActiveUser: () => ({ getEmail: () => user }), getEffectiveUser: () => ({ getEmail: () => user }) },
    HtmlService: {}, Logger: { log() {} },
  });
  vm.runInContext(readFileSync('gas/dist/main.js', 'utf8'), ctx);
});

describe('Adapter Apps Script (file build thật)', () => {
  it('có đủ hàm API toàn cục và dấu phiên bản build', () => {
    for (const f of ['doGet', 'apiWhoAmI', 'apiList', 'apiGet360', 'apiCreate', 'apiAttach', 'apiSubmit', 'apiReply', 'apiCloseForwarded', 'tqn6Version']) expect(typeof ctx[f]).toBe('function');
    const b = call('tqn6Version'); expect(b.commit).toBeTruthy(); expect(b.builtAt).toMatch(/^\d{4}-/);
  });
  it('RFI trọn vòng ghi vào tab TQN6_*, đọc lại đúng từ sheet', () => {
    const d = call('apiCreate', 'gas-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch', files: [{ name: 'a.pdf', url: 'https://drive.google.com/x', source: 'DRIVE' }] });
    expect(d.docNo).toMatch(/-COA-TBC-RFI-001$/);
    expect(ss.sheets.get('TQN6_DOCS').length).toBe(2);
    call('apiSubmit', 'gas-0002', d.documentId);
    user = 'cdt@coa.vn';
    const r = call('apiReply', 'gas-0003', d.documentId, '1', '');
    expect(r.status).toBe('CLOSED');
    const v = call('apiGet360', d.documentId);
    expect(v.revisions[0].files[0].name).toBe('a.pdf');
    expect(v.events.map((e: any) => e.action)).toEqual(['CREATE', 'SUBMIT', 'REPLY']);
    expect(locks).toBeGreaterThanOrEqual(3);
    // kết quả trả về web không được chứa Date (google.script.run sẽ trả null) – lỗi thấy trên DEV 22:28
    const l = call('apiList', 'RFI'); expect(l.length).toBe(1); expect(typeof l[0].createdAt).toBe('string');
    const hasDate = (o: any): boolean => o instanceof Date || (o && typeof o === 'object' && Object.values(o).some(hasDate));
    expect(hasDate(ctx.apiList('RFI'))).toBe(false);
  });
  it('gửi trùng (bấm 2 lần) → 1 hồ sơ; thiếu quyền → lỗi rõ, không ghi', () => {
    user = 'nt@tbc.vn';
    const a = call('apiCreate', 'gas-dup-01', { type: 'RFI', title: 'Hỏi cao độ hố ga' });
    const b = call('apiCreate', 'gas-dup-01', { type: 'RFI', title: 'Hỏi cao độ hố ga' });
    expect(b.documentId).toBe(a.documentId);
    expect(ss.sheets.get('TQN6_DOCS').length).toBe(3);
    user = 'la@x.vn';
    expect(call('apiCreate', 'gas-x-0001', { type: 'RFI', title: 'Người lạ tạo' }).code).toBe('FORBIDDEN');
    expect(ss.sheets.get('TQN6_DOCS').length).toBe(3);
  });
  it('đang có người khác giữ khóa → báo bận, không ghi dở', () => {
    user = 'nt@tbc.vn'; busyLock = true;
    expect(call('apiCreate', 'gas-busy-01', { type: 'RFI', title: 'Thử lúc bận' }).error).toMatch(/BUSY/);
    busyLock = false;
    expect(ss.sheets.get('TQN6_DOCS').length).toBe(3);
  });
  it('apiDo: 1 lần gọi trả cả kết quả + danh sách + hồ sơ 360 + thời gian; apiBoot đủ dữ liệu mở web', () => {
    user = 'nt@tbc.vn';
    const r = call('apiDo', 'create', 'do-000001', { type: 'RFI', title: 'Thử gộp 1 lần gọi', files: [{ name: 'x.pdf', url: 'https://drive.google.com/x', source: 'DRIVE' }] });
    expect(r.result.docNo).toMatch(/RFI-\d{3}$/); expect(r.list.length).toBeGreaterThan(0); expect(r.v360.doc.documentId).toBe(r.result.documentId);
    expect(typeof r.ms.total).toBe('number'); expect(r.ms.calls).toBeTruthy(); expect(r.dv).toBeTruthy();
    const s = call('apiDo', 'submit', 'do-000002', { id: r.result.documentId }); expect(s.v360.doc.status).toBe('SUBMITTED');
    const b = call('apiBoot', r.result.documentId); expect(b.who.email).toBe('nt@tbc.vn'); expect(b.v360.doc.documentId).toBe(r.result.documentId);
    expect(call('apiDo', 'xoa', 'do-000003', {}).code).toBe('INVALID');
  });
  it('7. Cache cũ ≠ dữ liệu thật: khóa cache theo dv → sau khi người khác ghi, mở lại thấy bản mới; peek báo đổi', () => {
    user = 'nt@tbc.vn';
    const c = call('apiDo', 'create', 'ca-000001', { type: 'RFI', title: 'Thử cache', files: [{ name: 'x.pdf', url: 'https://drive.google.com/x', source: 'DRIVE' }] });
    const b1 = call('apiBoot', c.result.documentId), dv1 = call('apiPeek').dv;
    expect(b1.v360.doc.status).toBe('DRAFT'); expect(b1.dv).toBe(dv1);
    expect(call('apiBoot', c.result.documentId).ms.calls.docSheets ?? 0).toBe(0);   // lần 2: lấy từ cache, không đọc Sheets
    call('apiDo', 'submit', 'ca-000002', { id: c.result.documentId });                 // "người khác" ghi
    const dv2 = call('apiPeek').dv; expect(dv2).not.toBe(dv1);
    const b2 = call('apiBoot', c.result.documentId); expect(b2.v360.doc.status).toBe('SUBMITTED'); expect(b2.list.find((x: any) => x.documentId === c.result.documentId).status).toBe('SUBMITTED');
  });
  it('6b. Xung đột qua API: gửi version cũ → CONFLICT', () => {
    user = 'nt@tbc.vn';
    const c = call('apiDo', 'create', 'cf-000001', { type: 'RFI', title: 'Thử xung đột', files: [{ name: 'x.pdf', url: 'https://drive.google.com/x', source: 'DRIVE' }] });
    const s = call('apiDo', 'submit', 'cf-000002', { id: c.result.documentId, ev: c.result.version });
    user = 'cdt@coa.vn'; call('apiDo', 'reply', 'cf-000003', { id: c.result.documentId, code: '1', note: '', ev: s.v360.doc.version });
    user = 'cdt2@coa.vn'; const r = call('apiDo', 'reply', 'cf-000004', { id: c.result.documentId, code: '2', note: 'muộn', ev: s.v360.doc.version });
    expect(r.code).toBe('CONFLICT');
  });
  it('Upload qua Apps Script thật (giả DriveApp): job COMPLETED, tên theo mã_rev, hash; thiếu quyền Drive → DRIVE_PERM, Thử lại xong', () => {
    user = 'nt@tbc.vn';
    const c = call('apiDo', 'create', 'up-c00001', { type: 'RFI', title: 'Thử tải file' });
    const b64 = Buffer.from('%PDF-1.4 noi dung that').toString('base64'), want = createHash('sha256').update('%PDF-1.4 noi dung that').digest('hex');
    drive.deny = true;
    const e = call('apiUpload', 'job-a00001', c.result.documentId, 'Phieu RFI.pdf', 'application/pdf', b64, want);
    expect(e.code).toBe('DRIVE_PERM'); expect(e.job.failedAt).toBe('UPLOADING');
    drive.deny = false;
    const r = call('apiUpload', 'job-a00001', c.result.documentId, 'Phieu RFI.pdf', 'application/pdf', b64, want);
    expect(r.job.state).toBe('COMPLETED'); expect(r.job.sha256).toBe(want); expect(r.job.storedName).toBe(c.result.docNo + '_rev00_Phieu RFI.pdf');
    expect(r.v360.revisions[0].files[0].url).toMatch(/^https:\/\/drive.google.com\/file\/d\/D\d+/);
    const n = drive.files.size; call('apiUpload', 'job-a00001', c.result.documentId, 'Phieu RFI.pdf', 'application/pdf', b64, want);
    expect(drive.files.size).toBe(n);
    const bad = call('apiUpload', 'job-a00002', c.result.documentId, 'x.pdf', 'application/pdf', b64, 'ffff');
    expect(bad.code).toBe('INTEGRITY');
    expect(typeof r.ms.drive).toBe('number');
  });
});

