// Chạy ĐÚNG file gas/dist/main.js (bản sẽ push) trong giả lập Apps Script: kiểm SheetsRepo, LockService, quyền, idempotency.
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import vm from 'node:vm';

function fakeSs() {
  const sheets = new Map<string, any[][]>();
  const mk = (name: string) => {
    const g = () => sheets.get(name)!;
    const sh: any = {
      getName: () => name,
      getLastRow: () => g().length,
      setFrozenRows() {},
      appendRow(row: any[]) { g().push(row.map(v => (v === null || v === undefined ? '' : v))); },
      getRange(r: number, c: number, nr = 1, nc = 1) {
        return {
          getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (g()[r - 1 + i] || [])[c - 1 + j] ?? '')),
          setValues(v: any[][]) { v.forEach((row, i) => { const rr = (g()[r - 1 + i] ||= []); row.forEach((x, j) => { rr[c - 1 + j] = x === null || x === undefined ? '' : x; }); }); return this; },
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
const call = (fn: string, ...a: any[]) => JSON.parse(JSON.stringify(ctx[fn](...a)));

beforeAll(() => {
  if (!existsSync('gas/dist/main.js')) throw new Error('Chạy npm run build trước');
  ss = fakeSs();
  const props: Record<string, string> = { TQN6_MEMBERS: JSON.stringify({ 'nt@tbc.vn': ['CONTRACTOR'], 'cdt@coa.vn': ['REVIEWER'] }) };
  ctx = vm.createContext({
    console, JSON, Date, Math,
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k: string) => props[k] ?? null }) },
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
});
