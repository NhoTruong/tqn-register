// Cầu nối giao diện ↔ backend. Chạy trong Apps Script: google.script.run. Chạy xem trước: dùng CHÍNH lõi nghiệp vụ trong bộ nhớ.
import { RegisterService } from '../core/service';
import { MemoryRepo, NoLock } from '../core/repo';
import { coastalConfig } from '../core/config';
import { classify } from '../core/classify';
import type { DocType, FileRef } from '../core/types';

type Res<T> = Promise<T>;
const inGas = typeof (globalThis as any).google !== 'undefined' && !!(globalThis as any).google?.script?.run;

function run<T>(fn: string, ...args: unknown[]): Res<T> {
  return new Promise((ok, bad) => {
    (globalThis as any).google.script.run
      .withSuccessHandler((r: any) => (r && r.error ? bad(new Error(r.error)) : ok(r)))
      .withFailureHandler((e: Error) => bad(e))[fn](...args);
  });
}

// ----- chế độ xem trước: 2 người dùng mẫu để thử cả hai phía -----
const demoUsers = { 'nhathau@demo': ['CONTRACTOR'], 'cdt@demo': ['REVIEWER'] } as Record<string, any>;
let demoMe = 'nhathau@demo';
const demo = new RegisterService(coastalConfig(demoUsers), new MemoryRepo(), new NoLock(), { now: () => new Date() });
(() => {   // dữ liệu mẫu ghi rõ DEMO
  const f: FileRef = { name: 'DEMO-RFI-gach.pdf', url: 'https://example.com/demo.pdf', source: 'LINK' };
  const a = { email: 'nhathau@demo', kind: 'HUMAN' as const }, c = { email: 'cdt@demo', kind: 'HUMAN' as const };
  const d1 = demo.create(a, 'seed-0001', { type: 'RFI', title: '[DEMO] Yêu cầu thông tin gạch porcelain và KCT', files: [f] });
  demo.submit(a, 'seed-0002', d1.documentId);
  const d2 = demo.create(a, 'seed-0003', { type: 'RFI', title: '[DEMO] Yêu cầu thông tin gạch mosaic', files: [f] });
  demo.submit(a, 'seed-0004', d2.documentId); demo.reply(c, 'seed-0005', d2.documentId, '2', 'Bổ sung catalogue màu');
  demo.create(a, 'seed-0006', { type: 'RFI', title: '[DEMO] Hỏi cao độ hố ga trục 5', files: [] });
})();
const me = () => ({ email: demoMe, kind: 'HUMAN' as const });
const wrap = <T,>(f: () => T): Res<T> => { try { return Promise.resolve(f()); } catch (e) { return Promise.reject(e); } };

export const isDemo = !inGas;
export function setDemoUser(e: string) { demoMe = e; }

export const api = {
  whoAmI: (): Res<{ email: string; roles: string[]; build: any }> => inGas ? run('apiWhoAmI') : wrap(() => ({ email: demoMe, roles: demo.rolesOf(demoMe), build: { env: 'PREVIEW', commit: 'local', builtAt: '', version: 'demo' } })),
  list: (t?: DocType) => inGas ? run<any[]>('apiList', t) : wrap(() => demo.list(t)),
  get360: (id: string) => inGas ? run<any>('apiGet360', id) : wrap(() => demo.get360(id)),
  create: (idem: string, input: { type: DocType; title: string; files?: FileRef[] }) => inGas ? run<any>('apiCreate', idem, input) : wrap(() => demo.create(me(), idem, input)),
  attach: (idem: string, id: string, f: FileRef) => inGas ? run<any>('apiAttach', idem, id, f) : wrap(() => demo.attach(me(), idem, id, f)),
  submit: (idem: string, id: string) => inGas ? run<any>('apiSubmit', idem, id) : wrap(() => demo.submit(me(), idem, id)),
  reply: (idem: string, id: string, code: string, note: string) => inGas ? run<any>('apiReply', idem, id, code, note) : wrap(() => demo.reply(me(), idem, id, code, note)),
  closeForwarded: (idem: string, id: string, note: string) => inGas ? run<any>('apiCloseForwarded', idem, id, note) : wrap(() => demo.closeForwarded(me(), idem, id, note)),
  classify: (name: string) => inGas ? run<any>('apiClassify', name) : wrap(() => classify(name)),
  options: () => coastalConfig().types,
};

/** idem_key: sinh 1 lần cho mỗi thao tác của người dùng; bấm 2 lần / mạng gửi lại vẫn dùng cùng khóa. */
export function newIdem() { return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10); }
