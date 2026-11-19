// Fault injection theo yêu cầu ChatGPT (issue #2): lỗi ở từng bước ghi, rồi retry cùng idem → trạng thái ra sao?
// File này GHI LẠI hiện trạng (kể cả lỗi) – không phải test "phải xanh".
import { describe, it, expect } from 'vitest';
import { RegisterService } from '../src/core/service';
import { MemoryRepo, NoLock } from '../src/core/repo';
import { coastalConfig } from '../src/core/config';
import type { Actor } from '../src/core/types';

const NT: Actor = { email: 'nt@tbc.vn', kind: 'HUMAN' };
const LINK = { name: 'a.pdf', url: 'https://drive.google.com/a', source: 'DRIVE' as const };
type Buoc = 'saveDocument' | 'saveRevision' | 'appendEvent' | 'putIdem' | 'setCounter';
function chay(buoc: Buoc, op: 'create' | 'submit') {
  const repo = new MemoryRepo(), svc = new RegisterService(coastalConfig({ 'nt@tbc.vn': ['CONTRACTOR'] }), repo, new NoLock(), { now: () => new Date(2026, 10, 20, 9) });
  const pre = op === 'submit' ? svc.create(NT, 'pre-00001', { type: 'RFI', title: 'Hồ sơ có sẵn', files: [LINK] }) : null;
  const goc = (repo as any)[buoc].bind(repo); let hong = true;
  (repo as any)[buoc] = (...a: any[]) => { if (hong) { hong = false; throw new Error('Sheets lỗi giả lập ở ' + buoc); } return goc(...a); };
  const lam = () => op === 'create' ? svc.create(NT, 'idem-00001', { type: 'RFI', title: 'Hỏi thử lỗi giữa chừng', files: [LINK] }) : svc.submit(NT, 'idem-00002', pre!.documentId);
  let lan1 = 'OK'; try { lam(); } catch (e) { lan1 = 'LỖI'; }
  let lan2 = 'OK'; try { lam(); } catch (e) { lan2 = 'LỖI: ' + (e as Error).message; }
  const docs = [...repo.docs.values()], revs = [...repo.revs.values()];
  return { buoc, op, lan1, lan2, soHoSo: docs.length - (pre ? 1 : 0), soMa: new Set(docs.map(d => d.docNo)).size, hoSoThieuRev: docs.filter(d => !revs.some(r => r.documentId === d.documentId)).length,
    suKienOK: repo.events.filter(e => e.ok && e.action === op.toUpperCase()).length, trangThai: pre ? repo.docs.get(pre.documentId)!.status : docs[0]?.status };
}
describe('Fault injection – hiện trạng (19/11)', () => {
  const kq: any[] = [];
  for (const op of ['create', 'submit'] as const) for (const b of ['setCounter', 'saveDocument', 'saveRevision', 'appendEvent', 'putIdem'] as Buoc[])
    it(op + ' lỗi ở ' + b, () => { const r = chay(b, op); kq.push(r); expect(r).toBeTruthy(); });
  it('in bảng', () => { console.table(kq); });
});

// ===== Acceptance (ChatGPT issue #2 điểm 5): assert đầy đủ. Hiện CHƯA ĐẠT → đánh dấu it.fails;
// sau khi có journal + khôi phục phải đổi it.fails → it và PASS. =====
describe('Acceptance sau journal – hiện phải FAIL', () => {
  for (const op of ['create', 'submit'] as const) for (const b of ['setCounter', 'saveDocument', 'saveRevision', 'appendEvent', 'putIdem'] as Buoc[]) {
    const ky = op === 'create' && (b === 'setCounter' || b === 'saveDocument') || op === 'submit' && (b === 'setCounter' || b === 'saveDocument' || b === 'saveRevision');
    (ky ? it : it.fails)(op + ' lỗi ở ' + b + ': thử lại → OK, 1 hồ sơ, 1 mã, 0 mồ côi, 1 sự kiện thành công', () => {
      const r = chay(b, op);
      expect(r.lan2).toBe('OK');
      if (op === 'create') { expect(r.soHoSo).toBe(1); expect(r.soMa).toBe(1); }
      expect(r.hoSoThieuRev).toBe(0);
      expect(r.suKienOK).toBe(1);
    });
  }
});
