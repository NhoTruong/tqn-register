// 10 kịch bản anh Nho yêu cầu (19/11 22:59): chống trùng, đồng thời, upload lỗi giữa chừng, mất mạng, xung đột, quyền Drive.
import { describe, it, expect, beforeEach } from 'vitest';
import { RegisterService, BizError } from '../src/core/service';
import { MemoryRepo, type Locker } from '../src/core/repo';
import { coastalConfig } from '../src/core/config';
import { runUpload, MemoryFileStore, MemoryJobStore } from '../src/core/upload';
import { createHash } from 'node:crypto';
import type { Actor } from '../src/core/types';

const NT: Actor = { email: 'nt@tbc.vn', kind: 'HUMAN' }, NT2: Actor = { email: 'nt2@tbc.vn', kind: 'HUMAN' };
const C1: Actor = { email: 'cdt@coa.vn', kind: 'HUMAN' }, C2: Actor = { email: 'cdt2@coa.vn', kind: 'HUMAN' };
const LINK = { name: 'a.pdf', url: 'https://drive.google.com/a', source: 'DRIVE' as const };
const members = { 'nt@tbc.vn': ['CONTRACTOR'], 'nt2@tbc.vn': ['CONTRACTOR'], 'cdt@coa.vn': ['REVIEWER'], 'cdt2@coa.vn': ['REVIEWER'] } as any;
const sha = (b: unknown) => createHash('sha256').update(Buffer.from(b as string)).digest('hex');

/** Khóa thật kiểu LockService: đang giữ thì người khác nhận BUSY; có thể ép bận. */
class TestLock implements Locker { held = false; forceBusy = 0; calls = 0;
  withLock<T>(fn: () => T): T { this.calls++; if (this.forceBusy > 0) { this.forceBusy--; throw new Error('BUSY: đang bận'); } if (this.held) throw new Error('BUSY'); this.held = true; try { return fn(); } finally { this.held = false; } } }

let repo: MemoryRepo, lock: TestLock, now: Date, svc: RegisterService, files: MemoryFileStore, jobs: MemoryJobStore;
const mk = () => new RegisterService(coastalConfig(members), repo, lock, { now: () => now });
const up = (actor: Actor, jobId: string, id: string, name: string, content: string) =>
  runUpload({ svc, files, jobs, actor, now: () => now, sha256: sha, folderKey: 'COA_TEST' }, { jobId, documentId: id, name, mime: 'application/pdf', size: content.length, bytes: content });
beforeEach(() => { repo = new MemoryRepo(); lock = new TestLock(); now = new Date(2026, 10, 19, 9); svc = mk(); files = new MemoryFileStore(); jobs = new MemoryJobStore(); });
const count = (action: string) => repo.events.filter(e => e.action === action && e.ok).length;

describe('An toàn ghi', () => {
  it('1. Double-click tạo hồ sơ → 1 hồ sơ, 1 mã, 1 dòng nhật ký', () => {
    const a = svc.create(NT, 'click-0001', { type: 'RFI', title: 'Hỏi cao độ hố ga' });
    const b = svc.create(NT, 'click-0001', { type: 'RFI', title: 'Hỏi cao độ hố ga' });
    expect(b.documentId).toBe(a.documentId); expect(repo.docs.size).toBe(1); expect(count('CREATE')).toBe(1);
  });
  it('2. Hai người cấp số cùng lúc → số khác nhau, không trùng; bộ đếm bị lùi vẫn không trùng', () => {
    const s2 = mk(), nos = new Set<string>();
    for (let i = 0; i < 10; i++) { nos.add(svc.create(NT, 'a-' + i + '-xxxx', { type: 'RFI', title: 'Người A lần ' + i }).docNo); nos.add(s2.create(NT2, 'b-' + i + '-xxxx', { type: 'RFI', title: 'Người B lần ' + i }).docNo); }
    expect(nos.size).toBe(20);
    repo.counters.clear();                                         // ai đó xóa nhầm bộ đếm
    expect(nos.has(svc.create(NT, 'c-0000-xxxx', { type: 'RFI', title: 'Sau khi mất bộ đếm' }).docNo)).toBe(false);
    // đang ghi dở (khóa đang giữ) → người thứ hai nhận BUSY, không ghi song song
    lock.held = true; expect(() => s2.create(NT2, 'd-0000-xxxx', { type: 'RFI', title: 'Bấm cùng lúc' })).toThrow(/BUSY/); lock.held = false;
  });
  it('3. Upload lên Drive xong nhưng ghi Sheets lỗi → giữ job + fileId; Thử lại cùng job không tải lại, 1 file, 1 bản ghi', () => {
    const d = svc.create(NT, 'cr-000001', { type: 'RFI', title: 'Hỏi gạch porcelain' });
    const orig = repo.saveRevision.bind(repo); let failOnce = true;
    repo.saveRevision = (r) => { if (failOnce && r.files.length) { failOnce = false; throw new Error('Sheets: Service timed out'); } orig(r); };
    expect(() => up(NT, 'job-000001', d.documentId, 'gach.pdf', 'NOI-DUNG-1')).toThrow(/COMMITTING/);
    const j = jobs.get('job-000001')!; expect(j.state).toBe('FAILED'); expect(j.failedAt).toBe('COMMITTING'); expect(j.fileId).toBeTruthy();
    expect(files.files.size).toBe(1);
    const r = up(NT, 'job-000001', d.documentId, 'gach.pdf', 'NOI-DUNG-1');
    expect(r.job.state).toBe('COMPLETED'); expect(files.files.size).toBe(1);
    expect(svc.get360(d.documentId).revisions[0].files.length).toBe(1);
    expect(svc.get360(d.documentId).revisions[0].files[0].fileId).toBe(j.fileId);
  });
  it('4. Mất mạng khi đang lưu (máy chủ đã ghi, web không nhận được) → gửi lại cùng khóa trả đúng kết quả, không ghi thêm', () => {
    const d = svc.create(NT, 'cr-000002', { type: 'RFI', title: 'Hỏi cửa sổ', files: [LINK] });
    const r1 = svc.submit(NT, 'sub-000001', d.documentId);       // phản hồi bị mất trên đường về
    now = new Date(now.getTime() + 60e3);
    const r2 = svc.submit(NT, 'sub-000001', d.documentId);       // web Thử lại
    expect(r2).toEqual(r1); expect(count('SUBMIT')).toBe(1);
  });
  it('5. Thử lại sau timeout/BUSY → lần sau ghi đúng 1 lần', () => {
    const d = svc.create(NT, 'cr-000003', { type: 'RFI', title: 'Hỏi lan can', files: [LINK] });
    lock.forceBusy = 1;
    expect(() => svc.submit(NT, 'sub-000002', d.documentId)).toThrow(/BUSY/);
    expect(svc.get360(d.documentId).doc.status).toBe('DRAFT');    // chưa ghi gì
    expect(svc.submit(NT, 'sub-000002', d.documentId).status).toBe('SUBMITTED');
    expect(svc.submit(NT, 'sub-000002', d.documentId).status).toBe('SUBMITTED'); expect(count('SUBMIT')).toBe(1);
  });
  it('6. Người sửa hồ sơ đã được người khác cập nhật → CONFLICT, không ghi đè', () => {
    const d = svc.create(NT, 'cr-000004', { type: 'RFI', title: 'Hỏi màu sơn', files: [LINK] });
    const s = svc.submit(NT, 'sub-000003', d.documentId), seen = s.version;   // 2 CĐT cùng mở, thấy cùng bản
    svc.reply(C1, 'rep-000001', d.documentId, '2', 'Bổ sung catalogue', seen);
    let er: any; try { svc.reply(C2, 'rep-000002', d.documentId, '1', '', seen); } catch (e) { er = e; }
    expect(er).toBeInstanceOf(BizError); expect(er.code).toBe('CONFLICT');
    const v = svc.get360(d.documentId); expect(v.doc.status).toBe('NEED_INFO'); expect(v.revisions[0].replyCode).toBe('2');
    expect(v.doc.version).toBe(seen + 1);
  });
  it('8. File cùng tên khác nội dung → giữ cả 2, tên thứ 2 có đuôi hash; cùng nội dung → không tải lại', () => {
    const d = svc.create(NT, 'cr-000005', { type: 'RFI', title: 'Hỏi bản vẽ cầu thang' });
    const a = up(NT, 'job-000101', d.documentId, 'BV cau thang.pdf', 'BAN-1');
    const b = up(NT, 'job-000102', d.documentId, 'BV cau thang.pdf', 'BAN-2');
    expect(a.job.storedName).toBe(d.docNo + '_rev00_BV cau thang.pdf');
    expect(b.job.storedName).toBe(d.docNo + '_rev00_BV cau thang_' + sha('BAN-2').slice(0, 8) + '.pdf');
    const c = up(NT, 'job-000103', d.documentId, 'ten-khac.pdf', 'BAN-1');
    expect(c.job.note).toMatch(/cùng nội dung/); expect(files.files.size).toBe(2);
    expect(svc.get360(d.documentId).revisions[0].files.length).toBe(2);
  });
  it('9. Thiếu quyền Google Drive → DRIVE_PERM rõ ràng, job FAILED ở UPLOADING, không có bản ghi', () => {
    const d = svc.create(NT, 'cr-000006', { type: 'RFI', title: 'Hỏi bản vẽ móng' });
    files.failNext = 'Exception: You do not have permission to call DriveApp.createFile';
    let er: any; try { up(NT, 'job-000201', d.documentId, 'mong.pdf', 'X'); } catch (e) { er = e; }
    expect(er.code).toBe('DRIVE_PERM'); expect(er.message).toMatch(/Cho phép/);
    expect(jobs.get('job-000201')!.failedAt).toBe('UPLOADING'); expect(files.files.size).toBe(0);
    expect(svc.get360(d.documentId).revisions[0].files.length).toBe(0);
    expect(up(NT, 'job-000201', d.documentId, 'mong.pdf', 'X').job.state).toBe('COMPLETED');   // cấp quyền xong → Thử lại cùng job
  });
  it('10. Chạy lại cùng JobID đã xong → không sinh hồ sơ/file/bản ghi trùng', () => {
    const d = svc.create(NT, 'cr-000007', { type: 'RFI', title: 'Hỏi khe lún' });
    up(NT, 'job-000301', d.documentId, 'khe.pdf', 'K'); up(NT, 'job-000301', d.documentId, 'khe.pdf', 'K'); up(NT, 'job-000301', d.documentId, 'khe.pdf', 'K');
    expect(files.files.size).toBe(1); expect(svc.get360(d.documentId).revisions[0].files.length).toBe(1); expect(count('ATTACH')).toBe(1);
  });
  it('Kiểm định dạng / dung lượng / hồ sơ đã trình / sai quyền trước khi tải', () => {
    const d = svc.create(NT, 'cr-000008', { type: 'RFI', title: 'Hỏi kính', files: [LINK] });
    expect(() => up(NT, 'job-000401', d.documentId, 'virus.exe', 'x')).toThrow(/Định dạng/);
    expect(() => up(NT, 'job-000402', d.documentId, 'a.pdf', '')).toThrow(/rỗng/);
    svc.submit(NT, 'sub-000009', d.documentId);
    expect(() => up(NT, 'job-000403', d.documentId, 'b.pdf', 'y')).toThrow(/đã trình/);
    expect(() => up(C1, 'job-000404', d.documentId, 'c.pdf', 'z')).toThrow(/quyền/);
    expect(files.files.size).toBe(0);
  });
});
