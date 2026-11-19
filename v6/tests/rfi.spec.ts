import { describe, it, expect, beforeEach } from 'vitest';
import { RegisterService, BizError } from '../src/core/service';
import { MemoryRepo, NoLock } from '../src/core/repo';
import { coastalConfig } from '../src/core/config';
import { classify } from '../src/core/classify';
import { formatNo, nextNumber, noKey } from '../src/core/numbering';
import { isUlid } from '../src/core/ids';
import type { Actor } from '../src/core/types';

const NT: Actor = { email: 'nt@tbc.vn', kind: 'HUMAN' };
const CDT: Actor = { email: 'cdt@coa.vn', kind: 'HUMAN' };
const AI: Actor = { email: 'ai@tqn', kind: 'AI' };
const LA: Actor = { email: 'la@x.vn', kind: 'HUMAN' };
const FILE = { name: 'RFI-gach.pdf', url: 'https://drive.google.com/file/d/abc', source: 'DRIVE' as const };

let now = new Date(2026, 10, 19, 9, 0);
const clock = { now: () => now };
let seed = 1; const rand = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
let repo: MemoryRepo, svc: RegisterService;

beforeEach(() => {
  now = new Date(2026, 10, 19, 9, 0); seed = 1;
  repo = new MemoryRepo();
  svc = new RegisterService(coastalConfig({ 'nt@tbc.vn': ['CONTRACTOR'], 'cdt@coa.vn': ['REVIEWER'], 'ai@tqn': ['DC'] }), repo, new NoLock(), clock, rand);
});

describe('Numbering Engine', () => {
  it('cấp mã theo mẫu dự án, ID nghiệp vụ là ULID (không phải số dòng)', () => {
    const d = svc.create(NT, 'idem-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch porcelain', files: [FILE] });
    expect(d.docNo).toBe('261119-COA-TBC-RFI-001');
    expect(isUlid(d.documentId)).toBe(true);
    const d2 = svc.create(NT, 'idem-0002', { type: 'RFI', title: 'Yêu cầu thông tin gạch mosaic' });
    expect(d2.docNo).toBe('261119-COA-TBC-RFI-002');
  });
  it('bỏ qua số đã dùng (kể cả khác ngày đầu mã) – chống trùng', () => {
    const cfg = coastalConfig();
    const r = nextNumber(cfg, 'RFI', new Date(2026, 10, 20), 0, ['261118-COA-TBC-RFI-001', '261119-COA-TBC-RFI-002 rev00']);
    expect(r.docNo).toBe('261120-COA-TBC-RFI-003');
    expect(noKey('261116-COA-TBC-RFA-COS-001-rev02')).toBe('COA-TBC-RFA-COS-001');
  });
  it('mẫu có ô lạ thì báo lỗi rõ', () => {
    expect(() => formatNo('{abc}', { cfg: coastalConfig(), type: 'RFI', date: new Date(), seq: 1 })).toThrow(/ô lạ/);
  });
});

describe('Idempotency + Audit', () => {
  it('gửi lại cùng idem_key không tạo hồ sơ thứ hai, không ghi log thứ hai', () => {
    const a = svc.create(NT, 'idem-same-1', { type: 'RFI', title: 'Yêu cầu thông tin gạch', files: [FILE] });
    const b = svc.create(NT, 'idem-same-1', { type: 'RFI', title: 'Yêu cầu thông tin gạch', files: [FILE] });
    expect(b.documentId).toBe(a.documentId);
    expect(repo.docs.size).toBe(1);
    expect(repo.events.filter(e => e.action === 'CREATE').length).toBe(1);
  });
  it('thiếu idem_key → từ chối', () => {
    expect(() => svc.create(NT, '', { type: 'RFI', title: 'Yêu cầu thông tin' })).toThrow(/idem_key/);
  });
  it('lệnh lỗi cũng được ghi audit (ok=false)', () => {
    const d = svc.create(NT, 'idem-a-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch' });
    expect(() => svc.submit(NT, 'idem-a-0002', d.documentId)).toThrow(/Chưa gắn file/);
    expect(repo.events.some(e => e.action === 'SUBMIT' && !e.ok)).toBe(true);
  });
});

describe('Quyền + AI không được duyệt', () => {
  it('người ngoài danh sách không tạo được', () => {
    expect(() => svc.create(LA, 'idem-x-0001', { type: 'RFI', title: 'Yêu cầu thông tin' })).toThrow(BizError);
  });
  it('AI không tạo hồ sơ chính thức, không trả lời/duyệt', () => {
    expect(() => svc.create(AI, 'idem-ai-0001', { type: 'RFI', title: 'AI soạn thử nội dung' })).toThrow(/AI/);
    const d = svc.create(NT, 'idem-b-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch', files: [FILE] });
    svc.submit(NT, 'idem-b-0002', d.documentId);
    expect(() => svc.reply(AI, 'idem-b-0003', d.documentId, '1', '')).toThrow(/AI/);
  });
  it('nhà thầu không tự trả lời hồ sơ của mình', () => {
    const d = svc.create(NT, 'idem-c-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch', files: [FILE] });
    svc.submit(NT, 'idem-c-0002', d.documentId);
    expect(() => svc.reply(NT, 'idem-c-0003', d.documentId, '1', '')).toThrow(/quyền/);
  });
});

describe('RFI từ đầu đến cuối (vertical slice)', () => {
  it('Tạo → trình → CĐT yêu cầu bổ sung (ô 2) → trình lại rev01 → CĐT trả lời (ô 1) → đóng; đủ lịch sử', () => {
    const d = svc.create(NT, 'idem-e-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch porcelain và KCT', files: [FILE] });
    const s = svc.submit(NT, 'idem-e-0002', d.documentId);
    expect(s.status).toBe('SUBMITTED'); expect(s.holder).toBe('CĐT');
    expect(s.dueAt).toBe(new Date(2026, 10, 22, 9, 0).toISOString());           // SLA 3 ngày

    now = new Date(2026, 10, 20, 10, 0);
    expect(() => svc.reply(CDT, 'idem-e-0003', d.documentId, '2', '')).toThrow(/ghi ý kiến/);
    const r1 = svc.reply(CDT, 'idem-e-0004', d.documentId, '2', 'Bổ sung catalogue màu');
    expect(r1.status).toBe('NEED_INFO'); expect(r1.holder).toBe('CONTRACTOR');
    expect(svc.reminders()).toEqual([{ to: 'nt@tbc.vn', docNo: d.docNo, title: d.title, why: 'Bên duyệt yêu cầu bổ sung / trình lại' }]);

    expect(() => svc.submit(NT, 'idem-e-0005', d.documentId)).toThrow(/rev01/);         // chưa gắn file bổ sung
    svc.attach(NT, 'idem-e-0006', d.documentId, { name: 'catalogue.pdf', url: 'https://1drv.ms/b/x', source: 'ONEDRIVE' });
    const s2 = svc.submit(NT, 'idem-e-0007', d.documentId);
    expect(s2.currentRev).toBe(1); expect(s2.status).toBe('SUBMITTED');
    expect(svc.get360(d.documentId).revisions.map(r => r.files.map(f => f.name))).toEqual([['RFI-gach.pdf'], ['catalogue.pdf']]);
    expect(() => svc.attach(NT, 'idem-e-0008', d.documentId, FILE)).toThrow(/đang chờ duyệt/);
  });
  it('quá hạn tự tính + nhắc bên duyệt; trả lời ô 1 thì đóng; Document 360 có 2 revision và toàn bộ sự kiện', () => {
    const d = svc.create(NT, 'idem-f-0001', { type: 'RFI', title: 'Yêu cầu thông tin gạch mosaic', files: [FILE] });
    svc.submit(NT, 'idem-f-0002', d.documentId);
    svc.reply(CDT, 'idem-f-0003', d.documentId, '2', 'Thiếu mẫu');
    now = new Date(2026, 10, 20, 9, 0);
    svc.attach(NT, 'idem-f-0035', d.documentId, { name: 'mau.pdf', url: 'https://1drv.ms/b/y', source: 'ONEDRIVE' });
    svc.submit(NT, 'idem-f-0004', d.documentId);
    now = new Date(2026, 10, 25, 9, 0);                                             // hạn 23/11 → trễ 2 ngày
    expect(svc.list('RFI')[0].overdueDays).toBe(2);
    expect(svc.reminders()[0]).toMatchObject({ to: 'REVIEWER', why: 'Quá hạn trả lời 2 ngày' });
    const c = svc.reply(CDT, 'idem-f-0005', d.documentId, '1', '');
    expect(c.status).toBe('CLOSED'); expect(svc.reminders()).toEqual([]);
    const v = svc.get360(d.documentId);
    expect(v.revisions.map(r => [r.revNo, r.replyCode])).toEqual([[0, '2'], [1, '1']]);
    expect(v.events.map(e => e.action)).toEqual(['CREATE', 'SUBMIT', 'REPLY', 'ATTACH', 'SUBMIT', 'REPLY']);
    expect(() => svc.submit(NT, 'idem-f-0006', d.documentId)).toThrow(/Không trình được/);
  });
  it('ô 3 chuyển tiếp → đóng khi có trả lời của bên được chuyển tới', () => {
    const d = svc.create(NT, 'idem-g-0001', { type: 'RFI', title: 'Hỏi cao độ hố ga', files: [FILE] });
    svc.submit(NT, 'idem-g-0002', d.documentId);
    expect(svc.reply(CDT, 'idem-g-0003', d.documentId, '3', 'Chuyển TVTK').status).toBe('FORWARDED');
    expect(svc.closeForwarded(CDT, 'idem-g-0004', d.documentId, 'TVTK: cao độ -1.20').status).toBe('CLOSED');
  });
  it('ô không có trên phiếu dự án bị từ chối', () => {
    const d = svc.create(NT, 'idem-h-0001', { type: 'RFI', title: 'Hỏi cao độ hố ga', files: [FILE] });
    svc.submit(NT, 'idem-h-0002', d.documentId);
    expect(() => svc.reply(CDT, 'idem-h-0003', d.documentId, 'A', 'x')).toThrow(/không có trên phiếu/);
  });
});

describe('Document Classification (chỉ đề xuất)', () => {
  it('nhận RFI/RFA/SDS/MSS từ tên file, có rev', () => {
    expect(classify('261118-COA-TBC-RFI-001 rev00.pdf')).toMatchObject({ type: 'RFI', rev: 0 });
    expect(classify('03. RFA. Yêu cầu phê duyệt.docx').type).toBe('RFA');
    expect(classify('Shop drawing thép móng R01.dwg')).toMatchObject({ type: 'SDS', rev: 1 });
    expect(classify('BPTC cọc khoan nhồi.pdf').type).toBe('MSS');
    expect(classify('anh_chup.jpg').type).toBeNull();
  });
});
