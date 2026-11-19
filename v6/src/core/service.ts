// Register Service – một hệ nghiệp vụ duy nhất (Numbering, Revision, Workflow, Audit, quyền, idempotency).
import { ulid, type Rand } from './ids';
import { nextNumber } from './numbering';
import type { Locker, Repo } from './repo';
import type { Actor, DocType, DocumentRec, EventRec, FileRef, ProjectConfig, RevisionRec, Role, DocStatus } from './types';

export class BizError extends Error { constructor(public code: string, msg: string) { super(msg); } }

export interface Clock { now(): Date }

export interface Doc360 { doc: DocumentRec; revisions: RevisionRec[]; events: EventRec[]; overdueDays: number }

const DAY = 864e5;

export class RegisterService {
  constructor(private cfg: ProjectConfig, private repo: Repo, private lock: Locker, private clock: Clock, private rand: Rand = Math.random) {}

  // ---------- quyền ----------
  rolesOf(email: string): Role[] { return this.cfg.members[email.toLowerCase()] || []; }
  private need(actor: Actor, roles: Role[], what: string) {
    const has = this.rolesOf(actor.email);
    if (!has.includes('ADMIN') && !roles.some(r => has.includes(r))) throw new BizError('FORBIDDEN', `${actor.email} không có quyền ${what} (cần ${roles.join('/')})`);
  }

  // ---------- khung chung cho mọi lệnh ghi: khóa + idempotency + audit ----------
  private mutate<T>(action: string, actor: Actor, idemKey: string, docId: string | null, fn: () => { result: T; revNo: number | null; detail: object; docId?: string }): T {
    if (!idemKey || idemKey.length < 8) throw new BizError('IDEM_REQUIRED', 'Thiếu idem_key cho lệnh ' + action);
    return this.lock.withLock(() => {
      const k = this.cfg.projectId + ':' + action + ':' + idemKey;
      const old = this.repo.getIdem(k);
      if (old) return JSON.parse(old) as T;                       // gửi lại → trả đúng kết quả cũ, không ghi thêm
      try {
        const r = fn();
        this.event(action, actor, r.docId ?? docId, r.revNo, idemKey, r.detail, true);
        this.repo.putIdem(k, JSON.stringify(r.result));
        return r.result;
      } catch (e) {
        this.event(action, actor, docId, null, idemKey, { error: (e as Error).message }, false);
        throw e;
      }
    });
  }
  private event(action: string, actor: Actor, docId: string | null, revNo: number | null, idemKey: string | null, detail: object, ok: boolean) {
    const at = this.clock.now();
    this.repo.appendEvent({ eventId: ulid(at.getTime(), this.rand), at: at.toISOString(), actor: actor.email, actorKind: actor.kind, action, documentId: docId, revNo, idemKey, detail: JSON.stringify(detail).slice(0, 2000), ok });
  }
  private load(id: string): DocumentRec { const d = this.repo.getDocument(id); if (!d) throw new BizError('NOT_FOUND', 'Không có hồ sơ ' + id); return d; }
  private rev(d: DocumentRec): RevisionRec {
    const r = this.repo.getRevisions(d.documentId).find(x => x.revNo === d.currentRev);
    if (!r) throw new BizError('DATA', 'Thiếu revision ' + d.currentRev + ' của ' + d.docNo); return r;
  }
  /** Bản đang soạn: DRAFT → rev hiện tại; bị yêu cầu bổ sung/trình lại → rev kế tiếp (tạo nếu chưa có). */
  private draftRev(d: DocumentRec): RevisionRec {
    if (d.status === 'DRAFT') return this.rev(d);
    const n = d.currentRev + 1, ex = this.repo.getRevisions(d.documentId).find(x => x.revNo === n);
    if (ex) return ex;
    const r: RevisionRec = { documentId: d.documentId, revNo: n, submittedAt: null, repliedAt: null, replyCode: null, replyText: null, replyNote: null, files: [] };
    this.repo.saveRevision(r); return r;
  }
  private setStatus(d: DocumentRec, s: DocStatus, holder: string, dueAt: string | null) { d.status = s; d.holder = holder; d.dueAt = dueAt; d.updatedAt = this.clock.now().toISOString(); }

  // ---------- 1. tạo hồ sơ (Numbering Engine) ----------
  create(actor: Actor, idemKey: string, input: { type: DocType; title: string; files?: FileRef[] }): DocumentRec {
    this.need(actor, ['DC', 'CONTRACTOR'], 'tạo hồ sơ');
    if (actor.kind !== 'HUMAN') throw new BizError('FORBIDDEN', 'Chỉ người dùng mới tạo hồ sơ chính thức (AI chỉ soạn nháp)');
    const title = String(input.title || '').trim();
    if (title.length < 5) throw new BizError('INVALID', 'Nội dung hồ sơ quá ngắn');
    return this.mutate('CREATE', actor, idemKey, null, () => {
      const now = this.clock.now(), key = this.cfg.projectId + ':' + input.type;
      const existing = this.repo.listDocuments(this.cfg.projectId, input.type).map(d => d.docNo);
      const { docNo, seq } = nextNumber(this.cfg, input.type, now, this.repo.getCounter(key), existing);
      this.repo.setCounter(key, seq);
      const d: DocumentRec = { documentId: ulid(now.getTime(), this.rand), projectId: this.cfg.projectId, docType: input.type, docNo, title, status: 'DRAFT', currentRev: 0, owner: actor.email.toLowerCase(), holder: 'CONTRACTOR', dueAt: null, createdAt: now.toISOString(), updatedAt: now.toISOString() };
      this.repo.saveDocument(d);
      this.repo.saveRevision({ documentId: d.documentId, revNo: 0, submittedAt: null, repliedAt: null, replyCode: null, replyText: null, replyNote: null, files: input.files || [] });
      return { result: d, revNo: 0, detail: { docNo, title }, docId: d.documentId };
    });
  }

  attach(actor: Actor, idemKey: string, docId: string, file: FileRef): RevisionRec {
    this.need(actor, ['DC', 'CONTRACTOR'], 'gắn file');
    return this.mutate('ATTACH', actor, idemKey, docId, () => {
      const d = this.load(docId);
      if (!['DRAFT', 'NEED_INFO', 'RESUBMIT_REQUIRED'].includes(d.status)) throw new BizError('STATE', 'Hồ sơ đã trình, không gắn thêm file vào bản đang chờ duyệt – hãy tạo rev mới');
      if (!/^https:\/\//.test(file.url)) throw new BizError('INVALID', 'Link file phải là https');
      const r = this.draftRev(d); r.files.push(file); this.repo.saveRevision(r);
      return { result: r, revNo: r.revNo, detail: { file: file.name } };
    });
  }

  // ---------- 2. trình ----------
  submit(actor: Actor, idemKey: string, docId: string): DocumentRec {
    this.need(actor, ['DC', 'CONTRACTOR'], 'trình hồ sơ');
    return this.mutate('SUBMIT', actor, idemKey, docId, () => {
      const d = this.load(docId), t = this.cfg.types[d.docType]!;
      if (!['DRAFT', 'NEED_INFO', 'RESUBMIT_REQUIRED'].includes(d.status)) throw new BizError('STATE', 'Không trình được khi đang ' + d.status);
      const r = this.draftRev(d);                                   // trình lại → revision mới, lịch sử cũ giữ nguyên
      if (!r.files.length) throw new BizError('INVALID', 'Chưa gắn file hồ sơ' + (r.revNo ? ' cho rev' + String(r.revNo).padStart(2, '0') : '') + ', không trình được');
      d.currentRev = r.revNo;
      const now = this.clock.now();
      r.submittedAt = now.toISOString(); this.repo.saveRevision(r);
      this.setStatus(d, 'SUBMITTED', t.reviewer, new Date(now.getTime() + t.slaDays * DAY).toISOString());
      this.repo.saveDocument(d);
      return { result: d, revNo: d.currentRev, detail: { dueAt: d.dueAt } };
    });
  }

  // ---------- 3. bên duyệt trả lời theo đúng ô trên phiếu (AI không được) ----------
  reply(actor: Actor, idemKey: string, docId: string, code: string, note: string): DocumentRec {
    if (actor.kind !== 'HUMAN') throw new BizError('FORBIDDEN', 'AI/hệ thống không được trả lời hay phê duyệt hồ sơ kỹ thuật');
    this.need(actor, ['REVIEWER'], 'trả lời/phê duyệt');
    return this.mutate('REPLY', actor, idemKey, docId, () => {
      const d = this.load(docId), t = this.cfg.types[d.docType]!;
      if (d.status !== 'SUBMITTED') throw new BizError('STATE', 'Chỉ trả lời hồ sơ đang chờ duyệt (đang ' + d.status + ')');
      const o = t.options.find(x => x.code === code);
      if (!o) throw new BizError('INVALID', 'Ô "' + code + '" không có trên phiếu ' + d.docType + ' của dự án');
      if (o.result !== 'CLOSE' && !note.trim()) throw new BizError('INVALID', 'Chọn ' + code + ' phải ghi ý kiến');
      const r = this.rev(d), now = this.clock.now();
      Object.assign(r, { repliedAt: now.toISOString(), replyCode: o.code, replyText: o.text, replyNote: note.trim() || null }); this.repo.saveRevision(r);
      const map: Record<string, [DocStatus, string]> = { CLOSE: ['CLOSED', ''], NEED_INFO: ['NEED_INFO', 'CONTRACTOR'], FORWARD: ['FORWARDED', t.reviewer], RESUBMIT: ['RESUBMIT_REQUIRED', 'CONTRACTOR'], REJECT: ['REJECTED', ''] };
      const [s, h] = map[o.result];
      this.setStatus(d, s, h, null); this.repo.saveDocument(d);
      return { result: d, revNo: d.currentRev, detail: { code: o.code, text: o.text } };
    });
  }

  /** Chuyển tiếp xong, bên nhận trả lời → đóng (RFI ô 3). */
  closeForwarded(actor: Actor, idemKey: string, docId: string, note: string): DocumentRec {
    if (actor.kind !== 'HUMAN') throw new BizError('FORBIDDEN', 'AI không được đóng hồ sơ');
    this.need(actor, ['REVIEWER'], 'đóng hồ sơ chuyển tiếp');
    return this.mutate('CLOSE_FORWARDED', actor, idemKey, docId, () => {
      const d = this.load(docId);
      if (d.status !== 'FORWARDED') throw new BizError('STATE', 'Hồ sơ không ở trạng thái chuyển tiếp');
      if (!note.trim()) throw new BizError('INVALID', 'Ghi câu trả lời của bên được chuyển tới');
      const r = this.rev(d); r.replyNote = [r.replyNote, 'Trả lời sau chuyển tiếp: ' + note.trim()].filter(Boolean).join('\n'); this.repo.saveRevision(r);
      this.setStatus(d, 'CLOSED', '', null); this.repo.saveDocument(d);
      return { result: d, revNo: d.currentRev, detail: { note } };
    });
  }

  // ---------- 4. Document 360° + danh sách ----------
  get360(docId: string): Doc360 {
    const doc = this.load(docId);
    return { doc, revisions: this.repo.getRevisions(docId), events: this.repo.listEvents(docId), overdueDays: this.overdue(doc) };
  }
  overdue(d: DocumentRec): number {
    if (!d.dueAt || d.status !== 'SUBMITTED') return 0;
    const n = Math.floor((this.clock.now().getTime() - new Date(d.dueAt).getTime()) / DAY);
    return n > 0 ? n : 0;
  }
  list(type?: DocType) { return this.repo.listDocuments(this.cfg.projectId, type).map(d => ({ ...d, overdueDays: this.overdue(d) })).sort((a, b) => b.createdAt.localeCompare(a.createdAt)); }

  /** Notification Engine: việc cần nhắc hôm nay (không gửi – adapter gửi). */
  reminders(): { to: string; docNo: string; title: string; why: string }[] {
    return this.list().flatMap(d => {
      if (d.status === 'SUBMITTED' && d.overdueDays > 0) return [{ to: 'REVIEWER', docNo: d.docNo, title: d.title, why: 'Quá hạn trả lời ' + d.overdueDays + ' ngày' }];
      if (d.status === 'NEED_INFO' || d.status === 'RESUBMIT_REQUIRED') return [{ to: d.owner, docNo: d.docNo, title: d.title, why: 'Bên duyệt yêu cầu bổ sung / trình lại' }];
      return [];
    });
  }
}
