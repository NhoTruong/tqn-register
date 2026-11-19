// Upload Engine: tải file thật theo phiếu công việc (job). Không phụ thuộc Drive/Sheets – chỉ qua FileStore/JobStore.
// Trạng thái: UPLOADING → VALIDATING → COMMITTING → COMPLETED (lỗi → FAILED, giữ bước + fileId để chạy lại cùng jobId).
// Chỉ báo xong khi có CẢ fileId trên kho file VÀ bản ghi revision hợp lệ.
import { BizError, type RegisterService } from './service';
import type { Actor, FileRef } from './types';

export type JobState = 'UPLOADING' | 'VALIDATING' | 'COMMITTING' | 'COMPLETED' | 'FAILED';
export interface UploadJob {
  jobId: string; documentId: string; docNo?: string; name: string; size: number; mime: string;
  sha256?: string; storedName?: string; fileId?: string; url?: string; revNo?: number;
  state: JobState; failedAt?: JobState; error?: string; note?: string; by: string; createdAt: string; updatedAt: string;
}
export interface StoredFile { id: string; url: string; name: string }
export interface FileStore {
  exists(id: string): boolean;
  create(folderKey: string, name: string, mime: string, bytes: unknown): StoredFile;   // lỗi quyền → ném Error có chữ "permission"
}
export interface JobStore { get(jobId: string): UploadJob | null; put(j: UploadJob): void }

export const MAX_BYTES = 25e6;
export const ALLOWED_EXT = /\.(pdf|dwg|dxf|xlsx|xls|xlsm|docx|doc|pptx|jpg|jpeg|png|zip|rar|7z|csv|txt|ifc|rvt|nwd|nwc|skp)$/i;

export interface UploadInput { jobId: string; documentId: string; name: string; mime: string; size: number; bytes: unknown; clientSha256?: string }
export interface UploadDeps { svc: RegisterService; files: FileStore; jobs: JobStore; actor: Actor; now: () => Date; sha256: (bytes: unknown) => string; folderKey: string }

const safeName = (s: string) => s.replace(/[\\/:*?"<>|#%\u0000-\u001f]+/g, '_').replace(/\s+/g, ' ').trim();

export function runUpload(dep: UploadDeps, inp: UploadInput): { job: UploadJob; rev: unknown } {
  if (!inp.jobId || inp.jobId.length < 8) throw new BizError('IDEM_REQUIRED', 'Thiếu jobId');
  const t = () => dep.now().toISOString();
  let job = dep.jobs.get(inp.jobId);
  const save = (st: JobState, extra: Partial<UploadJob> = {}) => { job = { ...job!, ...extra, state: st, updatedAt: t() }; dep.jobs.put(job); };
  const fail = (at: JobState, e: unknown): never => {
    const er = e as BizError; save('FAILED', { failedAt: at, error: er.message });
    throw new BizError(er.code || 'UPLOAD', er.message + ' [job ' + inp.jobId + ' dừng ở ' + at + (job!.fileId ? ', file đã lên kho – bấm Thử lại sẽ không tải lại' : '') + ']');
  };
  const fileRef = (): FileRef => ({ name: job!.storedName!, url: job!.url!, source: 'DRIVE', sha256: job!.sha256, fileId: job!.fileId, size: job!.size, by: job!.by, at: job!.createdAt });
  const commit = () => {
    save('COMMITTING');
    let rev: unknown;
    try { rev = dep.svc.attach(dep.actor, 'up-' + inp.jobId, inp.documentId, fileRef()); } catch (e) { return fail('COMMITTING', e); }
    save('COMPLETED', { error: undefined, failedAt: undefined });
    return { job: job!, rev };
  };

  // chạy lại cùng jobId: đã xong → trả lại kết quả (attach idempotent theo khóa up-jobId), không sinh bản ghi/file mới
  if (job && job.state === 'COMPLETED') return { job, rev: dep.svc.attach(dep.actor, 'up-' + inp.jobId, inp.documentId, fileRef()) };
  if (job && job.documentId !== inp.documentId) throw new BizError('INVALID', 'jobId đã dùng cho hồ sơ khác');
  if (job && job.fileId && dep.files.exists(job.fileId)) return commit();        // file đã lên kho, lần trước lỗi lúc ghi sổ

  // 1. UPLOADING: nhận file + kiểm sơ bộ
  if (!job) { job = { jobId: inp.jobId, documentId: inp.documentId, name: inp.name, size: inp.size, mime: inp.mime, state: 'UPLOADING', by: dep.actor.email, createdAt: t(), updatedAt: t() }; dep.jobs.put(job); }
  else save('UPLOADING');
  try {
    if (!inp.name || !ALLOWED_EXT.test(inp.name)) throw new BizError('INVALID', 'Định dạng không cho phép: ' + inp.name + ' (nhận pdf, dwg, xlsx, docx, ảnh, zip…)');
    if (!(inp.size > 0)) throw new BizError('INVALID', 'File rỗng');
    if (inp.size > MAX_BYTES) throw new BizError('INVALID', 'File quá 25 MB – dán link OneDrive thay vì tải lên');
    if (inp.bytes == null) throw new BizError('INVALID', 'Mất dữ liệu file – chọn lại file rồi Thử lại');
  } catch (e) { return fail('UPLOADING', e); }

  // 2. VALIDATING: quyền, trạng thái hồ sơ, mã, hash, trùng nội dung / trùng tên
  let target: ReturnType<RegisterService['attachTarget']>;
  try {
    target = dep.svc.attachTarget(dep.actor, inp.documentId);
    save('VALIDATING', { docNo: target.doc.docNo, revNo: target.revNo });
    const sha = dep.sha256(inp.bytes);
    if (inp.clientSha256 && inp.clientSha256 !== sha) throw new BizError('INTEGRITY', 'File nhận được khác file gửi đi (hash lệch) – mạng chập chờn, bấm Thử lại');
    const ext = (inp.name.match(ALLOWED_EXT) || [''])[0].toLowerCase(), base = safeName(inp.name.slice(0, inp.name.length - ext.length)).slice(0, 60);
    const same = target.files.find(f => f.sha256 && f.sha256 === sha);
    if (same) {                                                     // cùng nội dung đã gắn → không tải thêm, không ghi thêm
      save('COMPLETED', { sha256: sha, storedName: same.name, url: same.url, fileId: same.fileId, note: 'Đã có file cùng nội dung trong rev này – không tải lại' });
      return { job: job!, rev: null };
    }
    let stored = target.doc.docNo + '_rev' + String(target.revNo).padStart(2, '0') + '_' + base + ext;
    if (target.files.some(f => f.name === stored)) stored = stored.slice(0, -ext.length) + '_' + sha.slice(0, 8) + ext;   // cùng tên khác nội dung
    save('VALIDATING', { sha256: sha, storedName: stored });
  } catch (e) { return fail('VALIDATING', e); }

  // 3. lên kho file (lưu fileId ngay để lỗi bước sau không tải lại)
  try {
    const f = dep.files.create(dep.folderKey, job!.storedName!, inp.mime || 'application/octet-stream', inp.bytes);
    save('COMMITTING', { fileId: f.id, url: f.url });
  } catch (e) {
    const m = (e as Error).message || '';
    return fail('UPLOADING', /permission|quyền|authoriz|access denied|insufficient/i.test(m) ? new BizError('DRIVE_PERM', 'Thiếu quyền Google Drive – mở lại trang và bấm Cho phép (' + m.slice(0, 80) + ')') : e);
  }
  // 4. COMMITTING → COMPLETED
  return commit();
}

/** Bộ nhớ cho test / xem trước. */
export class MemoryFileStore implements FileStore {
  files = new Map<string, { name: string; bytes: unknown }>(); failNext: string | null = null; n = 0;
  exists(id: string) { return this.files.has(id); }
  create(_k: string, name: string, _m: string, bytes: unknown) {
    if (this.failNext) { const m = this.failNext; this.failNext = null; throw new Error(m); }
    const id = 'f' + ++this.n; this.files.set(id, { name, bytes }); return { id, url: 'https://drive.google.com/file/d/' + id, name };
  }
}
export class MemoryJobStore implements JobStore {
  m = new Map<string, UploadJob>();
  get(id: string) { const j = this.m.get(id); return j ? { ...j } : null; }
  put(j: UploadJob) { this.m.set(j.jobId, { ...j }); }
}
