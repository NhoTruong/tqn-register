// Cổng lưu trữ. Nghiệp vụ chỉ nói chuyện qua interface này.
// Bản cài: MemoryRepo (test, xem trước), SheetsRepo (Apps Script), sau này PostgresRepo (ASP.NET Core).
import type { DocumentRec, EventRec, RevisionRec, DocType } from './types';

export interface Repo {
  begin?(): void;                                  // gọi đầu mỗi lệnh ghi (trong khóa): bỏ bộ đệm đọc để đọc dữ liệu mới nhất
  getDocument(id: string): DocumentRec | null;
  listDocuments(projectId: string, type?: DocType): DocumentRec[];
  saveDocument(d: DocumentRec): void;               // insert hoặc update theo documentId
  getRevisions(documentId: string): RevisionRec[];
  saveRevision(r: RevisionRec): void;               // khóa (documentId, revNo)
  appendEvent(e: EventRec): void;                   // chỉ thêm
  listEvents(documentId: string): EventRec[];
  getCounter(key: string): number;
  setCounter(key: string, value: number): void;
  getIdem(key: string): string | null;              // kết quả JSON đã lưu
  putIdem(key: string, resultJson: string): void;
}

/** Khóa toàn cục (LockService trong Apps Script). */
export interface Locker { withLock<T>(fn: () => T): T }

export class NoLock implements Locker { withLock<T>(fn: () => T): T { return fn(); } }

export class MemoryRepo implements Repo {
  docs = new Map<string, DocumentRec>();
  revs = new Map<string, RevisionRec>();
  events: EventRec[] = [];
  counters = new Map<string, number>();
  idem = new Map<string, string>();
  getDocument(id: string) { const d = this.docs.get(id); return d ? { ...d } : null; }
  listDocuments(projectId: string, type?: DocType) {
    return [...this.docs.values()].filter(d => d.projectId === projectId && (!type || d.docType === type)).map(d => ({ ...d }));
  }
  saveDocument(d: DocumentRec) { this.docs.set(d.documentId, { ...d }); }
  getRevisions(documentId: string) {
    return [...this.revs.values()].filter(r => r.documentId === documentId).sort((a, b) => a.revNo - b.revNo).map(r => ({ ...r, files: [...r.files] }));
  }
  saveRevision(r: RevisionRec) { this.revs.set(r.documentId + '#' + r.revNo, { ...r, files: [...r.files] }); }
  appendEvent(e: EventRec) { this.events.push({ ...e }); }
  listEvents(documentId: string) { return this.events.filter(e => e.documentId === documentId).map(e => ({ ...e })); }
  getCounter(key: string) { return this.counters.get(key) ?? 0; }
  setCounter(key: string, v: number) { this.counters.set(key, v); }
  getIdem(key: string) { return this.idem.get(key) ?? null; }
  putIdem(key: string, v: string) { this.idem.set(key, v); }
}
