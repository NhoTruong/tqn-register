// Adapter Google Sheets cho Repo. Tab TQN6_* trong FILE ĐÍCH theo môi trường (DEV/STAGING/PROD) – không ghi vào sổ gốc.
// Đọc theo lô (1 getValues mỗi tab mỗi request), ghi theo dòng xác định – không lặp từng ô.
import type { Repo } from '../core/repo';
import type { DocType, DocumentRec, EventRec, RevisionRec } from '../core/types';

const TABS = {
  DOCS: ['documentId', 'projectId', 'docType', 'docNo', 'title', 'status', 'currentRev', 'owner', 'holder', 'dueAt', 'createdAt', 'updatedAt'],
  REVS: ['documentId', 'revNo', 'submittedAt', 'repliedAt', 'replyCode', 'replyText', 'replyNote', 'filesJson'],
  EVENTS: ['eventId', 'at', 'actor', 'actorKind', 'action', 'documentId', 'revNo', 'idemKey', 'detail', 'ok'],
  KV: ['key', 'value', 'updatedAt'],
} as const;
type TabName = keyof typeof TABS;

export class SheetsRepo implements Repo {
  private cache: Partial<Record<TabName, unknown[][]>> = {};
  constructor(private ss: GoogleAppsScript.Spreadsheet.Spreadsheet, private prefix = 'TQN6_') {}

  private sh(t: TabName) {
    const name = this.prefix + t;
    let s = this.ss.getSheetByName(name);
    if (!s) { s = this.ss.insertSheet(name); s.getRange(1, 1, s.getMaxRows(), TABS[t].length).setNumberFormat('@'); s.getRange(1, 1, 1, TABS[t].length).setValues([[...TABS[t]]]).setFontWeight('bold'); s.setFrozenRows(1); }
    return s;
  }
  private rows(t: TabName): unknown[][] {
    if (!this.cache[t]) { const s = this.sh(t), n = s.getLastRow() - 1; this.cache[t] = n > 0 ? s.getRange(2, 1, n, TABS[t].length).getValues() : []; }
    return this.cache[t]!;
  }
  private upsert(t: TabName, match: (r: unknown[]) => boolean, row: unknown[]) {
    const all = this.rows(t), i = all.findIndex(match), s = this.sh(t);
    if (i >= 0) { s.getRange(i + 2, 1, 1, row.length).setValues([row]); all[i] = row; }
    else { s.getRange(all.length + 2, 1, 1, row.length).setValues([row]); all.push(row); }
  }
  private obj<T>(t: TabName, r: unknown[]): T {
    // Sheets tự đổi chữ ISO thành ngày → trả lại chuỗi ISO (google.script.run không trả được Date lồng trong object)
    const o: Record<string, unknown> = {};
    TABS[t].forEach((k, i) => { const v = r[i]; o[k] = v === '' ? null : v instanceof Date ? v.toISOString() : v; });
    return o as T;
  }

  getDocument(id: string) { const r = this.rows('DOCS').find(x => x[0] === id); return r ? this.doc(r) : null; }
  private doc(r: unknown[]): DocumentRec { const d = this.obj<DocumentRec>('DOCS', r); d.currentRev = Number(d.currentRev); d.docNo = String(d.docNo); return d; }
  listDocuments(projectId: string, type?: DocType) { return this.rows('DOCS').filter(x => x[1] === projectId && (!type || x[2] === type)).map(r => this.doc(r)); }
  saveDocument(d: DocumentRec) { this.upsert('DOCS', x => x[0] === d.documentId, TABS.DOCS.map(k => (d as unknown as Record<string, unknown>)[k] ?? '')); }
  getRevisions(documentId: string): RevisionRec[] {
    return this.rows('REVS').filter(x => x[0] === documentId).map(r => {
      const o = this.obj<Record<string, unknown>>('REVS', r);
      return { documentId, revNo: Number(o.revNo), submittedAt: o.submittedAt as string | null, repliedAt: o.repliedAt as string | null, replyCode: o.replyCode == null ? null : String(o.replyCode), replyText: o.replyText as string | null, replyNote: o.replyNote as string | null, files: JSON.parse(String(o.filesJson || '[]')) };
    }).sort((a, b) => a.revNo - b.revNo);
  }
  saveRevision(r: RevisionRec) {
    this.upsert('REVS', x => x[0] === r.documentId && Number(x[1]) === r.revNo, [r.documentId, r.revNo, r.submittedAt ?? '', r.repliedAt ?? '', r.replyCode ?? '', r.replyText ?? '', r.replyNote ?? '', JSON.stringify(r.files)]);
  }
  appendEvent(e: EventRec) { const s = this.sh('EVENTS'); s.appendRow(TABS.EVENTS.map(k => (e as unknown as Record<string, unknown>)[k] ?? '')); this.cache.EVENTS = undefined; }
  listEvents(documentId: string) { return this.rows('EVENTS').filter(x => x[5] === documentId).map(r => { const e = this.obj<EventRec>('EVENTS', r); e.ok = String(e.ok) === 'true'; e.revNo = e.revNo == null ? null : Number(e.revNo); return e; }); }
  private kv(key: string) { const r = this.rows('KV').find(x => x[0] === key); return r ? String(r[1]) : null; }
  private setKv(key: string, v: string) { this.upsert('KV', x => x[0] === key, [key, v, new Date().toISOString()]); }
  getCounter(key: string) { return Number(this.kv('cnt:' + key) || 0); }
  setCounter(key: string, v: number) { this.setKv('cnt:' + key, String(v)); }
  getIdem(key: string) { return this.kv('idem:' + key); }
  putIdem(key: string, v: string) { this.setKv('idem:' + key, v); }
}

/** Khóa tài liệu: hai người bấm cùng lúc → người sau chờ tối đa 20 s, không chạy song song. */
export const gasLock = {
  withLock<T>(fn: () => T): T {
    const l = LockService.getDocumentLock() || LockService.getScriptLock();
    if (!l.tryLock(20000)) throw new Error('BUSY: Hệ thống đang bận ghi, thử lại sau vài giây');
    try { return fn(); } finally { SpreadsheetApp.flush(); l.releaseLock(); }
  },
};
