// Điểm vào Apps Script. Hàm toàn cục = API cho Web App (google.script.run). Chỉ adapter, không có nghiệp vụ.
import { RegisterService, BizError } from '../core/service';
import { coastalConfig } from '../core/config';
import { classify } from '../core/classify';
import { SheetsRepo, gasLock } from './sheetsRepo';
import { syncRow } from './sheetSync';
import { P } from './perf';
import { runUpload, type FileStore, type JobStore } from '../core/upload';
import type { Actor, DocType, FileRef, ProjectConfig } from '../core/types';

declare const global: Record<string, unknown>;
declare const __BUILD__: { commit: string; builtAt: string; env: string; version: string };
declare const __ENVCFG__: { storeId: string; members: Record<string, string[]> };

function props() { return PropertiesService.getScriptProperties(); }
function cfg(): ProjectConfig {
  const raw = props().getProperty('TQN6_MEMBERS');                // {"email":["ROLE"]} – không chứa mật khẩu/khóa
  return coastalConfig(raw ? JSON.parse(raw) : (__ENVCFG__.members as ProjectConfig['members']));
}
function store(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const id = props().getProperty('TQN6_STORE_ID') || __ENVCFG__.storeId;   // file lưu theo môi trường (DEV/STAGING/PROD)
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}
function me(): Actor { return { email: (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail()).toLowerCase(), kind: 'HUMAN' }; }
function svc() { return new RegisterService(cfg(), new SheetsRepo(store()), gasLock, { now: () => new Date() }); }
function safe<T>(fn: () => T): T | { error: string; code: string } {
  try { return fn(); } catch (e) { const er = e as BizError; return { error: er.message, code: er.code || 'ERROR', ms: P.out() } as any; }
}

global.doGet = () => HtmlService.createHtmlOutputFromFile('dist/index.html').setTitle('TQN Register').addMetaTag('viewport', 'width=device-width, initial-scale=1');

const T = <F extends (...a: any[]) => any>(f: F) => ((...a: any[]) => { P.reset(); return f(...a); }) as F;

// ---------- phiên bản dữ liệu + bộ đệm máy chủ ----------
// dv đổi sau mỗi lệnh ghi thành công. Khóa cache gồm dự án + dv → không bao giờ trả dữ liệu cũ sau khi có người ghi.
// Sửa tay thẳng vào tab TQN6_* (không qua web) thì cache tự hết hạn sau TTL.
const TTL = 300;
function dv(): string { return props().getProperty('TQN6_DV') || '0'; }
function bumpDv(): string { const v = Date.now().toString(36) + Math.random().toString(36).slice(2, 6); props().setProperty('TQN6_DV', v); return v; }
function cache() { try { return CacheService.getScriptCache(); } catch (e) { return null; } }
function cGet<T>(k: string): T | null { const c = cache(); if (!c) return null; const v = P.time('cache', () => c.get(k)); return v ? JSON.parse(v) as T : null; }
function cPut(k: string, o: unknown) { const c = cache(), j = JSON.stringify(o); if (c && j.length < 95000) P.time('cache', () => c.put(k, j, TTL)); }
const proj = () => cfg().projectId;
function listCached(s: RegisterService, ver: string) {
  const k = 'L:' + proj() + ':' + ver, hit = cGet<any[]>(k); if (hit) return hit;
  const l = s.list('RFI'); cPut(k, l); return l;
}
function v360Cached(s: RegisterService, id: string, ver: string) {
  const k = 'V:' + id + ':' + ver, hit = cGet<any>(k); if (hit) return hit;
  const v = s.get360(id); cPut(k, v); return v;
}
function syncSo(s: RegisterService, v: any, who: Actor): string {
  try { return P.time('ghiSo', () => syncRow(store(), v.doc.docType, v.doc, v.revisions, who.email.split('@')[0])); }
  catch (e) { return 'Lỗi ghi sổ ' + v.doc.docType + ': ' + (e as Error).message; }
}
/** Sau lệnh ghi: đổi dv, ghi sổ, trả đủ dữ liệu cho web (1 lần gọi), đặt sẵn cache cho người khác mở. */
function afterWrite(s: RegisterService, who: Actor, id: string | null, extra: object) {
  const ver = bumpDv(), list = s.list('RFI'), v360 = id ? s.get360(id) : null;
  cPut('L:' + proj() + ':' + ver, list); if (v360) cPut('V:' + id + ':' + ver, v360);
  const so = v360 ? syncSo(s, v360, who) : '';
  const out = { ...extra, list, v360, soRFI: so, dv: ver, ms: P.out() };
  console.log(JSON.stringify({ op: (extra as any).op, ms: out.ms }));
  return out;
}

global.apiWhoAmI = T(() => safe(() => { const a = me(); return { email: a.email, roles: svc().rolesOf(a.email), build: __BUILD__ }; }));
global.apiList = T((type?: DocType) => safe(() => svc().list(type)));
global.apiGet360 = T((id: string) => safe(() => { const ver = dv(); return { ...v360Cached(svc(), id, ver), dv: ver, ms: P.out() }; }));
/** Hỏi nhẹ: chỉ đọc 1 thuộc tính (không mở Sheets) – web gọi định kỳ để biết có ai vừa sửa. */
global.apiPeek = () => ({ dv: dv() });
global.apiCreate = T((idem: string, input: { type: DocType; title: string; files?: FileRef[] }) => safe(() => svc().create(me(), idem, input)));
global.apiAttach = T((idem: string, id: string, f: FileRef) => safe(() => svc().attach(me(), idem, id, f)));
global.apiSubmit = T((idem: string, id: string, ev?: number) => safe(() => svc().submit(me(), idem, id, ev)));
global.apiReply = T((idem: string, id: string, code: string, note: string, ev?: number) => safe(() => svc().reply(me(), idem, id, code, note, ev)));
global.apiCloseForwarded = T((idem: string, id: string, note: string, ev?: number) => safe(() => svc().closeForwarded(me(), idem, id, note, ev)));
/** Một lần gọi cho mỗi thao tác: ghi (khóa + chống trùng + kiểm version) + trả danh sách, hồ sơ 360, dv, thời gian từng chặng. */
global.apiDo = T((op: string, idem: string, a: any) => safe(() => {
  const s = svc(), who = me();
  let r: any;
  P.time('nghiepVu', () => {
    if (op === 'create') r = s.create(who, idem, a);
    else if (op === 'attach') r = s.attach(who, idem, a.id, a.file);
    else if (op === 'submit') r = s.submit(who, idem, a.id, a.ev);
    else if (op === 'reply') r = s.reply(who, idem, a.id, a.code, a.note || '', a.ev);
    else if (op === 'closeForwarded') r = s.closeForwarded(who, idem, a.id, a.note || '', a.ev);
    else throw new BizError('INVALID', 'Thao tác lạ: ' + op);
  });
  return afterWrite(s, who, (r && r.documentId) || a.id || null, { op, result: r });
}));

// ---------- Upload Drive theo job ----------
class DriveStore implements FileStore {
  exists(id: string) { try { return P.time('drive', () => !DriveApp.getFileById(id).isTrashed()); } catch (e) { return false; } }
  create(folderKey: string, name: string, mime: string, bytes: unknown) {
    return P.time('drive', () => {
      const p = props(), k = 'TQN6_FOLDER_' + folderKey; let folder: GoogleAppsScript.Drive.Folder;
      const fid = p.getProperty(k);
      try { folder = DriveApp.getFolderById(String(fid)); } catch (e) { folder = DriveApp.createFolder('TQN Register – ' + folderKey + ' – file hồ sơ'); p.setProperty(k, folder.getId()); }
      const f = folder.createFile(Utilities.newBlob(bytes as number[], mime, name));   // thư mục riêng tư của chủ sở hữu, không chia sẻ công khai
      return { id: f.getId(), url: f.getUrl(), name: f.getName() };
    });
  }
}
const propsJobs: JobStore = {
  get: (id) => { const v = props().getProperty('TQN6_JOB_' + id); return v ? JSON.parse(v) : null; },
  put: (j) => {
    props().setProperty('TQN6_JOB_' + j.jobId, JSON.stringify(j));
    if (j.state === 'COMPLETED' && Math.random() < 0.05) {          // dọn job đã xong quá 7 ngày
      const all = props().getProperties(), old = Date.now() - 7 * 864e5;
      Object.keys(all).filter(k => k.indexOf('TQN6_JOB_') === 0).forEach(k => { try { const x = JSON.parse(all[k]); if (x.state === 'COMPLETED' && Date.parse(x.updatedAt) < old) props().deleteProperty(k); } catch (e) { } });
    }
  },
};
const hex = (b: number[]) => b.map(x => ('0' + ((x + 256) % 256).toString(16)).slice(-2)).join('');
global.apiUpload = T((jobId: string, id: string, name: string, mime: string, b64: string, clientSha?: string) => {
  try {
    const s = svc(), who = me(), bytes = b64 ? Utilities.base64Decode(b64) : null;
    const r = runUpload({ svc: s, files: new DriveStore(), jobs: propsJobs, actor: who, now: () => new Date(), folderKey: cfg().code + '_' + __BUILD__.env,
      sha256: (b) => hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b as number[])) },
      { jobId, documentId: id, name, mime, size: bytes ? bytes.length : 0, bytes, clientSha256: clientSha });
    return afterWrite(s, who, id, { op: 'upload', job: r.job });
  } catch (e) { const er = e as BizError; return { error: er.message, code: er.code || 'ERROR', job: propsJobs.get(jobId), ms: P.out() }; }
});
global.apiJob = (jobId: string) => propsJobs.get(jobId);
global.apiBoot = T((id?: string) => safe(() => {               // mở web: 1 lần gọi, ưu tiên cache theo dv
  const s = svc(), a = me(), ver = dv(), list = listCached(s, ver), pick = id || (list[0] && list[0].documentId);
  const out = { who: { email: a.email, roles: s.rolesOf(a.email), build: __BUILD__ }, list, v360: pick ? v360Cached(s, pick, ver) : null, dv: ver, ms: P.out() };
  console.log(JSON.stringify({ op: 'boot', ms: out.ms }));
  return out;
}));
global.apiClassify = (name: string) => classify(name);
global.apiConfig = () => safe(() => { const c = cfg(); return { types: c.types, code: c.code }; });

/** Kiểm phiên bản đang chạy: so với RELEASES.md và git commit. */
global.tqn6Version = () => { Logger.log(JSON.stringify(__BUILD__)); return __BUILD__; };
