// Điểm vào Apps Script. Hàm toàn cục = API cho Web App (google.script.run). Chỉ adapter, không có nghiệp vụ.
import { RegisterService, BizError } from '../core/service';
import { coastalConfig } from '../core/config';
import { classify } from '../core/classify';
import { SheetsRepo, gasLock } from './sheetsRepo';
import { syncRow } from './sheetSync';
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
  try { return fn(); } catch (e) { const er = e as BizError; return { error: er.message, code: er.code || 'ERROR' }; }
}

global.doGet = () => HtmlService.createHtmlOutputFromFile('dist/index.html').setTitle('TQN Register').addMetaTag('viewport', 'width=device-width, initial-scale=1');

global.apiWhoAmI = () => safe(() => { const a = me(); return { email: a.email, roles: svc().rolesOf(a.email), build: __BUILD__ }; });
global.apiList = (type?: DocType) => safe(() => svc().list(type));
global.apiGet360 = (id: string) => safe(() => svc().get360(id));
global.apiCreate = (idem: string, input: { type: DocType; title: string; files?: FileRef[] }) => safe(() => svc().create(me(), idem, input));
global.apiAttach = (idem: string, id: string, f: FileRef) => safe(() => svc().attach(me(), idem, id, f));
global.apiSubmit = (idem: string, id: string) => safe(() => svc().submit(me(), idem, id));
global.apiReply = (idem: string, id: string, code: string, note: string) => safe(() => svc().reply(me(), idem, id, code, note));
global.apiCloseForwarded = (idem: string, id: string, note: string) => safe(() => svc().closeForwarded(me(), idem, id, note));
/** Một lần gọi cho mỗi thao tác: ghi + trả luôn danh sách và hồ sơ 360 (bỏ 2 lần gọi lại). Có ghi thời gian. */
global.apiDo = (op: string, idem: string, a: any) => safe(() => {
  const t0 = Date.now(), s = svc(), who = me();
  let r: any;
  if (op === 'create') r = s.create(who, idem, a);
  else if (op === 'attach') r = s.attach(who, idem, a.id, a.file);
  else if (op === 'submit') r = s.submit(who, idem, a.id);
  else if (op === 'reply') r = s.reply(who, idem, a.id, a.code, a.note || '');
  else if (op === 'closeForwarded') r = s.closeForwarded(who, idem, a.id, a.note || '');
  else throw new BizError('INVALID', 'Thao tác lạ: ' + op);
  const id = (r && r.documentId) || a.id, t1 = Date.now();
  const v360 = id ? s.get360(id) : null;
  let soRFI = '';
  try { if (v360) soRFI = syncRow(store(), 'RFI', v360.doc, v360.revisions, who.email.split('@')[0]); } catch (e) { soRFI = 'Lỗi ghi sổ RFI: ' + (e as Error).message; }
  const out = { result: r, list: s.list('RFI'), v360, soRFI, ms: { ghi: t1 - t0, doc: Date.now() - t1 } };
  console.log(JSON.stringify({ op, ms: out.ms, ok: true }));
  return out;
});
/** Tải file thật lên Drive (thư mục riêng của môi trường), tên theo mã + rev, trả FileRef để gắn vào hồ sơ. */
global.apiUpload = (idem: string, id: string, name: string, mime: string, b64: string) => safe(() => {
  const s = svc(), d = s.get360(id).doc;
  if (b64.length > 34e6) throw new BizError('INVALID', 'File quá 25 MB – dán link OneDrive thay vì tải lên');
  const p = props(), key = 'TQN6_FOLDER_' + __BUILD__.env; let fid = p.getProperty(key), folder: GoogleAppsScript.Drive.Folder;
  try { folder = DriveApp.getFolderById(String(fid)); } catch (e) { folder = DriveApp.createFolder('TQN Register V6 – ' + __BUILD__.env + ' – file hồ sơ'); p.setProperty(key, folder.getId()); }
  const rev = d.status === 'DRAFT' ? d.currentRev : d.currentRev + 1, ext = (name.match(/\.[a-z0-9]+$/i) || [''])[0];
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', d.docNo + '_rev' + ('0' + rev).slice(-2) + '_' + name.replace(ext, '').slice(0, 60) + ext));
  return (global.apiDo as any)('attach', idem, { id, file: { name: file.getName(), url: file.getUrl(), source: 'DRIVE' } });
});
global.apiBoot = (id?: string) => safe(() => {               // mở web: 1 lần gọi lấy đủ người dùng + danh sách + hồ sơ đầu
  const t0 = Date.now(), s = svc(), a = me(), list = s.list('RFI'), pick = id || (list[0] && list[0].documentId);
  const out = { who: { email: a.email, roles: s.rolesOf(a.email), build: __BUILD__ }, list, v360: pick ? s.get360(pick) : null, ms: Date.now() - t0 };
  console.log(JSON.stringify({ op: 'boot', ms: out.ms }));
  return out;
});
global.apiClassify = (name: string) => classify(name);
global.apiConfig = () => safe(() => { const c = cfg(); return { types: c.types, code: c.code }; });

/** Kiểm phiên bản đang chạy: so với RELEASES.md và git commit. */
global.tqn6Version = () => { Logger.log(JSON.stringify(__BUILD__)); return __BUILD__; };
