// Điểm vào Apps Script. Hàm toàn cục = API cho Web App (google.script.run). Chỉ adapter, không có nghiệp vụ.
import { RegisterService, BizError } from '../core/service';
import { coastalConfig } from '../core/config';
import { classify } from '../core/classify';
import { SheetsRepo, gasLock } from './sheetsRepo';
import type { Actor, DocType, FileRef, ProjectConfig } from '../core/types';

declare const global: Record<string, unknown>;
declare const __BUILD__: { commit: string; builtAt: string; env: string; version: string };

function props() { return PropertiesService.getScriptProperties(); }
function cfg(): ProjectConfig {
  const raw = props().getProperty('TQN6_MEMBERS');                // {"email":["ROLE"]} – không chứa mật khẩu/khóa
  return coastalConfig(raw ? JSON.parse(raw) : {});
}
function store(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const id = props().getProperty('TQN6_STORE_ID');                 // file lưu theo môi trường; thiếu thì dùng file gắn script
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
global.apiClassify = (name: string) => classify(name);
global.apiConfig = () => safe(() => { const c = cfg(); return { types: c.types, code: c.code }; });

/** Kiểm phiên bản đang chạy: so với RELEASES.md và git commit. */
global.tqn6Version = () => { Logger.log(JSON.stringify(__BUILD__)); return __BUILD__; };
