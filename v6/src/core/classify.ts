// Document Classification – ĐỀ XUẤT loại hồ sơ từ tên file/tiêu đề email. Không tự lưu; người dùng xác nhận.
import type { DocType } from './types';

const RULES: [DocType, RegExp][] = [
  ['RFI', /\bRFI\b|y[eê]u\s*c[aầ]u\s*th[oô]ng\s*tin|request\s*for\s*information/i],
  ['RFA', /\bRFA\b|y[eê]u\s*c[aầ]u\s*ph[eê]\s*duy[eệ]t|request\s*for\s*approval/i],
  ['RFC', /\bRFC\b|y[eê]u\s*c[aầ]u\s*thay\s*đ[oổ]i|request\s*for\s*change/i],
  ['SDS', /\bSDS\b|shop\s*drawing|b[aả]n\s*v[eẽ]\s*(thi\s*c[oô]ng|chi\s*ti[eế]t)/i],
  ['MSS', /\bM[0O]?SS?\b|method\s*statement|bi[eệ]n\s*ph[aá]p\s*thi\s*c[oô]ng|BPTC/i],
  ['MAS', /\bMAS\b|material\s*(approval|submittal)|ph[eê]\s*duy[eệ]t\s*v[aậ]t\s*li[eệ]u/i],
];

export interface Suggestion { type: DocType | null; confidence: number; reason: string; rev: number | null }

export function classify(name: string): Suggestion {
  const s = String(name || '');
  const rev = (() => { const m = s.match(/rev\s*[-_.]?\s*(\d{1,2})/i) || s.match(/\bR(\d{2})\b/); return m ? +m[1] : null; })();
  const hits = RULES.filter(([, re]) => re.test(s));
  if (hits.length === 1) return { type: hits[0][0], confidence: /\b[A-Z]{3}\b/.test(s) ? 0.95 : 0.75, reason: 'Khớp mẫu ' + hits[0][0], rev };
  if (hits.length > 1) return { type: hits[0][0], confidence: 0.4, reason: 'Khớp nhiều loại: ' + hits.map(h => h[0]).join(', ') + ' – cần người chọn', rev };
  return { type: null, confidence: 0, reason: 'Không nhận ra loại hồ sơ', rev };
}
