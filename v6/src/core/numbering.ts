// Numbering Engine – cấp mã theo mẫu của dự án. Bộ đếm theo (dự án, loại); gọi trong khóa (LockService ở adapter).
import type { DocType, ProjectConfig } from './types';

export interface NumberCtx { cfg: ProjectConfig; type: DocType; date: Date; seq: number }

function pad(n: number, w: number) { const s = String(n); return s.length >= w ? s : '0'.repeat(w - s.length) + s; }

export function formatNo(pattern: string, c: NumberCtx): string {
  const y = c.date.getFullYear() % 100, m = c.date.getMonth() + 1, d = c.date.getDate();
  return pattern.replace(/\{(\w+)(?::(\d+))?\}/g, (_, k: string, w?: string) => {
    switch (k) {
      case 'yymmdd': return pad(y, 2) + pad(m, 2) + pad(d, 2);
      case 'yyyy': return String(c.date.getFullYear());
      case 'project': return c.cfg.code;
      case 'contractor': return c.cfg.contractor;
      case 'type': return c.type;
      case 'seq': return pad(c.seq, w ? +w : 3);
      default: throw new Error('Mẫu mã có ô lạ: {' + k + '}');
    }
  });
}

/** Phần so khớp chống trùng: bỏ ngày đầu mã để "261111-…-001" và "261116-…-001" coi là cùng số. */
export function noKey(docNo: string): string {
  return docNo.trim().toUpperCase().replace(/^\d{6}-/, '').replace(/[-\s_]*REV\s*\d+$/i, '');
}

/**
 * Cấp số kế tiếp: lấy max(bộ đếm, số lớn nhất đang có) + 1, rồi bỏ qua mọi số đã dùng.
 * existing: các mã đã có cùng dự án + loại (đọc 1 lần, batch).
 */
export function nextNumber(cfg: ProjectConfig, type: DocType, date: Date, counter: number, existing: string[]): { docNo: string; seq: number } {
  const t = cfg.types[type];
  if (!t) throw new Error('Dự án chưa cấu hình loại ' + type);
  const used = new Set(existing.map(noKey));
  let seq = counter + 1;
  for (let guard = 0; guard < 10000; guard++, seq++) {
    const docNo = formatNo(t.pattern, { cfg, type, date, seq });
    if (!used.has(noKey(docNo))) return { docNo, seq };
  }
  throw new Error('Không cấp được số mới cho ' + type);
}
