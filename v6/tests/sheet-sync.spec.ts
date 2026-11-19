// Ghi thẳng vào sổ RFI: dò cột theo tiêu đề, cập nhật đúng dòng theo Mã, lần đầu lấy dòng trống.
import { describe, it, expect } from 'vitest';
import { findCols, syncRow } from '../src/backend/sheetSync';

const H3 = ['STT', 'Mã hồ sơ', 'Nội dung', 'Người thực hiện', 'Tình trạng', 'Rev', 'Ngày trình duyệt', 'Nhà thầu đệ trình', '', 'Chủ đầu tư', '', '', 'CĐT trả lời · theo phiếu'];
const H4 = ['', '', '', '', '', '', '', 'Kế hoạch', 'Thực tế', 'Kế hoạch', 'Thực tế', 'Vướng mắc', ''];
function fake() {
  const g: any[][] = [[], [], H3, H4, [], ['I', '', 'PHẦN XÂY DỰNG'], ['1', '', ''], ['2', '', '']];
  const links: Record<string, string> = {};
  (globalThis as any).SpreadsheetApp = { newRichTextValue: () => { let t = '', u = ''; const b: any = { setText: (x: string) => (t = x, b), setLinkUrl: (x: string) => (u = x, b), build: () => ({ t, u }) }; return b; } };
  const sh: any = {
    getLastColumn: () => H3.length, getLastRow: () => g.length,
    getRange: (r: number, c: number, nr = 1, nc = 1) => ({
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (g[r - 1 + i] || [])[c - 1 + j] ?? '')),
      setValue: (v: any) => { (g[r - 1] ||= [])[c - 1] = v; },
      setValues: (v: any[][]) => { v[0].forEach((x, j) => { (g[r - 1] ||= [])[c - 1 + j] = x; }); },
      setRichTextValue: (v: any) => { (g[r - 1] ||= [])[c - 1] = v.t; links[r + ',' + c] = v.u; },
    }),
  };
  return { g, links, ss: { getSheetByName: (n: string) => (n === 'RFI' ? sh : null) } as any };
}
const doc: any = { docNo: '261119-COA-TBC-RFI-001', title: 'Hỏi gạch', status: 'SUBMITTED', currentRev: 0 };
const revs: any = [{ revNo: 0, submittedAt: '2026-11-19T08:00:00Z', files: [{ name: 'a.pdf', url: 'https://drive.google.com/a' }] }];

describe('Ghi sổ RFI', () => {
  it('dò đúng cột theo tiêu đề', () => {
    const c = findCols([H3, H4, []]);
    expect(c).toMatchObject({ code: 1, subject: 2, owner: 3, status: 4, rev: 5, submitted: 6, ntActual: 8, cdtActual: 10, cdtNote: 11, pick: 12 });
  });
  it('lần đầu: dòng trống đầu tiên (bỏ dòng mục I); lần sau: cùng dòng theo Mã', () => {
    const f = fake();
    expect(syncRow(f.ss, 'RFI', doc, revs, 'nho')).toBe('RFI!7');
    expect(f.g[6].slice(1, 6)).toEqual(['261119-COA-TBC-RFI-001', 'Hỏi gạch', 'nho', 'Đã trình duyệt', '00']);
    expect(f.links['7,2']).toBe('https://drive.google.com/a');
    const closed = { ...doc, status: 'CLOSED' }, r2 = [{ ...revs[0], repliedAt: '2026-11-20T08:00:00Z', replyCode: '1', replyText: 'Xem trả lời bên dưới', replyNote: 'OK' }];
    expect(syncRow(f.ss, 'RFI', closed, r2, 'nho')).toBe('RFI!7');
    expect(f.g[6][4]).toBe('Đã phản hồi'); expect(f.g[6][11]).toBe('OK'); expect(f.g[6][12]).toBe('1 · Xem trả lời bên dưới');
    expect(f.g[7][1] ?? '').toBe('');
  });
  it('không có tab → báo, không lỗi', () => { expect(syncRow(fake().ss, 'RFX', doc, revs, 'x')).toMatch(/Không có tab/); });
});
