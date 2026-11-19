// Ghi thẳng vào sổ RFI quen thuộc (cùng file với TQN6_*): mỗi hồ sơ 1 dòng, tìm theo Mã; chưa có thì lấy dòng trống đầu tiên.
// Dò cột theo tiêu đề (dòng 3–5) nên sổ dự án khác chỉ cần cùng tên cột. Lỗi ở đây không làm hỏng thao tác chính (ghi log).
import type { DocumentRec, RevisionRec } from '../core/types';

const ST: Record<string, string> = { DRAFT: 'Đang thực hiện', SUBMITTED: 'Đã trình duyệt', FORWARDED: 'Đã trình duyệt', NEED_INFO: 'Trả về', RESUBMIT_REQUIRED: 'Trả về', CLOSED: 'Đã phản hồi', REJECTED: 'Không duyệt' };
const norm = (s: unknown) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();

export function findCols(hdr: unknown[][]) {
  const c: Record<string, number> = {}; let cdtStart = -1, cdtEnd = -1, ntStart = -1, ntEnd = -1;
  const top = hdr[0] || [];
  top.forEach((v, i) => {
    const t = norm(v);
    if (t.startsWith('chu dau tu')) { cdtStart = i; cdtEnd = i + 1; while (cdtEnd < top.length && !norm(top[cdtEnd])) cdtEnd++; }
    if (t.startsWith('nha thau de trinh')) { ntStart = i; ntEnd = i + 1; while (ntEnd < top.length && !norm(top[ntEnd])) ntEnd++; }
  });
  for (let r = 0; r < hdr.length; r++) hdr[r].forEach((v, i) => {
    const t = norm(v);
    if (t === 'ma ho so' && c.code == null) c.code = i;
    else if (t.startsWith('noi dung') && c.subject == null) c.subject = i;
    else if (t.startsWith('nguoi thuc hien') && c.owner == null) c.owner = i;
    else if (t.startsWith('tinh trang') && c.status == null) c.status = i;
    else if (t === 'rev' && c.rev == null) c.rev = i;
    else if (t.startsWith('ngay trinh duyet') && c.submitted == null) c.submitted = i;
    else if (t === 'thuc te' && ntStart >= 0 && i >= ntStart && i < ntEnd && c.ntActual == null) c.ntActual = i;
    else if (t === 'thuc te' && cdtStart >= 0 && i >= cdtStart && i < cdtEnd && c.cdtActual == null) c.cdtActual = i;
    else if (t.startsWith('vuong mac') && cdtStart >= 0 && i >= cdtStart && i < cdtEnd && c.cdtNote == null) c.cdtNote = i;
    else if (t.includes('cdt tra loi') && t.includes('theo phieu')) c.pick = i;
  });
  return c;
}

export function syncRow(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, tab: string, d: DocumentRec, revs: RevisionRec[], ownerName: string): string {
  const sh = ss.getSheetByName(tab); if (!sh) return 'Không có tab ' + tab;
  const lastC = sh.getLastColumn(), hdr = sh.getRange(3, 1, 3, lastC).getValues(), c = findCols(hdr);
  if (c.code == null || c.subject == null) return 'Không dò được cột Mã/Nội dung ở tab ' + tab;
  const first = 6, n = Math.max(sh.getLastRow() - first + 1, 1), data = sh.getRange(first, 1, n, lastC).getValues();
  let r = data.findIndex(x => String(x[c.code]).trim() === d.docNo);
  if (r < 0) r = data.findIndex(x => !String(x[c.code]).trim() && !String(x[c.subject]).trim() && String(x[0]).trim() !== '' && !/^[IVX]+$/.test(String(x[0]).trim()));
  if (r < 0) r = data.length;
  const row = first + r, cur = revs.find(x => x.revNo === d.currentRev), set = (col: number | undefined, v: unknown) => { if (col != null) sh.getRange(row, col + 1).setValue(v); };
  set(c.code, d.docNo); set(c.subject, d.title); set(c.owner, ownerName);
  try { set(c.status, ST[d.status] || d.status); } catch (e) { /* ô có danh sách chọn không chứa chữ này */ }
  set(c.rev, ('0' + d.currentRev).slice(-2));
  if (cur && cur.submittedAt) { const t = new Date(cur.submittedAt); set(c.submitted, t); set(c.ntActual, t); }
  if (cur && cur.repliedAt) { set(c.cdtActual, new Date(cur.repliedAt)); set(c.cdtNote, cur.replyNote || ''); try { set(c.pick, cur.replyCode + ' · ' + cur.replyText); } catch (e) { } }
  const f = cur && cur.files[0]; if (f) { try { sh.getRange(row, c.code + 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(d.docNo).setLinkUrl(f.url).build()); } catch (e) { } }
  return tab + '!' + row;
}
