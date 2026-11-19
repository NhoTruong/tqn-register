"""Xuat Register .xlsx -> JSON (gia tri da tinh + cong thuc) de test Apps Script bang Node. Chi doc."""
import sys, json, datetime, openpyxl
src, out, live = sys.argv[1], sys.argv[2], (sys.argv[3] == 'live' if len(sys.argv) > 3 else False)
wv = openpyxl.load_workbook(src, data_only=True); wf = openpyxl.load_workbook(src)
def enc(v):
    if isinstance(v, datetime.datetime): return {'__d': v.isoformat()}
    if isinstance(v, datetime.date): return {'__d': datetime.datetime(v.year, v.month, v.day).isoformat()}
    return v
fx = {}
for ws in wv.worksheets:
    f = wf[ws.title]; R, C = ws.max_row, min(ws.max_column, 40)
    vals = [[enc(ws.cell(r, c).value) for c in range(1, C + 1)] for r in range(1, R + 1)]
    forms = [[(f.cell(r, c).value if isinstance(f.cell(r, c).value, str) and f.cell(r, c).value.startswith('=') else '') for c in range(1, C + 1)] for r in range(1, R + 1)]
    fx[ws.title] = {'values': vals, 'formulas': forms}
if live:   # dan ten file "Link hồ sơ" doc duoc tren ban live 19/11 09:45 (cot P, dong 23-30)
    P = ['260611-COA-TBC-RFC-001 Rev01.pdf', '261116-COA-TBC-RFI-001 rev01.pdf', '261116-COA-TBC-RFI-002 rev01.pdf',
         '261116-COA-TBC-RFI-003 rev01.pdf', '261116-COA-TBC-RFA-COS-002-rev02.pdf', '261116-COA-TBC-RFA-COS-001-rev02.pdf',
         '261116-COA-TBC-RFC-002 rev01.pdf', '261118-COA-TBC-RFC-003 rev01-Yêu cầu thay đổi bề dày thép tấm rọ cọc.pdf']
    s = fx['Sổ trình duyệt']['values']
    for i, p in enumerate(P): s[22 + i][15] = p
    print('C23..C30:', [s[22 + i][2] for i in range(8)])
json.dump(fx, open(out, 'w', encoding='utf-8'), ensure_ascii=False)
print('ok', len(fx), 'sheet')
