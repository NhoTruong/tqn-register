"""Mo phong dung logic modTQN_Import/modTQN_Rules (MAX1) tren ban sao Register, CHI DOC.
Muc dich: lay bang chung loi truoc khi quyet dinh KEEP/REFACTOR. Khong phai VBA chay that."""
import sys, re, csv, hashlib, unicodedata, openpyxl, datetime
src = sys.argv[1]; out = sys.argv[2]
h0 = hashlib.sha256(open(src,'rb').read()).hexdigest()
wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
def clean(v):
    if v is None: return ''
    s = str(v).replace('\t',' ').replace('\r',' ').replace('\n',' ').strip()
    while '  ' in s: s = s.replace('  ',' ')
    return s
def dockey(no):  # modTQN_Rules.bas TQN_DocKey
    s = clean(no)
    if len(s) > 7 and s[6] == '-' and s[:6].isdigit(): s = s[7:].strip()
    p = s.upper().rfind('REV')
    if p >= 0:
        tail = s[p+3:].strip()
        if 0 < len(tail) <= 3 and tail.isdigit():
            s = s[:p].strip()
            while s and s[-1] in '-_': s = s[:-1].strip()
    return s.upper()
def rev_from_no(no):
    s = clean(no).upper(); p = s.rfind('REV')
    if p < 0: return ''
    t = s[p+3:].strip(); return '%02d' % int(t) if 0 < len(t) <= 3 and t.isdigit() else ''
def revision(v, no):
    s = clean(v).upper()
    if s.startswith('REV'): s = s[3:].strip()
    if s == '': return rev_from_no(no)
    if len(s) < 4 and s.replace('.','',1).isdigit(): return '%02d' % int(float(s))
    return s
def noacc(s): return ''.join(c for c in unicodedata.normalize('NFD', s.lower()) if unicodedata.category(c) != 'Mn').replace('đ','d')
def status(v):  # TQN_NormalizeStatus
    s = noacc(clean(v))
    if s == '': return 'DRAFT'
    if 'khong duyet' in s or 'tu choi' in s or 'tra ve' in s: return 'REJECTED_RESUBMIT'
    if 'duyet co' in s or 'chap thuan co' in s: return 'APPROVED_WITH_COMMENTS'
    if 'da duyet' in s or 'chap thuan' in s: return 'APPROVED'
    if 'da phan hoi' in s: return 'RESPONDED'
    if 'da hoan thanh' in s: return 'CLOSED'
    if 'da trinh' in s or 'dang cho' in s or 'chua phan hoi' in s: return 'SUBMITTED'
    if 'dang tham dinh' in s or 'dang xu ly' in s: return 'IN_REVIEW'
    return 'DRAFT'
TABS = [('SD','SDS',13,18),('M0S','MSS',13,18),('OTS','OTS',13,18),('MAT','MAS',13,17),
        ('SCH','SCH',12,17),('RFI','RFI',12,24),('RFC','RFC',12,16),('RFA','RFA',12,16)]
seen = {}; rows_out = []; ev = []
for tab, typ, ctv, ccd in TABS:
    ws = wb[tab]; data = list(ws.iter_rows(values_only=True))
    for r in range(6, len(data)+1):
        row = data[r-1] + (None,)*30
        no = clean(row[1])
        if not no or '-' not in no or len(no) < 12 or 'COA' not in no.upper(): continue
        subj = clean(row[2])
        if not subj: ev.append((tab, r, no, 'BO QUA: thieu noi dung')); continue
        rv = revision(row[5], no); key = dockey(no); k = (typ, key, rv)
        tv, cd = row[ctv-1], row[ccd-1]
        tvd = tv if isinstance(tv, datetime.datetime) else None
        if tv not in (None, '') and tvd is None:
            ev.append((tab, r, no, f'COT TVGS {openpyxl.utils.get_column_letter(ctv)} la chu, khong phai ngay -> bi bo trong: "{clean(tv)[:60]}"'))
        if k in seen:
            ev.append((tab, r, no, f'MAT DONG: trung Type+Key+Rev voi {seen[k]} -> TQN_AppendRecord bo qua, khong bao loi rieng | "{subj[:50]}"'))
            continue
        seen[k] = f'{tab}!{r}'
        st = status(row[4])
        if clean(row[4]) and noacc(clean(row[4])) == 'da hoan thanh':
            ev.append((tab, r, no, 'Trang thai "Da hoan thanh" -> CLOSED (trong SD nghia la ve xong noi bo, chua trinh)'))
        rows_out.append((typ, no, key, rv, st, clean(row[4]), tab, r))
with open(out, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f); w.writerow(['Sheet','Dong','Ma','Phat hien']); w.writerows(ev)
h1 = hashlib.sha256(open(src,'rb').read()).hexdigest()
print('SHA256 truoc:', h0); print('SHA256 sau  :', h1, '| khong doi' if h0 == h1 else '| THAY DOI!')
print('Ban ghi nhap duoc:', len(rows_out))
from collections import Counter
print('Theo loai:', dict(Counter(x[0] for x in rows_out)))
print('Phat hien:', len(ev)); [print(' ', *e) for e in ev]
