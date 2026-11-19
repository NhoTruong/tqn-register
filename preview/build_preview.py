"""Dựng bản xem thử Register chạy đúng code Apps Script trong trình duyệt.
Đọc test/fixture_live.json (chụp Register), bỏ tab có thông tin cá nhân, ghi preview/site/{data.json,tqn.js,sidebar.html,index.html}."""
import json, os, re, sys, hashlib
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "preview", "site"); os.makedirs(OUT, exist_ok=True)
fx = json.load(open(os.path.join(ROOT, "test", "fixture_live.json"), encoding="utf-8"))
PII = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+|(?<!\d)0\d{9}(?!\d)|(?<!\d)0\d{2}[ .]\d{3}[ .]\d{4}(?!\d)")
sheets, hidden_pii = [], []
for name, sh in fx.items():
    vals = sh["values"]
    last = 0
    for i, r in enumerate(vals):
        if any(x not in ("", None) for x in r): last = i + 1
    vals = vals[:last]
    lastc = 0
    for r in vals:
        for j in range(len(r) - 1, -1, -1):
            if r[j] not in ("", None): lastc = max(lastc, j + 1); break
    vals = [r[:lastc] for r in vals]
    flat = " ".join(str(x) for r in vals for x in r if isinstance(x, (str, int, float)))
    if PII.search(flat):
        hidden_pii.append(name)
        vals = [["(Tab này có số điện thoại/email nhân sự nên không đưa vào bản xem thử. Bản Google Sheets thật vẫn giữ nguyên.)"]]
        forms = {}
    else:
        forms = {}
        for i, r in enumerate(sh.get("formulas", [])[:last]):
            for j, f in enumerate(r[:lastc]):
                if f: forms[f"{i+1},{j+1}"] = f
    sheets.append({"name": name, "values": vals, "formulas": forms})
data = {"source": "Chụp Register (fixture_live) — chỉ đọc", "sha256_fixture": hashlib.sha256(open(os.path.join(ROOT, "test", "fixture_live.json"), "rb").read()).hexdigest(),
        "sheets": sheets, "pii_hidden": hidden_pii,
        "html": {n: open(os.path.join(ROOT, "apps-script", n + ".html"), encoding="utf-8").read() for n in ["Sidebar", "Workflow", "NewDoc", "Manager"]},
        "weather_raw": json.load(open(os.path.join(ROOT, "weather-poc", "data", "COASTAL_20261119_1114.json"), encoding="utf-8"))["body"]}
json.dump(data, open(os.path.join(OUT, "data.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
code = "".join(open(os.path.join(ROOT, "apps-script", f + ".gs"), encoding="utf-8").read() + "\n" for f in ["Config", "Read", "Import", "QC", "Remind", "Weather", "Workflow", "Ai", "Theme", "Auto", "Menu", "Docs", "Manager"])
open(os.path.join(OUT, "tqn.js"), "w", encoding="utf-8").write(code)
open(os.path.join(OUT, "sidebar.html"), "w", encoding="utf-8").write(open(os.path.join(ROOT, "apps-script", "Sidebar.html"), encoding="utf-8").read())
print("tab:", len(sheets), "| ẩn vì thông tin cá nhân:", hidden_pii, "| data.json", os.path.getsize(os.path.join(OUT, "data.json")) // 1024, "KB")

# ---- Định dạng gốc của Register (màu nền, chữ đậm, gộp ô, độ rộng cột, tô màu có điều kiện) ----
import openpyxl
THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']
def _col(c):
    try:
        if c is None: return None
        if c.type == 'rgb' and isinstance(c.rgb, str) and len(c.rgb) == 8 and c.rgb[:2] != '00': base = c.rgb[2:]
        elif c.type == 'theme' and c.theme is not None and c.theme < len(THEME): base = THEME[c.theme]
        else: return None
        t = c.tint or 0
        if t:
            rgb = [int(base[i:i + 2], 16) for i in (0, 2, 4)]
            rgb = [round(v + (255 - v) * t) if t > 0 else round(v * (1 + t)) for v in rgb]
            base = ''.join('%02X' % max(0, min(255, v)) for v in rgb)
        return '#' + base
    except Exception:
        return None
wb = openpyxl.load_workbook(os.path.join(ROOT, '..', 'zipv2', 'REG.xlsx'))
styles = {}
for s in sheets:
    if s['name'] in hidden_pii or s['name'] not in wb.sheetnames: continue
    ws = wb[s['name']]
    R = len(s['values']); C = max([len(r) for r in s['values']] + [1])
    st = {}
    for row in ws.iter_rows(min_row=1, max_row=R, max_col=C):
        for c in row:
            bg = _col(c.fill.fgColor) if c.fill and c.fill.fill_type == 'solid' else None
            fc = _col(c.font.color) if c.font and c.font.color else None
            b = 1 if c.font and c.font.b else 0
            al = c.alignment.horizontal if c.alignment and c.alignment.horizontal in ('center', 'right') else ''
            if bg or (fc and fc not in ('#000000',)) or b or al:
                st[f"{c.row},{c.column}"] = [bg or '', fc if fc != '#000000' else '', b, al]
    cw = {}
    for k, d in ws.column_dimensions.items():
        if d.width: cw[k] = round(d.width * 7 + 5)
    cf = []
    for rng in ws.conditional_formatting:
        for r in rng.rules:
            if r.type == 'cellIs' and r.operator == 'equal' and r.formula and r.dxf and r.dxf.fill:
                cf.append([str(rng.sqref), r.formula[0].strip('"'), _col(r.dxf.fill.fgColor) or _col(r.dxf.fill.bgColor) or '', _col(r.dxf.font.color) if r.dxf.font and r.dxf.font.color else ''])
    dvs = []
    for dv in ws.data_validations.dataValidation:
        if dv.type != 'list' or not dv.formula1 or '#REF' in dv.formula1: continue
        f1 = dv.formula1.strip()
        if f1.startswith('"'): opts = [x.strip() for x in f1.strip('"').split(',') if x.strip()]
        else:
            ref = f1.replace('$', ''); sh_name, rg = ref.split('!') if '!' in ref else (s['name'], ref)
            sh_name = sh_name.strip("'")
            opts = [str(c.value) for row in wb[sh_name][rg] for c in row if c.value not in (None, '')] if sh_name in wb.sheetnames else []
        if opts: dvs.append([str(dv.sqref), opts])
    styles[s['name']] = {'s': st, 'cw': cw, 'm': [str(m) for m in ws.merged_cells.ranges], 'cf': cf, 'dv': dvs}
data['styles'] = styles
json.dump(data, open(os.path.join(OUT, "data.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print('định dạng gốc:', len(styles), 'tab, data.json', os.path.getsize(os.path.join(OUT, "data.json")) // 1024, 'KB')
