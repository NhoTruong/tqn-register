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
        "weather_raw": json.load(open(os.path.join(ROOT, "weather-poc", "data", "COASTAL_20261119_1114.json"), encoding="utf-8"))["body"]}
json.dump(data, open(os.path.join(OUT, "data.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
code = "".join(open(os.path.join(ROOT, "apps-script", f + ".gs"), encoding="utf-8").read() + "\n" for f in ["Config", "Read", "Import", "QC", "Remind", "Weather", "Auto", "Menu"])
open(os.path.join(OUT, "tqn.js"), "w", encoding="utf-8").write(code)
open(os.path.join(OUT, "sidebar.html"), "w", encoding="utf-8").write(open(os.path.join(ROOT, "apps-script", "Sidebar.html"), encoding="utf-8").read())
print("tab:", len(sheets), "| ẩn vì thông tin cá nhân:", hidden_pii, "| data.json", os.path.getsize(os.path.join(OUT, "data.json")) // 1024, "KB")
