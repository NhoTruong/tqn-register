"""Kiểm thử Weather Engine POC. Chạy: python3 test_engine.py (cần đã chạy run_two_projects.py ra thư mục out/)."""
import copy, json, os, sys, tempfile, dataclasses
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import EngineError, normalize, require_complete, run, view_model
from run_two_projects import PROJECTS, HERE

ok = fail = 0
def check(name, cond, info=""):
    global ok, fail
    print(("PASS " if cond else "FAIL ") + name + (f" — {info}" if info else "")); ok += cond; fail += (not cond)

def expect(name, fn, state):
    try:
        fn(); check(name, False, "không bị chặn")
    except EngineError as e:
        check(name, e.state == state, f"{e.state}: {e}")

(pa, fa), (pb, fb) = PROJECTS
ra = json.load(open(os.path.join(HERE, fa), encoding="utf-8")); rb = json.load(open(os.path.join(HERE, fb), encoding="utf-8"))
ba = normalize(pa, ra, ra["_fetch"]["http_date"]); bb = normalize(pb, rb, rb["_fetch"]["http_date"])

# W1 đúng số liệu API: so view model với raw
vm = view_model(pa, ba, "2026-11-19", "06:00", "CHAY_TAY")
check("W1 nhiệt cao/thấp ngày = daily API", (vm["tmax"], vm["tmin"]) == (ra["body"]["daily"]["temperature_2m_max"][0], ra["body"]["daily"]["temperature_2m_min"][0]), f'{vm["tmax"]}/{vm["tmin"]}')
h12 = [x for x in ba["hours"] if x["time"].startswith("2026-11-19T12:00")][0]
check("W2 giờ 12:00 = raw (28.7°C, 100%, 0.5 mm)", (h12["temp_c"], h12["pop_pct"], h12["rain_mm"]) == (28.7, 100, 0.5))
check("W3 đúng múi giờ (+07:00)", h12["time"].endswith("+07:00"), h12["time"])
check("W4 tổng mưa ca sáng = tổng raw 06–11", vm["slots"][0]["rain_sum"] == round(sum(ra["body"]["hourly"]["precipitation"][6:12]), 1), str(vm["slots"][0]["rain_sum"]))
# W5 hai dự án cách ly, khác số liệu
vb = view_model(pb, bb, "2026-11-19", "06:00", "CHAY_TAY")
check("W5 hai dự án khác dữ liệu", (vm["tmax"], vm["grid"]) != (vb["tmax"], vb["grid"]), f'{vm["tmax"]} vs {vb["tmax"]}')
out = os.path.join(HERE, "out")
ma = json.load(open(os.path.join(out, "COASTAL_QNG/WEATHER/2026/11/20261119/manifest.json"), encoding="utf-8"))
mb = json.load(open(os.path.join(out, "DEMO_B/WEATHER/2026/11/20261119/manifest.json"), encoding="utf-8"))
check("W6 mỗi dự án đủ 12 file (3 khung x 2 theme mobile + share + desktop + PDF A4), có SHA256", len(ma["files"]) == 12 == len(mb["files"]) and all(ma["files"].values()))
check("W7 manifest ghi provider/kind/issued/giấy phép", ma["kind"] == "FORECAST" and ma["provider"] == "Open-Meteo" and ma["license_status"] == "CHUA_XAC_MINH")
from engine import line_path
lp = line_path([{"temp_c": 24}, {"temp_c": 29}], 300, 34, 25)
check("W15 đường nhiệt: nhiệt tăng thì điểm đi lên", lp["pts"][1][1] < lp["pts"][0][1], str(lp["pts"]))
check("W16 dải 8 ngày đúng daily API, đánh dấu hôm nay", [d["tmax"] for d in vm["days8"]] == ra["daily8"]["daily"]["temperature_2m_max"] and vm["days8"][0]["active"])
check("W17 biểu đồ 3 giờ: 00..21 + 00 ngày sau, đúng giá trị", [x["hour"] for x in vm["chart3h"]] == [0,3,6,9,12,15,18,21,0] and vm["chart3h"][4]["temp_c"] == ra["body"]["hourly"]["temperature_2m"][12])
# W8 chạy trùng cùng slot, cùng dữ liệu -> bỏ qua, không phát hành bản thứ hai
st, m = run(pa, os.path.join(HERE, fa), out, "2026-11-19", log=lambda *_: None)
check("W8 chạy trùng -> SKIPPED_DUPLICATE, vẫn version 1", st == "SKIPPED_DUPLICATE" and m["version"] == 1, st)
# W9 thiếu giờ -> PARTIAL, không render
bad = copy.deepcopy(ra);
for k in bad["body"]["hourly"]: bad["body"]["hourly"][k] = bad["body"]["hourly"][k][:15]
b2 = normalize(pa, bad, bad["_fetch"]["http_date"])
expect("W9 thiếu giờ 15–23 -> PARTIAL (không phát hành báo cáo thiếu)", lambda: require_complete(b2, "2026-11-19"), "PARTIAL")
# W10 sai đơn vị -> FAILED
bu = copy.deepcopy(ra); bu["body"]["hourly_units"]["temperature_2m"] = "°F"
expect("W10 đơn vị °F -> FAILED", lambda: normalize(pa, bu, bu["_fetch"]["http_date"]), "FAILED")
# W11 lệch múi giờ -> FAILED
bz = copy.deepcopy(ra); bz["body"]["utc_offset_seconds"] = 0
expect("W11 phản hồi UTC thay vì +7 -> FAILED", lambda: normalize(pa, bz, bz["_fetch"]["http_date"]), "FAILED")
# W12 GPS chưa xác nhận -> BLOCKED; W13 JSON hỏng -> FAILED
expect("W12 GPS chưa xác nhận -> BLOCKED", lambda: run(dataclasses.replace(pa, gps_confirmed=False), os.path.join(HERE, fa), out, "2026-11-19"), "BLOCKED")
with tempfile.TemporaryDirectory() as td:
    p = os.path.join(td, "x.json"); open(p, "w").write("{hỏng")
    expect("W13 JSON hỏng -> FAILED", lambda: run(pa, p, td, "2026-11-19"), "FAILED")
    check("W14 job lỗi không để lại manifest", not any(f == "manifest.json" for _, _, fs in os.walk(td) for f in fs))
print(f"TONG: PASS={ok} FAIL={fail}")
sys.exit(1 if fail else 0)
