"""TQN Weather Engine — POC (Python), dùng làm golden reference cho bản C# (ADR-009).

Luồng: raw JSON (provider) -> chuẩn hóa (WeatherBundle) -> kiểm đủ dữ liệu -> khóa chống trùng
-> view model -> HTML/CSS -> PNG (mobile 1080x1920, share 1080x1350, desktop 1920x1080) + PDF A4
-> manifest.json (SHA256). Không chụp màn hình/scrape ứng dụng thời tiết nào.
"""
import hashlib, json, math, os, sys, datetime as dt
from dataclasses import dataclass, field, asdict
from zoneinfo import ZoneInfo
from jinja2 import Environment, FileSystemLoader, select_autoescape

HERE = os.path.dirname(os.path.abspath(__file__))
REQUIRED_HOURLY = ["temperature_2m", "apparent_temperature", "relative_humidity_2m",
                   "precipitation_probability", "precipitation", "weather_code",
                   "wind_speed_10m", "wind_direction_10m", "wind_gusts_10m"]
EXPECTED_UNITS = {"temperature_2m": "°C", "apparent_temperature": "°C", "relative_humidity_2m": "%",
                  "precipitation_probability": "%", "precipitation": "mm",
                  "wind_speed_10m": "km/h", "wind_gusts_10m": "km/h"}

# Mã hiện tượng nội bộ TQN (chuẩn hóa từ WMO, độc lập nhà cung cấp)
WMO = {0: ("NANG", "Trời quang"), 1: ("IT_MAY", "Ít mây"), 2: ("MAY_RAI_RAC", "Mây rải rác"),
       3: ("NHIEU_MAY", "Nhiều mây"), 45: ("SUONG_MU", "Sương mù"), 48: ("SUONG_MU", "Sương mù đóng băng"),
       51: ("MUA_PHUN", "Mưa phùn nhẹ"), 53: ("MUA_PHUN", "Mưa phùn"), 55: ("MUA_PHUN", "Mưa phùn dày"),
       61: ("MUA", "Mưa nhẹ"), 63: ("MUA", "Mưa vừa"), 65: ("MUA_TO", "Mưa to"),
       80: ("MUA_RAO", "Mưa rào nhẹ"), 81: ("MUA_RAO", "Mưa rào"), 82: ("MUA_TO", "Mưa rào rất to"),
       95: ("DONG", "Dông"), 96: ("DONG", "Dông kèm mưa đá"), 99: ("DONG", "Dông kèm mưa đá lớn")}
HUONG = ["B", "BĐB", "ĐB", "ĐĐB", "Đ", "ĐĐN", "ĐN", "NĐN", "N", "NTN", "TN", "TTN", "T", "TTB", "TB", "BTB"]
THU = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ nhật"]
SLOTS = [("sang", "Sáng", 6, 11), ("chieu", "Chiều", 12, 17), ("toi", "Tối", 18, 23)]


class EngineError(Exception):
    def __init__(self, state, msg):
        super().__init__(msg); self.state = state


def sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def wmo(code):
    return WMO.get(int(code), ("KHAC", f"Mã {code}"))


def huong(deg):
    return HUONG[int((deg % 360) / 22.5 + 0.5) % 16]


@dataclass
class Project:
    code: str
    name: str
    address: str
    lat: float
    lon: float
    timezone: str
    gps_confirmed: bool
    provider: str = "OPEN_METEO"
    demo: bool = False
    gps_note: str = ""
    license_status: str = "CHUA_XAC_MINH"  # CHUA_XAC_MINH | DA_XAC_MINH


def normalize(project: Project, raw: dict, http_date: str) -> dict:
    """OpenMeteoProvider -> WeatherBundle (dict). Ném EngineError nếu thiếu/sai."""
    b = raw.get("body") if "body" in raw else raw
    if not isinstance(b, dict) or "hourly" not in b:
        raise EngineError("FAILED", "Phản hồi không đúng cấu trúc (thiếu hourly)")
    if b.get("utc_offset_seconds") != int(dt.datetime.now(ZoneInfo(project.timezone)).utcoffset().total_seconds()):
        raise EngineError("FAILED", f"Múi giờ phản hồi lệch với {project.timezone}")
    for k, u in EXPECTED_UNITS.items():
        if b.get("hourly_units", {}).get(k) != u:
            raise EngineError("FAILED", f"Đơn vị {k} không phải {u}")
    h = b["hourly"]
    n = len(h.get("time", []))
    for k in REQUIRED_HOURLY:
        if k not in h or len(h[k]) != n:
            raise EngineError("PARTIAL", f"Thiếu chuỗi giờ {k}")
    hours = []
    for i in range(n):
        vals = {k: h[k][i] for k in REQUIRED_HOURLY}
        if any(v is None for v in vals.values()):
            continue
        t = dt.datetime.fromisoformat(h["time"][i]).replace(tzinfo=ZoneInfo(project.timezone))
        key, text = wmo(vals["weather_code"])
        hours.append({"time": t.isoformat(), "date": t.date().isoformat(), "hour": t.hour,
                      "temp_c": vals["temperature_2m"], "feels_c": vals["apparent_temperature"],
                      "rh_pct": vals["relative_humidity_2m"], "pop_pct": vals["precipitation_probability"],
                      "rain_mm": vals["precipitation"], "wind_kmh": vals["wind_speed_10m"],
                      "gust_kmh": vals["wind_gusts_10m"], "wind_dir": huong(vals["wind_direction_10m"]),
                      "cond": key, "cond_text": text, "wmo": vals["weather_code"]})
    issued = dt.datetime.strptime(http_date, "%a, %d %b %Y %H:%M:%S GMT").replace(tzinfo=dt.timezone.utc)
    return {"kind": "FORECAST", "provider": "Open-Meteo", "model": "best_match",
            "issued_at_utc": issued.isoformat(), "grid_lat": b["latitude"], "grid_lon": b["longitude"],
            "elevation_m": b.get("elevation"), "current": b.get("current"), "daily": b.get("daily"),
            "hours": hours}


def require_complete(bundle: dict, local_date: str):
    """Không phát hành báo cáo trống/thiếu: phải đủ 06..23 của ngày báo cáo."""
    have = {x["hour"] for x in bundle["hours"] if x["date"] == local_date}
    missing = [hh for hh in range(6, 24) if hh not in have]
    if missing:
        raise EngineError("PARTIAL", f"Thiếu giờ {missing} của ngày {local_date} — không render")


def view_model(project: Project, bundle: dict, local_date: str, slot_label: str, run_kind: str):
    tz = ZoneInfo(project.timezone)
    day = [x for x in bundle["hours"] if x["date"] == local_date]
    d = dt.date.fromisoformat(local_date)
    nxt = (d + dt.timedelta(days=1)).isoformat()
    tomorrow = [x for x in bundle["hours"] if x["date"] == nxt]
    daily = bundle["daily"]; di = daily["time"].index(local_date)
    slots = []
    for key, name, a, z in SLOTS:
        hs = [x for x in day if a <= x["hour"] <= z]
        mx = max(hs, key=lambda x: x["pop_pct"])
        slots.append({"key": key, "name": name, "range": f"{a:02d}:00–{z:02d}:00", "hours": hs,
                      "tmin": min(x["temp_c"] for x in hs), "tmax": max(x["temp_c"] for x in hs),
                      "rain_sum": round(sum(x["rain_mm"] for x in hs), 1), "pop_max": mx["pop_pct"],
                      "pop_max_at": f'{mx["hour"]:02d}:00',
                      "wind_max": max(x["wind_kmh"] for x in hs), "gust_max": max(x["gust_kmh"] for x in hs),
                      "rh_avg": round(sum(x["rh_pct"] for x in hs) / len(hs)),
                      "main": max(set(x["cond_text"] for x in hs), key=[x["cond_text"] for x in hs].count)})
    work = [x for x in day if 6 <= x["hour"] <= 17]
    peak = max(work, key=lambda x: x["pop_pct"])
    summary = (f'{wmo(daily["weather_code"][di])[1]}. Cao {daily["temperature_2m_max"][di]:g}°, '
               f'thấp {daily["temperature_2m_min"][di]:g}°. Khả năng mưa giờ làm việc cao nhất '
               f'{peak["pop_pct"]}% lúc {peak["hour"]:02d}:00; tổng mưa cả ngày {daily["precipitation_sum"][di]:g} mm.')
    issued_local = dt.datetime.fromisoformat(bundle["issued_at_utc"]).astimezone(tz)
    days = []
    for j, ds in enumerate(daily["time"]):
        dd = dt.date.fromisoformat(ds)
        days.append({"label": "Hôm nay" if ds == local_date else THU[dd.weekday()], "date": dd.strftime("%d/%m"),
                     "tmax": daily["temperature_2m_max"][j], "tmin": daily["temperature_2m_min"][j],
                     "cond": wmo(daily["weather_code"][j])[0], "cond_text": wmo(daily["weather_code"][j])[1],
                     "rain": daily["precipitation_sum"][j], "pop_pct": daily["precipitation_probability_max"][j],
                     "active": ds == local_date})
    dist_km = haversine(project.lat, project.lon, bundle["grid_lat"], bundle["grid_lon"])
    return {"p": asdict(project), "date_text": f'{THU[d.weekday()]}, {d.strftime("%d/%m/%Y")}',
            "local_date": local_date, "slot": slot_label, "run_kind": run_kind,
            "issued_local": issued_local.strftime("%H:%M %d/%m/%Y"),
            "kind": bundle["kind"], "provider": bundle["provider"], "model": bundle["model"],
            "grid": f'{bundle["grid_lat"]:.4f}, {bundle["grid_lon"]:.4f}', "grid_km": round(dist_km, 1),
            "tmax": daily["temperature_2m_max"][di], "tmin": daily["temperature_2m_min"][di],
            "rain_day": daily["precipitation_sum"][di], "pop_day": daily["precipitation_probability_max"][di],
            "wind_day": daily["wind_speed_10m_max"][di], "gust_day": daily["wind_gusts_10m_max"][di],
            "rh_day": round(sum(x["rh_pct"] for x in day) / len(day)),
            "cond_day": wmo(daily["weather_code"][di])[0], "cond_day_text": wmo(daily["weather_code"][di])[1],
            "summary": summary, "slots": slots, "days": days, "tomorrow_hours": tomorrow,
            "peak": peak, "hours_all": [x for x in day if x["hour"] >= 6]}


def haversine(a1, o1, a2, o2):
    r = 6371.0; p1, p2 = math.radians(a1), math.radians(a2)
    dp, dl = p2 - p1, math.radians(o2 - o1)
    return 2 * r * math.asin(math.sqrt(math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2))


def render_html(vm: dict, layout: str, theme: str) -> str:
    env = Environment(loader=FileSystemLoader(os.path.join(HERE, "templates")), autoescape=select_autoescape())
    env.globals.update(line_path=line_path)
    return env.get_template("report.html").render(vm=vm, layout=layout, theme=theme)


def line_path(hours, w, h, pad=10, lo=None, hi=None):
    ts = [x["temp_c"] for x in hours]
    lo = min(ts) - 1 if lo is None else lo; hi = max(ts) + 1 if hi is None else hi
    n = len(ts); step = (w - 2 * pad) / max(n - 1, 1)
    pts = [(pad + i * step, h - pad - (t - lo) / (hi - lo) * (h - 2 * pad)) for i, t in enumerate(ts)]
    return {"d": "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts), "pts": pts}


OVERFLOW_JS = """[...document.querySelectorAll('.wrap *')].some(e=>{const r=e.getBoundingClientRect();return r.height>0&&(r.bottom>innerHeight+0.5||r.right>innerWidth+0.5)})"""
LAYOUTS = {"mobile": (540, 960, 2), "share": (540, 675, 2), "desktop": (960, 540, 2)}


def run(project: Project, raw_path: str, out_root: str, local_date: str, slot="06:00",
        run_kind="CHAY_TAY", themes=("sang", "toi"), log=print):
    """Một job: trả về (state, manifest|None). Idempotent theo (project, date, FORECAST, slot)."""
    started = dt.datetime.now(dt.timezone.utc)
    if not project.gps_confirmed:
        raise EngineError("BLOCKED", "GPS chưa xác nhận — không chạy")
    raw_bytes = open(raw_path, "rb").read()
    raw_sha = sha256(raw_bytes)
    try:
        raw = json.loads(raw_bytes)
    except json.JSONDecodeError as e:
        raise EngineError("FAILED", f"JSON hỏng: {e}")
    out_dir = os.path.join(out_root, project.code, "WEATHER", local_date[:4], local_date[5:7], local_date.replace("-", ""))
    man_path = os.path.join(out_dir, "manifest.json")
    version = 1
    if os.path.exists(man_path):
        old = json.load(open(man_path, encoding="utf-8"))
        if old["raw_sha256"] == raw_sha and old["state"] == "COMPLETE":
            log(f"[{project.code}] BỎ QUA: đã có báo cáo {local_date} {slot} cùng dữ liệu (chống trùng)")
            return "SKIPPED_DUPLICATE", old
        version = old["version"] + 1
    bundle = normalize(project, raw, raw["_fetch"]["http_date"])
    require_complete(bundle, local_date)
    os.makedirs(out_dir, exist_ok=True)
    open(os.path.join(out_dir, "raw.json"), "wb").write(raw_bytes)
    json.dump(bundle, open(os.path.join(out_dir, "normalized.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    vm = view_model(project, bundle, local_date, slot, run_kind)
    files = {}
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        br = pw.chromium.launch(executable_path="/usr/bin/google-chrome", args=["--no-sandbox"])
        for theme in themes:
            for lay, (w, h, s) in LAYOUTS.items():
                html = render_html(vm, lay, theme)
                pg = br.new_page(viewport={"width": w, "height": h}, device_scale_factor=s)
                pg.route("**/*", lambda r: r.abort() if not r.request.url.startswith(("data:", "about:")) else r.continue_())
                pg.set_content(html, wait_until="load")
                overflow = pg.evaluate(OVERFLOW_JS)
                if overflow:
                    raise EngineError("FAILED", f"Bố cục {lay}/{theme} bị tràn — không phát hành ảnh cắt chữ")
                name = f"weather_{slot.replace(':', '')}_forecast_{lay}_{theme}.png"
                pg.screenshot(path=os.path.join(out_dir, name), full_page=False)
                files[name] = None
                pg.close()
            pg = br.new_page(viewport={"width": 794, "height": 1123})
            pg.route("**/*", lambda r: r.abort() if not r.request.url.startswith(("data:", "about:")) else r.continue_())
            pg.set_content(render_html(vm, "a4", theme), wait_until="load")
            if pg.evaluate(OVERFLOW_JS):
                raise EngineError("FAILED", f"Bố cục A4/{theme} bị tràn — không phát hành PDF thiếu dòng")
            name = f"weather_{slot.replace(':', '')}_forecast_A4_{theme}.pdf"
            pg.pdf(path=os.path.join(out_dir, name), format="A4", print_background=True,
                   margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
            files[name] = None
            pg.close()
        br.close()
    for k in files:
        files[k] = sha256(open(os.path.join(out_dir, k), "rb").read())
    man = {"project": project.code, "local_date": local_date, "job_type": "FORECAST", "slot": slot,
           "run_kind": run_kind, "version": version, "state": "COMPLETE", "kind": "FORECAST",
           "provider": bundle["provider"], "model": bundle["model"], "issued_at_utc": bundle["issued_at_utc"],
           "license_status": project.license_status, "raw_sha256": raw_sha,
           "started_at_utc": started.isoformat(),
           "finished_at_utc": dt.datetime.now(dt.timezone.utc).isoformat(), "files": files,
           "drive_file_ids": "CHUA_TAI_LEN (BLOCKED: chua co service account)"}
    json.dump(man, open(man_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    log(f"[{project.code}] COMPLETE v{version}: {len(files)} file -> {out_dir}")
    return "COMPLETE", man
