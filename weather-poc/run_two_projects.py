"""Chạy job dự báo cho 2 dự án (Coastal + DEMO_B) từ dữ liệu Open-Meteo thật đã lấy lúc 11:14 19/11/2026."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import Project, run

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECTS = [
    (Project(code="COASTAL_QNG", name="The Rehaus 4 & 5 — Coastal Quảng Ngãi",
             address="KĐT sinh thái Coastal, Quảng Ngãi", lat=15.09903804958772, lon=108.89455225180114,
             timezone="Asia/Ho_Chi_Minh", gps_confirmed=True,
             gps_note="anh Nho gửi 11:12 19/11/2026"), "data/COASTAL_20261119_1114.json"),
    (Project(code="DEMO_B", name="Dự án DEMO B (thử cách ly)", address="Tọa độ thử tại TP.HCM — không phải công trình thật",
             lat=10.7769, lon=106.7009, timezone="Asia/Ho_Chi_Minh", gps_confirmed=True, demo=True,
             gps_note="tọa độ DEMO"), "data/DEMO_B_20261119_1114.json"),
]

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out")
    for p, f in PROJECTS:
        run(p, os.path.join(HERE, f), out, "2026-11-19", slot="06:00", run_kind="CHAY_TAY")
