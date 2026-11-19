# TQN Weather Engine — POC

Bản chạy thật bằng Python, dùng làm mẫu chuẩn để viết bản C# theo ADR-009.

## Chạy
1. `python3 run_two_projects.py` — tạo báo cáo cho 2 dự án vào `out/{DỰ_ÁN}/WEATHER/YYYY/MM/YYYYMMDD/`.
2. `python3 test_engine.py` — chạy 14 bài test.

## Dữ liệu
- File: `data/*.json`, dữ liệu Open-Meteo thật, lấy lúc 11:14 ngày 19/11/2026 qua trình duyệt trên máy anh.
- Đã đối chiếu SHA256 của chuỗi giờ với trình duyệt.
- Loại dữ liệu là FORECAST.

## Đầu ra mỗi dự án
- 6 file PNG:
  - mobile 1080×1920;
  - share 1080×1350;
  - desktop 1920×1080;
  - mỗi kích thước có 2 theme sáng/tối.
- 2 file PDF A4 (sáng/tối).
- `raw.json`, `normalized.json`, `manifest.json` (có SHA256).

## Chặn phát hành sai
- Thiếu giờ 06–23 → PARTIAL.
- Sai đơn vị hoặc sai múi giờ → FAILED.
- GPS chưa xác nhận → BLOCKED.
- Bố cục bị tràn chữ → FAILED.
- Chạy trùng slot → bỏ qua.

## Chưa làm
- Tải lên Drive: BLOCKED, chưa có service account.
- Lịch 06:00 trên cloud: chờ chọn hosting (ADR-007).
- Bản C#: BLOCKED, chưa có .NET SDK.
