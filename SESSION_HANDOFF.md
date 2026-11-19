# Bàn giao phiên — 19/11/2026
- **Nhánh:** `v3/g0-g2`, đi ra từ `main` d8066d9.
- **Đã làm:** G0, G1, G2 theo chỉ thị V3, gồm `docs/*`, `db/schema_v0.sql` và `db/tests/run_db_tests.sh` (17/17 PASS trên PostgreSQL 16 trong máy làm việc).
- **Chạy lại test DB:** dựng PostgreSQL cục bộ, rồi chạy `PGHOST=... PGPORT=... PGUSER=... bash db/tests/run_db_tests.sh`.
- **GPS Coastal anh gửi lúc 11:12:** 15.099038, 108.894552. Coi là đã xác nhận, nguồn: tin nhắn của anh.
- **Đang chờ anh:**
  1. Chọn hosting (ADR-007).
  2. Cho phép NuGet/dotnet hoặc build trên máy anh.
  3. Xác nhận giấy phép thời tiết: Open-Meteo thương mại hoặc phi thương mại, hay dùng Google kèm API key.
  4. Cài S1 lên bản sao Register.
- **Việc tiếp theo:** POC báo cáo thời tiết PNG/PDF cho 2 dự án (Coastal + DEMO_B), dùng dữ liệu Open-Meteo thật; sau đó Sprint 1 C#.
