# G0 — Môi trường (V3)
Ngày kiểm: 19/11/2026, 11:09 (giờ VN). Lệnh và kết quả gốc: `evidence/v3/G0_lenh_da_chay.txt`.

## Ma trận

| Hạng mục | Máy làm việc của Claude (Linux, cloud) | Máy anh Nho (DESKTOP-TILDKVA) | Trạng thái |
|---|---|---|---|
| Repo | `tqn-register` (git cục bộ). Nhánh `main` có 3 commit: 50360dd, a697db9, d8066d9. Nhánh mới là `v3/g0-g2` | Chưa có repo GitHub cho Register | HAVE (cục bộ). GitHub: NOT_HAVE |
| .NET 10 SDK | `dotnet: command not found`. Tải dotnet và NuGet bị proxy chặn (403) | Có .NET Framework 4.8.1 (đọc qua TQN_MCP_Hub). Chưa đọc được `dotnet --list-sdks` vì công cụ trên máy anh chỉ cho build theo danh sách trắng | **BLOCKED** |
| PostgreSQL | Có PostgreSQL 16.15 (cả server và client). Đã dựng cụm thử ở cổng 55432 và chạy test thật | Chưa kiểm | HAVE (để thử). Staging/prod: NOT_HAVE |
| Docker / Testcontainers | Có client, không có daemon | Chưa kiểm | NOT_HAVE |
| Node / Python | Node 22.23.5, Python 3.13.18 | — | HAVE |
| Excel Windows | — | Excel đang chạy, file `QL_TRINH_DUYET_COA.xlsx` đang mở | HAVE. Test VBA: NOT_RUN |
| Google Sheets (Register) | Có link gviz chỉ đọc. Chưa có quyền ghi | Anh là người quản lý | Đọc: HAVE. Ghi/Apps Script: chưa cài |
| Google Drive / Sheets API (OAuth, service account) | Chưa có credential. Theo chỉ thị, không được nhận credential qua chat | — | **BLOCKED** |
| Open-Meteo | Proxy chặn từ máy làm việc. Đã gọi thật được qua trình duyệt tích hợp trên máy anh (ngày 19/11) | — | Kỹ thuật: HAVE. Giấy phép thương mại: **BLOCKED** |
| Google Weather API | Cần API key và tài khoản thanh toán Google Cloud | — | **BLOCKED** |
| Hosting server C# (06:00 phải chạy cả khi máy tắt) | — | Không dùng máy anh làm server được | NOT_HAVE (cần chọn, xem ADR-007) |
| GPS công trường đã xác nhận | — | — | **BLOCKED**. Chờ anh gửi tọa độ |

## Hệ quả
- Phần C# thuộc G3: **viết được code nhưng chưa biên dịch được ở đây**. Muốn ra bản biên dịch cần một trong hai cách:
  - anh cho phép domain `api.nuget.org` và `builds.dotnet.microsoft.com` trong Cài đặt → Capabilities;
  - hoặc build trên máy anh, nếu máy có .NET 10 SDK.
- Mọi khẳng định về schema/DB ở G2 đều đã **chạy thật** trên PostgreSQL 16 (`evidence/v3/G2_db_test_log.txt`). Các thứ chỉ chạy ở máy làm việc **không** được coi là đã kiểm trên môi trường của anh.
