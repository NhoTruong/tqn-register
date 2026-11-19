# G2 — Quyết định kiến trúc (ADR)

## Ba phương án

| | P1. Excel .xlsm offline | P2. Google Sheets + Apps Script (dữ liệu trong Sheet, có khóa) | P3. Sheets/Excel làm client + backend DB (Supabase/PostgreSQL) + Cloud Run/Scheduler |
|---|---|---|---|
| Hợp với hiện trạng | Thấp: Register đang ở Google Sheets, cả nhóm dùng chung | **Cao**: cài thẳng vào bản Register live, không đổi cách nhập | Trung bình: phải dựng thêm hạ tầng |
| Đa người | Không an toàn (U25 BLOCKER) | Được với nhóm nhỏ nhờ `LockService`; có giới hạn quota | Tốt nhất (transaction) |
| Thời tiết 06:00 không mở máy | Không | Có, sai số trigger khoảng ±15 phút | Có, đúng phút (Cloud Scheduler) |
| In thẳng máy in, Word giữ đúng 100% | **Có** | Không in thẳng; Docs → PDF có thể lệch nhẹ | Tùy client |
| Chi phí, độ khó | Thấp | Thấp, không cần server | Cao: server, tài khoản, bảo trì |
| Rủi ro chính | lệch dữ liệu giữa các bản sao | quota 6 phút/lần chạy, giới hạn Sheet | vận hành, bảo mật, chi phí |

## ADR-001 — Nền tảng: **chọn P2 làm lõi, thiết kế sẵn lớp dữ liệu để lên P3**

- Lõi chạy trong Google Sheets + Apps Script, gắn vào Register live. Phần thêm vào là tab `TQN_*` và menu; không sửa 32 tab gốc.
- Excel Desktop chỉ là client phụ cho việc **in hàng loạt bằng mẫu Word gốc** (module VBA nhỏ, đọc file xuất từ Sheets). Không dùng .xlsm làm nơi lưu dữ liệu.
- Mọi đọc/ghi đi qua một lớp `Repository` (JS). Khi số dự án/người dùng vượt ngưỡng ở mục kiểm thử G8, đổi phần cài đặt của lớp này sang DB mà không sửa nghiệp vụ.
- Lý do: khớp đúng chỗ anh đang làm việc; chạy được ngay không cần server; đáp ứng "không mở Excel".

## ADR-002 — Nguồn sự thật

- Giai đoạn 1: 32 tab gốc vẫn là nơi người dùng nhập liệu. `TQN_DATA` là bảng chuẩn hóa sinh ra từ import idempotent, khóa `(project_id, doc_type, doc_key, revision, source_row_hash)`.
- Sổ trình duyệt là nguồn của Kết quả, Hạn, Tình trạng. Sổ chi tiết là nguồn của mã TBC và khối TVGS/CĐT. Ghép hai bên qua tên file hoặc mã; không ghép được thì đưa vào `TQN_QC`, không đoán.

## ADR-003 — Định danh hồ sơ và Revision

- `document_id` (ULID) cho hồ sơ logic; `revision_id` cho từng Rev; `submission_id` cho từng lượt trình. Số hiệu phát hành lưu nguyên văn, `doc_key` chỉ để so khớp.
- Trùng số hiệu khác nội dung (D3) → `CONFLICT`, giữ cả hai, chờ người quyết. Không xóa revision cũ.

## ADR-004 — Workflow

- Bảng trạng thái và chuyển trạng thái theo từng loại hồ sơ nằm trong cấu hình dự án (JSON có version). Mặc định Coastal dùng 7 trạng thái của Sổ trình duyệt; mỗi bên (TVGS/TVTK/CĐT) có `ReviewDecision` riêng A/B/C/D.

## ADR-005 — Thời tiết 06:00

- Apps Script trigger theo `timezone` dự án, ghi rõ `scheduled_for`/`fetched_at`/độ trễ; khóa theo `(project_id, ngày, 06:00, FORECAST)` chống chạy trùng. Dữ liệu từ nhà cung cấp có giấy phép phù hợp (Open-Meteo chỉ dùng thử nếu điều khoản cho phép). Render HTML → PDF/PNG, lưu Drive theo thư mục dự án. Lỗi thì ghi FAILED, không xuất ảnh.
- Cần đúng phút thì chuyển sang P3 (Cloud Scheduler). Đây là giới hạn đã biết, không hứa "đúng 06:00" trên Apps Script.

## ADR-006 — Quyền truy cập

- Quyền theo Google: chỉ người được chia sẻ file mới chạy menu. Vai trò (Admin / PM / DC / QAQC / Reviewer / Xem) lưu ở `TQN_CONFIG` và kiểm trong code trước mỗi thao tác ghi. Không dùng mật khẩu VBA làm bảo mật. Khóa API thời tiết để trong Script Properties, không để trong ô.
