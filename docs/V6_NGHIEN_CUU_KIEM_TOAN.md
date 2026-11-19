# TQN Register V6 — Nghiên cứu GitHub + Kiểm toán (19/11/2026)

## 1. Chín repo đã đọc (clone thật về máy làm việc, đọc file cấu hình + mã chính)

| Repo | Giấy phép | Commit cuối | Là gì | Build/test | Phù hợp TQN | Dùng thế nào |
|---|---|---|---|---|---|---|
| google/clasp | Apache-2.0 | 28/08/2026 | CLI chính thức đẩy/kéo code, deploy, nhiều môi trường | npm, mocha | **Rất cao** | Công cụ bắt buộc: push code, `.clasp.dev.json` / `.clasp.prod.json` |
| Mikoshiba-Kyu/gas-webapp-boilerplate | MIT | 26/05/2026 | React + TS + Vite → 1 file HTML; backend TS → rollup ra .gs; vitest + Playwright | yarn/npm, vitest, playwright | **Cao nhất cho Web App React** | **CHỌN LÀM STARTER DUY NHẤT** (lấy cấu hình build, không lấy code mẫu) |
| google/aside | Apache-2.0 | 18/06/2026 | Bộ khung TS chính thức: jest, eslint, tách dev/prod | npm, jest | Cao | Chỉ học ý tưởng tách DEV/PROD bằng 2 file clasp; không trộn khung |
| YunosukeYoshino/google-apps-script-starter | MIT | 18/08/2026 | Khung CLI dùng bun + oxlint + vite singlefile | bun | Trung bình | Không dùng (cần bun, không có React) |
| googleworkspace/apps-script-samples | Apache-2.0 | 30/07/2026 | Mẫu chính thức Gmail/Drive/Sheets/trigger | — | Tham khảo | Tra cách gọi Gmail/Drive cho Smart Intake |
| tanaikech/taking-advantage-of-google-apps-script | MIT | 08/11/2026 | Tổng hợp mẹo hiệu năng (batch, Sheets API) | — | Tham khảo | Nguyên tắc batch read/write |
| tanaikech/TableApp | MIT | 28/12/2025 | Thư viện bọc "Tables" của Sheets qua Sheets API | — | Thấp lúc này | Chưa dùng (sổ hiện tại không phải Table) |
| tyclaudius-ai/sheets-approval-appsscript | MIT | 20/02/2026 | 1 file Code.gs duyệt Pending/Approved/Rejected, LockService, hash để buộc duyệt lại khi sửa | — | Trung bình | Học ý: **sửa nội dung sau khi duyệt → bắt duyệt lại** (hash dòng) |
| gus-automates/invoice-approval-tracker | **Không có file giấy phép** | 29/03/2026 | Gmail → Drive → sheet hóa đơn | — | Thấp | **Không sao chép** (mặc định giữ bản quyền); chỉ đọc ý luồng Gmail→Drive |

Kết luận: 1 starter = **gas-webapp-boilerplate** (React/TS/Vite/vitest/Playwright, MIT) + **clasp** làm công cụ. Không trộn aside hay starter khác.

## 2. Nguồn đã đối chiếu

- MAX1 (VBA, ChatGPT): 5 module `.bas` + Excel starter — có trên máy làm việc.
- V2 (handoff ChatGPT): chỉ thị, kiến trúc, mẫu báo cáo — có.
- V3 (C# + PostgreSQL): chỉ thị + `db/schema_v0.sql` + G0–G2 Claude làm sáng 19/11 — có.
- **V4, V5: không tìm thấy trong các file anh đã gửi.** Bảng dưới không đánh giá V4/V5; anh gửi thì em bổ sung.
- Code Apps Script đang chạy thật trên Register (4.832 dòng, 25 file) — đã đọc.

## 3. Bảng KEEP / REFACTOR / REPLACE / DEPRECATE

| Thành phần | Nguồn | Quyết định | Lý do |
|---|---|---|---|
| Khối TQN QUẢN LÝ: chọn phiếu → tự Tình trạng/ngày, Ai đang giữ, trễ hạn, rev cũ (QuanLy.gs) | Apps Script | **KEEP** | đã chạy đúng trên sheet thật 19/11, 62 test PASS |
| Nhật ký TQN_NHAT_KY khóa người khác | Apps Script | **KEEP → REFACTOR** | đổi sang sự kiện có `event_id`, `doc_id`, `idem_key` (Audit Engine V6) |
| Mẫu phiếu A/B/C/D theo dự án (TQN_MAU_PHIEU) | Apps Script | **KEEP** | đây chính là cấu hình workflow đa dự án; V6 đọc lại cùng bảng |
| Tự cấp mã SD (TuDongMa.gs) | Apps Script | **REFACTOR** | chuyển thành Numbering Engine thuần TS, có khóa + chống trùng + test |
| Hộp Quản lý hồ sơ / Workflow (Manager, Workflow .html) | Apps Script | **REPLACE dần** | thay bằng Web App React V6; giữ chạy song song tới khi V6 đủ chức năng |
| Web.html (web cập nhật thử) | Apps Script | **DEPRECATE** | trùng với V6 React |
| Thời tiết (Weather.gs) | Apps Script | **KEEP** | đang chạy thật; V6 chỉ bọc vào Weather Engine |
| AI đọc phiếu / mẫu phiếu (Ai.gs) | Apps Script | **KEEP (chỉ đề xuất)** | AI không tự duyệt; giữ đúng nguyên tắc |
| Gọn sổ, hàng nút, ẩn tư vấn | Apps Script | **KEEP** | thao tác giao diện sheet, không phải nghiệp vụ |
| `TQN_DocKey`, `TQN_Revision`, chuẩn hóa chữ | MAX1 | **KEEP (port sang TS)** | đúng; dùng cho Classification |
| Cột cố định tvgsCols/cdtCols | MAX1 | **REPLACE** | đã thay bằng dò tiêu đề (Read.gs) |
| Dùng số dòng làm khóa | Apps Script cũ | **REPLACE** | V6 dùng `document_id` (ULID) + `revision_no`; số dòng chỉ là vị trí hiển thị |
| ADR-003 (document/revision/submission) | V3 | **KEEP** | V6 làm đúng mô hình này |
| schema_v0.sql (PostgreSQL) | V3 | **KEEP làm đích** | adapter V6 cùng tên trường → chuyển ASP.NET Core/PostgreSQL không đổi nghiệp vụ |
| Dashboard COUNTIF (TQN_DASH) | MAX1 | **DEPRECATE** | Report Engine thay |

## 4. Quyết định kiến trúc V6

- **Một hệ nghiệp vụ duy nhất**: `core/` viết TypeScript thuần (không gọi Google). Apps Script chỉ là adapter (Sheets/Gmail/Drive/Lock). Sau này ASP.NET Core chỉ thay adapter, hoặc port `core/` với cùng bộ test JSON.
- Định danh: `document_id` ULID; mã phát hành lấy từ Numbering Engine theo cấu hình dự án; `revision_no` tăng; mọi lệnh ghi có `idem_key` (gửi lại không tạo bản thứ hai).
- Lưu trữ giai đoạn 1: tab `TQN6_*` (DOCS, REVS, EVENTS, COUNTERS, IDEM, FILES) trong **file DEV riêng** — không đụng sổ gốc.
- Môi trường: `.clasp.dev.json`, `.clasp.staging.json`, `.clasp.prod.json`; Web App `executeAs: USER_ACCESSING`, `access: DOMAIN`/`MYSELF` — **không bao giờ ANYONE_ANONYMOUS**.
- AI: chỉ phân loại/trích xuất/soạn nháp; workflow chặn mọi quyết định duyệt không có người dùng vai trò Reviewer.
