# G1 — Ma trận KEEP / REFACTOR / REPLACE / DEPRECATE

Nguồn được xét: (A) bộ MAX1 của ChatGPT, (B) phần Claude đã làm sáng nay 19/11.

| Thành phần | Nguồn | Quyết định | Lý do / việc phải làm |
|---|---|---|---|
| Nguyên tắc chỉ đọc tab gốc, có hash trước/sau | A | **KEEP** | đúng P0-a |
| `TQN_CleanText`, `TQN_DocKey`, `TQN_Revision` (Rules.bas:1–60) | A | **KEEP** | đã chạy mô phỏng ra đúng; chuyển thành hàm dùng chung cho Apps Script (JS) và VBA, có test cùng bộ dữ liệu |
| `TQN_NormalizeStatus` (Rules.bas:79) | A | **REFACTOR** | bỏ map cứng; dùng bảng `status_vocabulary` theo dự án. "Đã hoàn thành" ≠ CLOSED (M2) |
| `TQN_AppendRecord` chống trùng (Data.bas:43) | A | **REFACTOR** | trùng thì ghi vào hàng chờ xử lý + báo rõ, không bỏ im lặng (M1) |
| Bảng cột cố định `tvgsCols/cdtCols` (Import.bas:23) | A | **REPLACE** | dò cột theo tiêu đề + bảng mapping theo dự án có version (M3) |
| Không đọc Sổ trình duyệt | A | **REPLACE** | Sổ trình duyệt là nguồn Kết quả/Hạn/Tình trạng; ghép với sổ chi tiết qua tên file (M4, D7) |
| `TQN_ExportWordForm` + 3 mẫu Word mới | A | **REPLACE** | dùng mẫu gốc của dự án có logo (B: `MAU_RFA/RFI/RFC/TRN.docx`), thay chữ theo đoạn, không giới hạn 255 ký tự (M5–M7) |
| `TQN_AUDIT`, `TQN_QC`, `TQN_CreateBackup` | A | **KEEP → REFACTOR** | giữ ý tưởng; với bản online, audit phải ghi ở backend, không chỉ trong sheet |
| `TQN_DASH` (đếm COUNTIF) | A | **DEPRECATE** | dashboard chuyển sang báo cáo R1–R5; tránh trùng logic với Sổ trình duyệt |
| Data dictionary 26 cột | A | **REFACTOR** | thêm `project_id`, `source_file_hash`, `raw_status`, `mapped_status`, `mapping_version`, tách `Submission` khỏi `Revision` |
| `report_demo.html`, ảnh mockup | A | KEEP làm tham chiếu | nhãn DEMO; không dùng số liệu |
| Kiểm mã theo QĐ COA (trùng mã, sai gói NP/HTKT, khoảng trắng) | B | **KEEP** | đã chạy trên 79 dòng thật, bắt được D3, D6; đưa vào bộ QC chung |
| Mẫu phiếu có ô `{{…}}` dựng từ mẫu gốc (RFA, RFI, RFC) + phiếu giao nhận TRN | B | **KEEP** | đã render PDF kiểm bố cục; giữ logo, khung |
| `TQN_TrinhDuyet.bas` v1 (đã giao 09:21) | B | **REFACTOR** | logic phiếu, giao nhận tốt; phải đổi nguồn sang Sổ trình duyệt và mapping theo dự án |
| `TQN_TrinhDuyet.src.bas` v2 (đang viết dở) | B | **ĐÓNG BĂNG** | chờ ADR-001; phần dùng lại: `U()` + `MessageBoxW` (tiếng Việt có dấu), `TimCot`, giao nhận gộp |
| Web mẫu giao diện (artifact 09:39) | B | **REFACTOR** | giữ bố cục; đổi sang 13 loại của Sổ trình duyệt, thêm `project_id`, nhãn DEMO |
| Trạng thái tự đặt 8 mức của Claude | B | **DEPRECATE** | dùng đúng 7 trạng thái Register đang có: Chưa trình, Đang chờ, Quá hạn, Đã duyệt, Duyệt có ý kiến, Không duyệt – trình lại, Đã phản hồi |
