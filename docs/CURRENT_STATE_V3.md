# G1 — Hiện trạng và phân loại module (bổ sung V3)
Tài liệu này bổ sung cho `G0_G1_HIEN_TRANG.md` và `G1_KEEP_REFACTOR_REPLACE.md` (commit 50360dd). Hai tài liệu đó vẫn còn hiệu lực. Ở đây em chỉ ghi phần thay đổi khi đặt hệ thống vào đích V3 (C# + PostgreSQL).

## 1. Code đang có (nguồn hiện hành)
| Module | File:dòng | Đang làm gì | Phân loại V3 | Lý do |
|---|---|---|---|---|
| Đọc Register theo tiêu đề | `apps-script/Read.gs:6-165` (`tqnFindCol_`, `tqnFindBlock_`, `tqnReadDetail_`, `tqnReadLog_`, `tqnExtractCodes_`) | Tìm cột theo tên khối và cột con, không cố định cột | **KEEP → port sang C#** | Đây là phần hiểu nghiệp vụ, đã test trên dữ liệu thật. C# sẽ dùng làm mapper của ImportBatch (G4), giữ nguyên quy tắc |
| Chuẩn hóa mã | `Config.gs:83-102` (`tqnCleanCode_`, `tqnDocKey_`, `tqnRev_`) | Bỏ tiền tố ngày và đuôi rev để ra khóa hồ sơ | **REFACTOR** | Đang chỉ hiểu Rev số. Phải tách thành `CodeProfile` có phiên bản (COASTAL_2026_V1) và hỗ trợ cả Rev A/B/C của bản vẽ |
| Nhập vào TQN_DATA | `Import.gs:61-111`, khóa ở `Import.gs:15` + `tqnImport_` | Upsert idempotent vào tab TQN_DATA | **REPLACE (ở G4)** bằng `import_batch` / `import_row_result` trong PostgreSQL. Tạm thời **KEEP** | TQN_DATA là một bản sao trên Sheet. Theo V3 không được để nó thành nguồn cạnh tranh. Khi PG chạy thì TQN_DATA chỉ còn là projection |
| QC 11 luật | `QC.gs:6-111` | Trùng mã, khoảng trắng, chưa vào sổ tổng, lệch tình trạng, sai công thức, quá hạn… | **KEEP → port** | Đây là luật nghiệp vụ. C# sẽ dùng chung bộ test fixture |
| Nhắc hạn | `Remind.gs:8-91` | Danh sách hạn ≤ N ngày, gửi email lúc 07:00 | **KEEP** đến G5, sau đó **REPLACE** bằng `Reminder`/`Notification` | Dùng tốt ngay bây giờ |
| Thời tiết | `Weather.gs:29-186` | Open-Meteo → tab "Theo dõi thời tiết", trigger 06:00 | **KEEP tạm** đến G8, sau đó **REPLACE** bằng Weather Engine C# (ADR-009) | Trigger Apps Script lệch khoảng ±15 phút. Chưa có ảnh PNG/PDF. Chưa xác minh giấy phép |
| Thanh bên | `Menu.gs:60-106`, `Sidebar.html` | KPI, lỗi, nhảy đến ô | **KEEP** (dùng ngay trong Sheet) | Web Dashboard của G6 là kênh thứ hai, không thay thanh bên |
| Kiểm thử | `test/run_tests.js`, `test/run_weather_tests.js` | Giả lập SpreadsheetApp | **KEEP** làm golden data | Trở thành fixture chung cho xUnit |
| VBA v1 in phiếu | `claude_19-11/` | In RFA/RFI/RFC/TRN từ mẫu Word | **KEEP** làm client Excel in phiếu | V3 giữ Excel VBA làm tiện ích |

## 2. Năm file `.bas` của MAX1 (baseline, chỉ tham khảo)
| File | Dòng | Phân loại | Bằng chứng |
|---|---|---|---|
| `modTQN_Install.bas` | 1-151 | **DEPRECATE** | Tạo các tab riêng của MAX1. Register đã có tab của mình. Không cài thêm vào workbook 32 tab |
| `modTQN_Import.bas` | 6-85 | **DEPRECATE** (lấy ý tưởng cho G4) | Dòng 21-24 cố định số cột (`tvgsCols = 13,13,13,13,12,…`), nên M0S lấy nhầm cột M thay vì N (lỗi M3). Không đọc Sổ trình duyệt (lỗi M4) |
| `modTQN_Data.bas` | 4-215 | **REPLACE** | Dòng 4-17 cấp ID bằng `max+1` trên Sheet: trùng khi nhiều người dùng, vi phạm chỉ thị III. Dòng 43 `TQN_FindDuplicate … Exit Function` bỏ qua dòng trùng mà không báo (lỗi M1) |
| `modTQN_Rules.bas` | 4-140 | **REFACTOR** (port các hàm thuần) | Dòng 92 map "Đã hoàn thành" thành CLOSED, sai với 7 trạng thái của Register (lỗi M2). `TQN_DocKey`/`TQN_Revision` có cùng ý tưởng với `Config.gs:91-102` |
| `modTQN_QC.bas` | 4-141 | **REFACTOR** phần QC, **REPLACE** phần Word | Dòng 113-128 dùng Word Find/Replace, bị giới hạn 255 ký tự (lỗi M5). Bộ in phiếu VBA v1 đã thay phần này |

## 3. Mười lỗi và rủi ro lớn nhất
| # | Mức | Lỗi / rủi ro | Hậu quả | Cách khắc phục |
|---|---|---|---|---|
| R1 | **P0** | Cấp số bằng `max+1`/COUNTIF trên Sheet (MAX1 `modTQN_Data.bas:4-17`, và cách cấp mã thủ công hiện nay) | Hai bên cùng nhập ra trùng mã hồ sơ đã phát hành | `number_sequence` + `next_number()` trong transaction. **Đã test 100 yêu cầu đồng thời, không trùng (T03)** |
| R2 | **P0** | MAX1 bỏ qua dòng trùng mà không báo (`modTQN_Data.bas:43`) | Mất hồ sơ khi import mà không ai biết | `import_row_result` bắt buộc ghi NEW/UNCHANGED/CONFLICT/UNMAPPED cho **mọi** dòng |
| R3 | **P0** | Sổ chi tiết có lỗi trong dữ liệu gốc: SDS-025/026 trùng mã (D3), SDS-001…004 và MSS-001 chưa vào sổ tổng (D4) | Import tự động sẽ gộp hoặc mất bản ghi | Không tự sửa. Import đánh dấu CONFLICT để người có quyền quyết định. QC Q01/Q05 đang báo các lỗi này |
| R4 | **P1** | Công thức `Q6 = P-M` sai ở RFA, RFC, Hồ sơ an toàn, SCH (D1) | Số ngày xử lý sai, báo quá hạn sai | Báo cho anh sửa trong Register (QC Q09). PG tự tính lại, không đọc ô công thức |
| R5 | **P1** | Đồng bộ hai chiều Sheets ↔ PG | Ghi đè lẫn nhau, mất dữ liệu | ADR-002: giai đoạn chuyển tiếp chỉ đồng bộ **một chiều theo từng module**. Không bao giờ có hai nơi cùng được ghi |
| R6 | **P1** | Khóa bản ghi S1 gồm `hash8(subject)` (`Import.gs:15`) | Sửa tên hồ sơ trong Register sẽ sinh thêm bản ghi mới trong TQN_DATA | G4 dùng `doc_key` + rev làm khóa. Đổi subject chỉ tạo diff để duyệt |
| R7 | **P1** | Giấy phép thời tiết: Open-Meteo bản miễn phí chỉ cho phi thương mại. Google Weather hạn chế lưu trữ nội dung | Vi phạm điều khoản khi lưu và phân phát PNG/PDF | ADR-009: chỉ bật nguồn đã xác minh giấy phép. Chưa xác minh thì không phát hành |
| R8 | **P1** | Chưa có hosting và .NET SDK (G0 BLOCKED) | Không đạt G3, phần 06:00 phải chạy khi máy tắt | Tạm giữ trigger Apps Script (đang chạy trên cloud Google). Anh chọn hosting theo ADR-007 |
| R9 | **P2** | S1 chưa chạy trên Google Sheets thật (mới chạy giả lập Node) | Có thể lệch API Google, quyền OAuth | Cài lên **bản sao** Register theo `apps-script/CAI_DAT.md` rồi chạy UAT S1 |
| R10 | **P2** | File Register có số điện thoại và email nhân sự | Lộ thông tin nếu đẩy lên GitHub public | Repo private. Fixture test bỏ thông tin cá nhân. `.gitignore` chặn `*.xlsx` gốc |

## 4. Ba phương án chuyển đổi
| | (A) Giữ Excel/Sheets làm chính | **(B) Lai, chuyển backend dần (chọn)** | (C) Viết lại có kiểm soát |
|---|---|---|---|
| Cách làm | Chỉ có Apps Script/VBA | Register Sheets vẫn là nơi các bên nhập. Mỗi module lần lượt chuyển nguồn chính sang PG: cấp số trước, rồi duyệt, rồi thời tiết | Dựng hệ thống mới, chuyển hẳn một lần |
| Ưu | Rẻ, các bên không phải đổi thói quen | Không gián đoạn, mỗi bước có test và đường lui | Kiến trúc sạch |
| Nhược | Không giải được R1/R5, khó dùng cho nhiều dự án | Phải giữ kỷ luật một chiều, chạy hai hệ trong một thời gian | Rủi ro mất dữ liệu cao, đổi thói quen các bên cùng lúc |
| Chi phí | Thấp | Trung bình, trải theo từng gate | Cao, dồn vào một lần |
| Kết luận | Chỉ phù hợp trước mắt | **Khuyến nghị, đúng với mặc định của chỉ thị** | Không chọn |
