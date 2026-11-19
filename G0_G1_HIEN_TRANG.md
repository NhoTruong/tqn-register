# G0 + G1 — Môi trường thật và kiểm kê hiện trạng

Ngày: 19/11/2026 10:20 (Asia/Ho_Chi_Minh) · Người lập: Claude · Bản: 0.1

## G0 — Môi trường (cái gì kiểm được, cái gì không)

| Hạng mục | Trạng thái | Bằng chứng |
|---|---|---|
| Máy chạy Claude | Linux trong cloud, Python 3, LibreOffice, Chrome headless, Playwright | đã dùng để render phiếu và test web |
| Máy anh Nho | Windows x64 (desktop-tildkva), đã nối qua Claude desktop | `get_device_info` 10:41 |
| Excel / Word Windows | **CHƯA KIỂM**. Chưa có thư mục nào được cấp quyền, chưa có shell trên máy anh | `connectedFolders: []` |
| Register live | Google Sheets, xem được qua trình duyệt tích hợp, chỉ đọc | đọc 32 tab bằng gviz lúc 09:10–09:45 |
| Register .xlsx trong gói | `HAUS QUANG NGAI- Register.xlsx`, 783 281 byte, SHA256 `47e537fe…59818f`, khớp manifest | `sha256sum` |
| Apps Script / Drive API / repo GitHub của anh | **KHÔNG CÓ QUYỀN**. Chưa nối Google Drive/Sheets connector, chưa có repo cho Register | — |
| Mã nguồn đang có | Chưa có repo cho Register. Đây là phiên đầu, toàn bộ nằm ở `/home/claude/build` và thư mục này | `MISSING_SOURCE` = không có repo cũ |

**Bản .xlsx trong gói cũ hơn bản live:** ô `Sổ trình duyệt!P23` (Link hồ sơ) trống trong .xlsx nhưng có tên file trên Google Sheets. Nguồn sự thật hiện tại là bản live.

## G1 — Kiểm kê 32 tab (đã sửa so với báo cáo 09:41)

Bộ đọc gviz trả về sheet đầu tiên khi không khớp tên, nên 7 tab (6 tab ẩn và `Lịch họp ` có dấu cách cuối tên) đã bị đọc nhầm lúc 09:41. Bảng dưới đây đọc trực tiếp từ file .xlsx.

| Nhóm | Tab | Ghi chú dữ liệu |
|---|---|---|
| Sổ tổng | **Sổ trình duyệt** | 42 dòng (B23:Q522). Công thức: K = I+J (hạn, theo ngày lịch); N = số ngày xử lý; **O = Tình trạng tự tính**: Kết quả → Chưa trình → Đã phản hồi → Quá hạn → Đang chờ. Thống kê C7:J19 đếm theo 13 loại. Ô thả xuống M (4 kết quả), H (TVGS/TVTK/CĐT/TVGS+CĐT/TBCONS nội bộ) |
| Sổ chi tiết có mã | SD (58), M0S (11), RFI (3), RFC (3), RFA (4) | Dữ liệu bắt đầu ở dòng 6. Cột khối TVGS/CĐT lệch nhau theo tab |
| Sổ chi tiết chưa có mã | OTS (16 mục "..."), MAT (4), SCH, CRP- Báo cáo, Hồ sơ an toàn, TRN-GNHS | khung sẵn, chưa phát hành |
| Văn bản | CV-TB đi, CV-TB đến, Chỉ thị – Phát sinh, Chỉ thị NT – Phát sinh, NCR | khung đủ cột, 0 dòng dữ liệu |
| Hiện trường | Theo dõi phát hành thiết kế, Hiện trường, Hao hụt bê tông (1 550 công thức), Theo dõi thời tiết (chỉ có tiêu đề) | |
| Nhân sự | DS BCH, DS Trực CN T7, Cập nhật kế hoạch công việc, Data | có SĐT, email cá nhân: không đưa lên GitHub |
| **Tab ẩn: mẫu của dự án trước** | SCH-Kế hoạch, OTS-MEP, Danh mục thẩm tra BPTC, Check, Cấp phối bê tông, DS Trực CN T4 | `Check!A1` = "Trung tâm hội nghị Phật Giáo Quốc Tế". Mã `TTHNPGQT-…`, ngày 2024–2025, nhà thầu Phước Thành. Register đang được **dùng lại từ dự án khác**: đúng lý do phải làm cấu hình theo dự án |
| Sót tên dự án cũ | `CV-TB đến` cột "Phước Thành phản hồi", `Theo dõi phát hành thiết kế` cột "BCH PHƯỚC THÀNH" | phải là tham số dự án |

## G1 — Lỗi nghiệp vụ trong Register (chỉ báo, không sửa)

| # | Mức | Chỗ | Lỗi |
|---|---|---|---|
| D1 | P0 | `RFA!Q6`, `RFC!Q6`, `Hồ sơ an toàn!Q6`, `SCH-Kế hoạch!Q6` (cả cột) | `=P-M`: lấy **ngày CĐT trừ "Số ngày phản hồi"** (M là số ngày, không phải ngày) → kết quả vô nghĩa. Đúng phải là `P-L` (từ TVGS) hoặc `P-I` (từ ngày trình) |
| D2 | P1 | `M0S!O23`/`S23` | Số ngày CĐT tính từ ngày trình (`R-I`), còn RFA tính từ TVGS: hai sổ đo khác nhau |
| D3 | P0 | SD dòng 30–33 | `SDS-025`, `SDS-026` mỗi mã dùng cho 2 bộ bản vẽ khác nhau (thang bộ / cột) |
| D4 | P1 | SD/M0S ↔ Sổ trình duyệt | SDS-001…004 và MSS-001 ghi "Đã trình duyệt 18/11" nhưng Sổ trình duyệt chưa có dòng (thống kê SD = 0, MS = 0) |
| D5 | P1 | `RFC!R4` và `RFA!R4` | Cùng một câu "CĐT đồng ý phê duyệt cọc thử…", không khớp nội dung hồ sơ (bể tự hoại / tiến độ) |
| D6 | P2 | `RFI!B6` | Khoảng trắng giữa mã: `261116- COA-…` |
| D7 | P1 | 2 hệ mã | Sổ trình duyệt dùng mã WF công ty `20261116.COT.TC38…`, sổ chi tiết dùng mã TBC `261116-COA-TBC-…`. Chỉ nối được qua cột Link hồ sơ (tên file PDF) |

## G1 — Lỗi của bộ MAX1 (có chạy mô phỏng làm bằng chứng)

Bằng chứng: `tools/dryrun_max1_import.py` chạy đúng logic `modTQN_Import` + `modTQN_Rules` trên bản sao .xlsx, chỉ đọc. SHA256 trước/sau không đổi. Kết quả: nhập 77 bản ghi, 16 phát hiện → `evidence/max1_dryrun_findings.csv`.

| # | Mức | file:dòng | Lỗi | Tác động dữ liệu |
|---|---|---|---|---|
| M1 | **P0** | modTQN_Data.bas:43 | Trùng Type+Key+Rev thì `Exit Function` im lặng | **Mất 2 dòng** SD!32, SD!33 (bản vẽ cột RE4/RE5) khỏi TQN_DATA; thông báo chỉ gộp vào "đã thấy/bỏ qua" |
| M2 | P0 | modTQN_Rules.bas:92 | "Đã hoàn thành" → `CLOSED` | 13 shop drawing công viên **vẽ xong nội bộ, chưa trình** bị xem như đã đóng |
| M3 | P0 | modTQN_Import.bas:23 | M0S lấy TVGS ở cột M ("Ghi trên hồ sơ" = chữ), ngày thật nằm ở N | Ngày TVGS của M0S luôn trống; ý kiến "Bổ sung biện pháp neo…" bị bỏ |
| M4 | P0 | toàn bộ | Không đọc **Sổ trình duyệt**, sổ duy nhất có Kết quả, Hạn, Ngày phản hồi | Dashboard MAX1 không biết hồ sơ nào đã duyệt / quá hạn thật |
| M5 | P1 | modTQN_QC.bas:120 | Word `Replacement.Text` giới hạn 255 ký tự | Nội dung dài gây lỗi 5854 |
| M6 | P1 | modTQN_QC.bas:124 | `MatchCase = False` + token chung chung | dễ thay nhầm chữ trong mẫu |
| M7 | P1 | 04_MAU_WORD | Mẫu dựng lại, **không logo, không đúng bố cục** dự án | không in nộp được |
| M8 | P2 | modTQN_Data.bas:17, Rules:123 | Tìm ID/tìm trùng duyệt từng ô | chậm khi > vài nghìn dòng |
| M9 | P2 | MsgBox chữ không dấu | VBA MsgBox không hiện được tiếng Việt có dấu | trải nghiệm |

Điểm tốt của MAX1 nên giữ: chỉ đọc tab gốc, `DocKey` bỏ ngày và rev, tách `RecordID` khỏi STT, `AUDIT`/`QC`, không đoán `DueDate`, file .bas thuần ASCII.
