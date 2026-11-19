# G2 — Kế hoạch triển khai có UAT

Nguyên tắc: một bước xong và có bằng chứng thì mới sang bước sau. Mọi thứ chạy trên **bản sao** Register trước.

| Sprint | Nội dung | Đầu ra | UAT phải PASS (ID trong `UAT_UNIVERSAL.csv`) | Cần từ anh |
|---|---|---|---|---|
| S1 (G3) | Lớp đọc chuẩn + mapping theo tiêu đề cho 11 tab kiểu trình duyệt + Sổ trình duyệt; import idempotent ra `TQN_DATA`; QC (D1–D7, trùng mã, ghép 2 hệ mã); cấu hình `COASTAL_2026_V1` + `DEMO_B` | Apps Script + test Node chạy trên dữ liệu thật đã xuất; báo cáo dry-run | Không sửa nguồn (hash), Alias M0S/MAT, Import idempotent, Revision, Source trace, Isolated project scope | 1 bản sao Register trên Drive (anh tạo hoặc cho em quyền), quyền chạy Apps Script trên bản sao |
| S2 (G4) | Tạo phiếu RFA/RFI/RFC/TRN từ mẫu gốc → Google Docs/PDF trong Drive; cấp mã mới khóa `LockService`; trình lại rev+1 | phiếu PDF mẫu đối chiếu bản Word gốc | Mẫu Word RFI, Tình trạng B/C/D, Uniqueness sequence, Due date missing | xác nhận phiếu PDF đạt hay phải in bằng Word |
| S3 (G5) | Báo cáo R1 Executive + R2 Document Control (HTML → PDF/PNG), nhãn nguồn/thời điểm | 2 báo cáo từ dữ liệu thật | KPI revision denominator, Báo cáo provenance, Responsive, Font tiếng Việt | **6 ảnh báo cáo tham chiếu** (chưa có trong gói lẫn trong cuộc trò chuyện này) |
| S4 (G6) | Thời tiết 06:00: GPS đã xác nhận, nhà cung cấp, trigger, lưu Drive, ghi `Theo dõi thời tiết` | ảnh R3 thật, log 3 ngày liên tiếp | Weather requires GPS, 2 cities, Time slot, Idempotency, Provider outage, Forecast vs observed | **tọa độ công trường đã xác nhận**; chọn nhà cung cấp và giấy phép |
| S5 | Module VBA in hàng loạt bằng mẫu Word gốc (client phụ) | .bas + hướng dẫn | Compile, in thử 3 phiếu (chạy trên máy anh) | chạy thử trên Excel Windows |
| S6 (G8–G9) | Kiểm thử tổng, sao lưu/khôi phục, bàn giao có version + SHA256 | gói phát hành | toàn bộ P0 | nghiệm thu |

**Chưa làm ở bước này:** chưa viết Apps Script, chưa đụng Register live, chưa gọi API thời tiết.
