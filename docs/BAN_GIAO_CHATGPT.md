# BÀN GIAO CHO CHATGPT – TQN Register (cập nhật 20/11 09:13)

## CẬP NHẬT MỚI NHẤT (09:13)
- **Đã đo bản nhanh cả nút trên DEV** (`tqnA_DoQuanLyNhanhDEV`, xem issue #1 comment 7028170988):
  - Cũ: **18,5 s**.
  - Nhanh, chưa cache (n=20): median **6,2 s**, p95 7,3 s.
  - Cache trúng (n=20): median **0,49 s**, p95 0,86 s.
  - So khớp JSON cũ/mới: `khop = true`, 0 chỗ lệch.
- **Kế hoạch nhanh hơn, anh Nho đồng ý xem xét, làm theo thứ tự:**
  1. Gộp `TQN_QUY_TRINH / PHIEU / GIAO_NHAN / AUDIT` + cấu hình vào cùng batchGet, cache `tqnConfig_`, khóa cache không gọi `tqnConfig_` → mục tiêu lần lạnh ~2 s (ước tính).
  2. Nạp dữ liệu từ cache ngay khi mở hộp thoại, bỏ 1 vòng khứ hồi; nút Làm mới vẫn đọc thật.
  3. Trigger 5 phút hâm sẵn cache, onEdit đổi `tqnDv`.
  Mục tiêu: bấm nút → thấy bảng < 1 s gần như mọi lần. Đo 20 lần trên DEV, anh Nho duyệt mới lên PROD.
- **V6 journal (issue #2):** Claude đã đồng ý 5 lỗ hổng ChatGPT nêu. Acceptance test có 5 ca `it.fails` (commit 681efc4). Việc tiếp: ADR-0004 rồi code journal; ChatGPT review và viết test đọc xen giữa.


Claude sắp hết phiên. ChatGPT làm tiếp trên repo này theo `PHOI_HOP.md`. Mọi trao đổi ghi ở issue #1, #2.

## 1. Trạng thái thật (đã kiểm / chưa kiểm)

### Nút Quản lý hồ sơ (issue #1, mục 9), mã cũ Apps Script
- **Đo trên bản sao DEV** (spreadsheet `[DEV] TQN Register - ban thu`, script project `1g47zWMx_BOBOm1g4_48uHT05Rvx1mlKNOeEl_3R4yvJtO3TCLGvtXS1G`, runtime **V8**, đã bật Sheets advanced service):
  - `tqnManagerData(true)` hiện tại: median **18,3 s**, p95 19,6 s (n=5). Chậm nhất là `tqnQlTheoMa_` (6–9 s) và `tqnDocsData` (3–4 s).
  - `tqnQlTheoMa_` (n=20):
    - B0 gốc: 6 473 / 8 398 ms (median / p95)
    - B1 `getDataRange` mỗi tab: 2 585 / 3 207 ms
    - **B2 `Sheets.Values.batchGet` 1 lệnh: 816 / 1 209 ms**
    - **A cache trúng: 114 / 210 ms**
    - Cả 3 cách đều khớp 68/68 mã.
  - `tqnDocsData`: chậm do **số lần gọi Sheets** (mỗi lần 130–230 ms), không phải do CPU. Kết quả 2 hồ sơ là **đúng**, chỉ có 2 NCR. **Lỗi tiềm ẩn:** tab "Chỉ thị" dò cột Nội dung ra cột A (STT).
- **Đã viết, CHƯA CÓ KẾT QUẢ:** `apps-script/thi_nghiem/QuanLyNhanh.gs`, gồm `tqnManagerDataNhanh` (1 batchGet cho mọi tab, có dự phòng B1), `tqnManagerDataCache` (cache theo `tqnDv`, TTL 120 s) và hàm đo `tqnA_DoQuanLyNhanhDEV` (so khớp sâu JSON cũ/mới + 20 lần không cache + 20 lần cache trúng). Đã dán vào bản DEV và bấm chạy lúc **00:07**. **Kết quả xem ở trang Executions của project DEV.** Việc đầu tiên: đọc kết quả, dán vào issue #1.
- **Chưa làm:**
  - Đo trọn vòng từ lúc bấm nút đến khi hộp thoại dùng được (điều kiện 3 của ChatGPT).
  - Cơ chế đổi `tqnDv` khi sửa sổ (onEdit + nút Làm mới bỏ qua cache).
  - Sửa dò cột tab Chỉ thị.
  - Mục 12: bật Sheets v4 trong manifest PROD.
  - **CHƯA ĐỤNG PROD.**
- **Cảnh báo vận hành:** trong trình soạn Apps Script, hàm mặc định đứng đầu danh sách là `tqnDailyJob`. Claude đã bấm nhầm 2 lần trên bản DEV (không gửi email, `gui_mail: 0`). **Phải nhìn tận mắt ô chọn hàm trước khi bấm Chạy.**

### Register V6 (issue #2), TypeScript + React, thư mục `v6/`
- 35 test PASS (`npm test`). Bản mới nhất (commit 9dfc7b9 trở đi) **chưa nạp DEV web app**, đang chờ chạy `cap_nhat_dev.bat` trên máy anh Nho.
- **P0 đã xác nhận bằng test** `v6/tests/loi-giua-chung.spec.ts`: lỗi Sheets giữa chừng khi tạo hồ sơ → bấm thử lại sinh **2 hồ sơ / 2 mã**. Lỗi khi submit → thử lại báo lỗi dù dữ liệu đã ghi.
- **Đã thống nhất hướng sửa** (issue #2): journal `TQN6_OPS` PREPARED → APPLYING → COMMITTED, ID xác định theo idem, khôi phục trước mỗi lệnh ghi; sổ bản chiếu có trạng thái `SYNC_*`; khóa cache gồm storeId + projectId + loại + dv. **Chưa code.**
- Chờ ChatGPT trả lời 2 câu: (1) hai request cùng khôi phục một OPS; (2) docNo đã cấp ở PREPARED nhưng thao tác bị bỏ.

## 2. Việc tiếp theo, theo thứ tự (cập nhật `docs/VIEC_CON_LAI.md`)
1. Lấy kết quả `tqnA_DoQuanLyNhanhDEV` → issue #1. Nếu `khop=false`, sửa chỗ lệch (thường do định dạng ngày/số của FORMATTED_VALUE).
2. Đo trọn vòng nút trên DEV (≥20 lần, cache trúng/trượt). Đạt median < 3 s thì đề xuất PR đưa lên PROD **kèm sao lưu `prod-baseline` mới**, chờ anh Nho duyệt.
3. V6 P0: viết journal + khôi phục trong `v6/src/core/service.ts` / `repo.ts` / `sheetsRepo.ts`. Claude đang giữ các file này; ChatGPT muốn làm thì ghi vào issue #2 để nhận việc. Bảng fault injection phải ra 0 hồ sơ trùng.
4. ADR-0003: nguồn chuẩn theo giai đoạn, kế hoạch nhập dữ liệu (mục 33).
5. Mục ChatGPT đã nhận: 8, 11, 12, 15, 31.

## 3. Cách kiểm
- `cd v6 && npm ci && npm run typecheck && npm test && npm run build`
- Mã cũ: `node test/run_quanly_test.js`
- Không triển khai PROD, không push đè `main` khi chưa test, không đưa key/mật khẩu vào repo.

## 4. Đường dẫn
- Issue #1: https://github.com/NhoTruong/tqn-register/issues/1
- Issue #2: https://github.com/NhoTruong/tqn-register/issues/2
- Toàn bộ việc còn lại: `docs/VIEC_CON_LAI.md`; đặc tả giao việc: `docs/GIAO_CHATGPT.md`; quy trình: `PHOI_HOP.md`; ADR: `docs/adr/`.
