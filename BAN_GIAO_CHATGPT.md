# Bàn giao cho ChatGPT – TQN Register (19/11/2026)

## 1. Đây là gì
- TQN Register là sổ quản lý hồ sơ trình duyệt (SD, RFA, RFC, RFI…) của dự án Coastal Quảng Ngãi, chạy trên **Google Sheets + Apps Script**. Mọi người dùng chung một file.
- Phiếu đã ký lưu trên **OneDrive**. Register chỉ giữ link tới phiếu.
- Mã nguồn nằm trong `apps-script/*.gs`. Test chạy bằng Node trong `test/`: `node test/run_quanly_test.js`, tương tự cho các file test khác.
- Dữ liệu thật dùng cho test: `test/fixture_live.json`. Ô ngày ghi dạng `{"__d":"2026-11-06T00:00:00"}`.
- Build gói cài: `python3 tools/build_dist.py`. Có file mới thì thêm tên vào mảng `GS` trong file này.

## 2. Luật bắt buộc (không được vi phạm)
1. **Chỉ ghi vào tab có tiền tố `TQN_`**, và tạo tab bằng `tqnOwnSheet_()`.
   - Tab gốc SD, RFA… chỉ được ghi vào khối "TQN QUẢN LÝ" bên phải. Riêng tự cấp mã thì được ghi, nhưng chỉ vào ô Mã đang trống.
   - Không sửa dữ liệu gốc khi chưa hỏi.
2. **Không chặn cứng người dùng.** Gặp sai sót thì cảnh báo, tô đỏ, ghi nhật ký. Ngoại lệ duy nhất: cột kết quả phiếu và cột link được khóa theo email (`TQN_PHAN_QUYEN`).
3. **Khóa API chỉ để trong Script Properties** (`ANTHROPIC_API_KEY`), không bao giờ để trong ô hay trong code.
4. **Tách logic thành hàm thuần** (không gọi `SpreadsheetApp`), đặt tên kiểu `tqnXxx_`, để test bằng Node và sau này dùng lại cho web app.
5. **Không xóa file trên Drive.** Không gửi mail khi người dùng chưa bật.
6. **Mọi chữ hiện cho người dùng viết tiếng Việt có dấu**, ngắn gọn.
7. Mỗi hàm mới phải có test trong `test/run_<ten>_test.js`, viết theo mẫu `test/run_quanly_test.js` (mock sheet + `vm`). **Toàn bộ test cũ phải vẫn PASS.**

## 3. Đã có (đừng làm lại)
- `QuanLy.gs`:
  - Tab mẫu phiếu `TQN_MAU_PHIEU`, ghi A/B/C/D đúng nguyên văn phiếu dự án.
  - Khối TQN QUẢN LÝ gồm: kết quả từng cấp | Chuyển tới | Ai đang giữ | Link OneDrive | Kiểm tra.
  - Nhật ký `TQN_NHAT_KY`, ghi ở hàm `tqnQuanLyKhiSua`.
  - Phân quyền theo email; AI đọc mẫu phiếu trắng.
- Các tính năng khác:
  - `TuDongMa.gs`: tự cấp mã.
  - `ThaFile.gs`: thả file.
  - `TienDo.gs`: hồ sơ chặn thi công.
  - `Weather.gs`: thời tiết.
  - `Remind.gs`: nhắc hạn.
  - `TrangChu.gs`: trang chủ.
  - `Ai.gs`: AI đọc hồ sơ, có hàm `tqnAiCall_(cfg, prompt, maxTokens)`.
  - `Auto.gs`: bật/tắt tự động; danh sách trình kích hoạt ở `TQN_AUTO_HANDLERS`.
- Các hàm dùng chung:
  - `tqnConfig_()`, có các trường `detailSheets`, `alias`, `hdrFrom`, `hdrTo`, `firstRow`.
  - `tqnDetailMap_(hdr)`: dò cột theo tiêu đề.
  - `tqnNoAccent_`, `tqnIsDate_`, `tqnAudit_`, `tqnSs_`.

## 4. Việc giao cho ChatGPT (làm theo thứ tự, xong việc nào báo việc đó)

### Việc G1 – Xuất PDF tổng hợp các mã vừa cập nhật
- Đầu vào: khoảng thời gian, mặc định là từ lần xuất trước đến hiện tại (lưu trong DocumentProperties `lastPdf`).
- Đọc `TQN_NHAT_KY` và gom theo **mã hồ sơ**. Mỗi mã lấy trạng thái mới nhất: TVGS, CĐT, Ai đang giữ, link OneDrive, người sửa cuối, giờ.
- Ghi ra tab `TQN_BAO_CAO_CAP_NHAT`:
  - Tiêu đề: tên dự án, khoảng thời gian, người xuất.
  - Bảng các mã đã cập nhật, chia nhóm theo loại hồ sơ.
- Xuất tab đó ra PDF khổ A4 ngang qua URL `https://docs.google.com/spreadsheets/d/<id>/export?format=pdf&gid=<gid>&portrait=false&fitw=true…`, có kèm `ScriptApp.getOAuthToken()`.
- Lưu PDF vào thư mục Drive `TQN Register – Hồ sơ/Báo cáo`. Hàm `tqnThaFolders_()` có sẵn cách tạo thư mục.
- Thêm mục menu "🖨 Xuất PDF các mã vừa cập nhật".
- Hàm gom phải là hàm thuần, có test: nhiều lần sửa cùng một mã thì lấy lần cuối; dòng CẢNH BÁO vẫn được liệt kê; khoảng thời gian rỗng thì báo "Không có cập nhật".

### Việc G2 – Tự kiểm sức khỏe hằng tuần
- Hàm `tqnKiemSucKhoe()` chạy mỗi thứ Hai lúc khoảng 07:10:
  - Kiểm đủ các trình kích hoạt trong `TQN_AUTO_HANDLERS`, thiếu cái nào thì **tự cài lại**.
  - Kiểm `TQN_MAU_PHIEU` không có lỗi, dùng `tqnQlDocMau_(rows)._loi`.
  - Đếm số dòng có "⚠" ở cột Kiểm tra của từng sổ.
  - Kiểm nhật ký vẫn còn khóa.
- Ghi kết quả vào `TQN_SUC_KHOE` (mỗi tuần một dòng) và vào `tqnAudit_`.
- Phần quyết định "cái nào thiếu, cần cài lại" phải là hàm thuần, có test.

### Việc G3 – Tab hướng dẫn cho đội
- Tạo tab `TQN_HUONG_DAN` gồm:
  - 5 bước cho người mới: chọn kết quả theo phiếu, dán link OneDrive, xem Ai đang giữ, sửa khi Kiểm tra báo ⚠, xem nhật ký.
  - Bảng ý nghĩa các ký hiệu ✔ ↩ ⏳ ✎ ⚠.
- Thêm mục menu để mở tab này.

## 5. Cách báo cáo lại
- Liệt kê: các file đã đổi, kết quả test (số PASS/FAIL của **mọi** file test), output của `build_dist.py`.
- Ghi rõ 4 trạng thái: Đã thiết kế / Đã viết code / Đã test Node / **CHƯA KIỂM trên sheet thật**.
- Không tự nạp lên sheet. Claude sẽ soát code rồi mới nạp.
