# Cài TQN Register (bản 1 file) – khoảng 5 phút

**Cài vào BẢN SAO trước.** Duyệt kết quả xong mới cài vào file các bên đang nhập.

1. **Tạo bản Google Sheets.** Trên Drive, mở Register và chọn **Tệp → Lưu dưới dạng Google Trang tính**. Nếu đã là Google Sheets thì chọn **Tệp → Tạo bản sao**.
2. **Gắn code.** Trong bản Google Sheets vừa tạo, chọn **Tiện ích mở rộng → Apps Script**. Xóa hết nội dung trong `Code.gs`, dán **toàn bộ** file `TQN_Register_1_FILE.gs` vào, rồi bấm **Lưu**. Không cần tạo file HTML.
3. **Tải lại sheet.** Trên thanh menu sẽ có **TQN Register**.
4. **Cấp quyền lần đầu.** Chọn **TQN Register → ▣ Quản lý hồ sơ**. Google hỏi quyền (bảng tính, Drive để lưu PDF phiếu, gửi mail nếu bật nhắc), anh tự bấm duyệt bằng tài khoản của anh.
5. **Bật tự động.** Chọn **TQN Register → ★ Bật tự động tất cả**. Tool sẽ cập nhật mỗi giờ, nhắc hạn mỗi sáng, đổi màu giao diện và màu thanh tab theo nhóm.

## Dùng hằng ngày

Mở **▣ Quản lý hồ sơ**:

- **Ba con số** cho biết Quá hạn, Sắp hạn, Đang chờ.
- **Ba nút chính:**
  1. Nhập và kiểm
  2. Tạo phiếu
  3. Giao nhận
- **Việc của từng hồ sơ:** đánh dấu ô vuông đầu dòng rồi bấm nút, hoặc bấm nút "việc tiếp theo" ngay trên dòng đó.
- **Phiếu:** mỗi phiếu được ghi vào tab `TQN_PHIEU` và lưu PDF vào thư mục Drive `TQN_PHIEU`. Phiếu giao nhận được cấp số `yymmdd-COA-TBC-TML-nnn`, không trùng số.
- **Thông tin trên phiếu** (công ty, người lập, số hợp đồng…): sửa ở tab `TQN_CONFIG`, các dòng `phieu_*`.

## Bật trợ lý AI (tùy chọn)

1. Lấy khóa API Claude của công ty.
2. Trong Apps Script, chọn **Cài đặt dự án → Thuộc tính tập lệnh**. Thêm thuộc tính `ANTHROPIC_API_KEY` và dán khóa vào đó. **Không dán khóa vào ô bảng tính, không gửi qua chat.**
3. Trong `TQN_CONFIG`, đặt `ai_enabled = Có`.
4. Mở **Quản lý hồ sơ → ✦ Đọc phiếu (AI)**, chọn ảnh hoặc PDF phiếu đã có ý kiến. Xem bảng so sánh rồi bấm **Ghi**. Tool chỉ ghi vào ô trống, không đè dữ liệu có sẵn.

## An toàn dữ liệu

- Tool chỉ ghi vào các tab tên `TQN_*`.
- Ngoại lệ 1: dòng thời tiết, ghi vào tab "Theo dõi thời tiết".
- Ngoại lệ 2: nút Đọc phiếu → Ghi. Nút này chỉ ghi vào ô trống, có xác nhận và có nhật ký `TQN_AUDIT`.
- Đổi màu chỉ đổi phần hiển thị. Bản gốc lưu ở tab ẩn `TQN_GIAO_DIEN_GOC`. Chọn **TQN Register → Trả lại giao diện gốc** để trả lại cả màu ô lẫn màu tab.

## Gỡ bỏ

**TQN Register → Tắt tự động tất cả**, rồi xóa code trong Apps Script. Các tab `TQN_*` có thể xóa tay. Tab gốc không bị ảnh hưởng.
