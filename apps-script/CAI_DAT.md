# Cài TQN Register S1 vào bản sao Register (khoảng 5 phút)

**Làm trên BẢN SAO trước.** Không cài thẳng vào file các bên đang nhập cho tới khi anh duyệt kết quả.

1. **Tạo bản Google Sheets.** Mở `HAUS QUANG NGAI- Register.xlsx` trên Drive, chọn **Tệp → Lưu dưới dạng Google Trang tính**. Đặt tên ví dụ `Register COA – TQN thử`. File .xlsx gốc giữ nguyên.
   - Bước này bắt buộc: file đang mở ở chế độ .xlsx thì không gắn được Apps Script, cũng không khóa vùng được.
2. **Gắn code.** Trong bản Google Sheets vừa tạo, chọn **Tiện ích mở rộng → Apps Script**.
   - Xóa nội dung mẫu, rồi tạo 7 file đúng tên và dán nội dung tương ứng: `Config.gs`, `Read.gs`, `Import.gs`, `QC.gs`, `Remind.gs`, `Weather.gs`, `Menu.gs`.
   - Tạo thêm file HTML tên `Sidebar` và dán `Sidebar.html`.
   - Bấm **Lưu**.
3. **Mở menu.** Quay lại sheet, tải lại trang. Trên thanh menu sẽ có **TQN Register**.
4. **Cấp quyền lần đầu.** Chọn **TQN Register → Cài đặt**. Google hỏi quyền, anh tự duyệt; script chạy bằng tài khoản của anh.
5. **Điền cấu hình.** Mở tab `TQN_CONFIG`, sửa các ô nền vàng nếu cần. Mặc định đã điền sẵn cho Coastal.
6. **Chạy thử.** Chọn **TQN Register → 1. Nhập & kiểm tra ngay**, rồi mở **Bảng điều khiển**.
7. **Bật nhắc hạn hằng ngày.**
   - Chọn **Bật nhắc hạn hằng ngày**. Mặc định chạy khoảng 7h; Google có thể lệch ±15 phút.
   - Muốn gửi mail thì sửa `TQN_CONFIG`: `remind_email_enabled = Có` và điền `remind_recipients`.

8. **Thời tiết 06:00.**
   - Điền trong `TQN_CONFIG`: `weather_lat`, `weather_lon` (tọa độ thật của công trường), `weather_gps_confirmed = Có`, `weather_license_ok = Có` (sau khi đã kiểm điều khoản nguồn dữ liệu), `weather_enabled = Có`.
   - Chọn **TQN Register → Thời tiết: cập nhật ngay** để thử, rồi **Thời tiết: bật 06:00 hằng ngày**.
   - Script chạy trên máy chủ Google: máy tắt, không ai mở file vẫn chạy.
   - Script ghi 1 dòng/ngày vào tab "Theo dõi thời tiết" đúng cột có sẵn. Không ghi đè dòng người nhập tay. Lỗi API thì không ghi số.

## Phần mềm thêm vào bản sao

Thêm 6 tab `TQN_*`. Tab gốc duy nhất được ghi là "Theo dõi thời tiết" (đã test U01, W13):

| Tab | Nội dung |
|---|---|
| `TQN_CONFIG` | cấu hình dự án |
| `TQN_DATA` | bảng chuẩn hóa, nhập lại bao nhiêu lần cũng không trùng dòng |
| `TQN_QC` | lỗi và cảnh báo, chỉ báo, không tự sửa |
| `TQN_NHAC_HAN` | danh sách nhắc hạn |
| `TQN_AUDIT` | nhật ký mỗi lần chạy |
| `TQN_THOI_TIET_RAW` | bản gốc phản hồi API thời tiết, cả lần lỗi |

## Gỡ bỏ

- Tắt nhắc hạn: **TQN Register → Tắt nhắc hạn hằng ngày**.
- Xóa 5 tab `TQN_*` và project Apps Script.
- Register không bị ảnh hưởng.
