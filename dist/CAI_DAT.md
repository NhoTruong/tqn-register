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

## Thời tiết công trường (đã bật trên file thật 19/11)

- **Tự động 4 lần/ngày: 6h, 11h, 13h, 18h.** Mỗi lần ghi lại đúng dòng hôm nay trong tab "Theo dõi thời tiết". Đổi giờ ở ô `weather_hours` trong `TQN_CONFIG`.
- **Mỗi dòng gồm:**
  - các cột gốc A–J;
  - ô Ghi chú: dòng đầu "📷 Xem bản tin hình" (bấm mở), sau đó mỗi buổi Sáng / Trưa / Chiều / Tối một dòng, rồi gió giật, mặt trời và ngày mai;
  - 4 ô biểu tượng Sáng 06–10h, Trưa 11–13h, Chiều 14–17h, Tối 18–23h.
- **Bản tin hình (minh chứng):** PDF chi tiết từng giờ 06:00–23:00, lưu ở Drive, thư mục "TQN Register – Bản tin thời tiết".
- **Chạy bù:** chọn **TQN Register → Thời tiết: cập nhật từ ngày… đến ngày…** (vd `01/11/2026 - 19/11/2026`). Ngày đã qua dùng số liệu mô hình Open-Meteo.
- **Dòng nhập tay** (ô Ghi chú không bắt đầu bằng 📷) không bao giờ bị ghi đè.
- **Tọa độ:** `weather_lat` / `weather_lon` = 15.099038 / 108.894552 (anh Nho gửi 11:12 19/11).
- **Nguồn:** Open-Meteo bản miễn phí, chỉ cho phép dùng phi thương mại. `weather_license_ok = Có` do anh Nho chốt.

## Tốc độ các nút

- **Bảng điều khiển, Quy trình hồ sơ:** lần đầu 15–20 giây, bấm lại trong 3 phút gần như mở ngay. Nút Làm mới luôn đọc số mới.
- **Nhập & kiểm tra, Xem nhắc hạn:** 15–20 giây, vì phải đọc 32 sổ.
- **Giao diện dễ nhìn:** 1–2 phút, chỉ cần bấm 1 lần.
- **Thời tiết theo khoảng ngày:** khoảng 8 giây mỗi ngày, vì mỗi ngày tạo 1 bản tin PDF.

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

## Trang chủ (thanh nút) – 19/11
Menu TQN Register → "⌂ Tạo / làm mới tab Trang chủ". Tab "TQN Trang chủ" đứng đầu, có 13 nút giống bản thử (Hằng ngày · Tự động · Thời tiết · AI · Giao diện) và dòng tình hình hồ sơ. Nút bấm được trên máy tính (ứng dụng điện thoại không chạy nút hình).

## Tự cấp mã SD · RFI – 19/11
- BPTC chỉ gõ "Nội dung chi tiết" ở dòng mới → ô "Mã hồ sơ" tự điền số tiếp theo, đúng mẫu đang dùng:
  - SD: `yymmdd-COA.NP-NT-SDS-nnn` (gói NP/HTKT lấy theo mã dòng trên)
  - RFI: `yymmdd-COA-TBC-RFI-nnn rev01`, Rev tự điền 01
- Chỉ điền ô trống (Mã, STT, Rev); không sửa mã đã có; khóa chống 2 người cấp trùng số; mỗi mã ghi vào tab phụ TQN_CAP_MA (trạng thái TỰ ĐỘNG).
- Bật/tắt: menu "Tự cấp mã SD · RFI khi gõ Nội dung: BẬT" / "Tự cấp mã: tắt". Thêm loại khác (RFA, RFC, MAT…): sửa khóa `auto_code_types` trong TQN_CONFIG, ví dụ `SDS RFI RFA`.
- Dòng cũ có nội dung mà thiếu mã: menu "Cấp mã cho các dòng còn trống mã" (xem trước danh sách rồi mới ghi).
- Kiểm & sửa mã nhập tay: menu "Kiểm & sửa mã nhập tay (trùng số, STT, khoảng trắng)". Hiện danh sách trước, bấm Có mới sửa. Mã trùng ĐÃ trình thì chỉ báo, không đổi. Mã cũ lưu ở TQN_CAP_MA (cột "Từ số hiệu").
- RFI tự đồng bộ mã: gõ/sửa "Ngày trình duyệt" hoặc "Rev" → phần ngày + rev trong mã tự đổi theo (giữ số thứ tự). Vd sửa Rev 02, ngày 20/11 → 261120-COA-TBC-RFI-002 rev02. Mã cũ lưu ở TQN_CAP_MA (ĐỒNG BỘ NGÀY/REV).
- Cột Mã chuyên nghiệp (19/11 16:01): ghi chú hướng dẫn ở tiêu đề "Mã hồ sơ"; ô mã tự cấp/tự đổi có ghi chú giờ + mã trước; gõ tay vào cột Mã → Google hỏi lại (bảo vệ kiểu cảnh báo), mã sai mẫu/trùng số tô đỏ + ghi chú; khoảng trắng thừa tự bỏ.

## Thả file tự phân phối – 19/11 16:29
- Thư mục Drive: "TQN Register – Hồ sơ › 0. THẢ FILE VÀO ĐÂY" (menu TQN Register → 📂 Thả file tự phân phối → nút Mở thư mục).
- Tên file có mã hồ sơ (vd 261117-COA.NP-NT-SDS-005.pdf). 15 phút/lần TQN chuyển file vào thư mục đúng sổ (SD, RFI…) và gắn link vào ô Mã hồ sơ – bấm mã là mở file. Ngày/rev trong tên khác sổ vẫn khớp theo số thứ tự.
- Không nhận ra mã → "Chưa nhận ra mã". Không xóa file nào. Nhật ký: tab TQN_THA_FILE.

## Tiến độ ↔ hồ sơ chặn thi công – 19/11 (đã viết code, CHƯA nạp lên sheet)
Menu TQN Register → "⛔ Tiến độ ↔ hồ sơ chặn thi công". Tạo tab "TQN Tiến độ – hồ sơ chặn": tự lấy 58 bản vẽ SD (mã, nội dung, hạng mục, khối RE4/RE5/Công viên). Anh điền ô vàng: Ngày thi công dự kiến + Việc ngoài trời. Cột "Chặn thi công?" tự báo ⛔ / ⚠ / ✔. Làm mới không mất dữ liệu đã điền.
