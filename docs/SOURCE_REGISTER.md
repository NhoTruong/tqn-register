# Sổ nguồn (V3)
SHA256 đầy đủ có trong `evidence/v3/sha256_goi_V3.txt` và `evidence/v3/sha256_nguon_nghiep_vu.txt`. Bảng `SHA256_MANIFEST.csv` của gói V3 khớp **17/17** file.

| # | Nguồn | SHA256 (16 ký tự đầu) | Xuất xứ | Độ tin cậy | Đã mở bytes |
|---|---|---|---|---|---|
| 1 | Code đang làm: `tqn-register` (Apps Script S1 + thời tiết, 965 dòng) | commit d8066d9 | Claude viết trong phiên này | Nguồn hiện hành | Có. Test 19/19 và 15/15 PASS (chạy giả lập Node) |
| 2 | CLAUDE_MASTER_DIRECTIVE_V3 .md/.docx/.pdf | ae0b229b / 28a4d92f / 082a64c5 | Anh gửi (ZIP V3 và chat 11:04) | Chỉ thị | Có. 3 file anh gửi qua chat trùng byte với bản trong ZIP |
| 3 | Lệnh phát triển V3 (attachment.txt) | — | Anh gửi chat 11:03 | Chỉ thị | Có |
| 4 | Bổ sung V3 Weather API (12 mục) | — | Anh gửi chat 11:10 | Chỉ thị | Có |
| 5 | CLAUDE_BOOT_PROMPT, PROJECT_CONFIG_EXAMPLE.json | ba54b2b5 / 9eec256f | ZIP V3 | Chỉ thị / mẫu | Có |
| 6 | V2 ALL_IN_ONE.zip | 240ae1d4 | ZIP V3, trùng bản đã đọc ở phiên trước | Baseline tham khảo | Có (đã đọc ở phiên trước) |
| 7 | MAX1_FULL.zip | 6bcb99cf | ZIP V3, nội dung trùng MAX1 trong V2 | Baseline tham khảo | Có |
| 8 | 5 file `.bas` MAX1 (730 dòng) | f7fd0078, 3d5eefe8, a0e58349, 53c5e7df, 243de6df | V2 → 00_BAN_MAX1_GOC | Baseline có lỗi (xem CURRENT_STATE) | Có, đọc toàn bộ |
| 9 | HAUS QUANG NGAI- Register.xlsx (32 tab, 6 tab ẩn) | 47e537fe | V2 | Nguồn nghiệp vụ, **không sửa** | Có |
| 10 | QUY ĐỊNH GHI MÃ SỐ HỒ SƠ Rev00 (PDF 14 trang) | 1b38470c | V2 | Ưu tiên cao nhất khi có mâu thuẫn | Có |
| 11 | Mẫu RFI.doc, RFC.doc, RFA.docx | 049af47e / f1a9d0c7 / eadc16a8 | V2 và chat (trùng byte) | Mẫu gốc | Có |
| 12 | 6 ảnh thời tiết 01–06 (945×2048) | 079bbec0 … 5c23afc0 | ZIP V3 | Chỉ tham chiếu trình bày, **không lấy số** | Có, đã xem cả 6 ảnh |
| 13 | ADDITIONAL_UAT_V3.csv (V3-001…020), GATE_STATUS_TEMPLATE.csv | cc13cfec / 7882afbf | ZIP V3 | Bảng nghiệm thu | Có |
| 14 | Register Google Sheets đang dùng | — | Link anh gửi | Bản sống | Chỉ đọc qua gviz. Không ghi |

## Nhận xét về 6 ảnh (tham chiếu UX)
- **01–03, nền xanh trời:** đều là cùng một màn hình, chụp sáng, chiều và tối.
  - Thẻ trên cùng gồm tên địa điểm, nhiệt độ lớn, mô tả, cao/thấp, cảm giác như, giờ cập nhật.
  - Dòng tóm tắt bằng chữ.
  - Dải 6 giờ, mỗi giờ có icon, nhiệt độ, đường nhiệt màu vàng và % mưa có hình giọt nước.
  - Thẻ "Mưa sắp tới" ghi %, kèm dòng "Hôm qua".
- **04–06, nền tối:** gồm ngày, cao/thấp, nhãn "Có mưa", dải giờ có % mưa, bộ chọn 7 ngày trong tuần và 3 dòng Lượng mưa / Gió / Độ ẩm.
- **Lấy về dùng:** bố cục chia sáng/chiều/tối, đường nhiệt độ, % mưa theo giờ, bộ chọn ngày, ba chỉ số.
- **Không lấy:** ảnh nền có người, linh vật ếch, logo hay chữ của Samsung, Google, AccuWeather, và mọi con số trên ảnh.
