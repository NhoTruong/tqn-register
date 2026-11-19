# TQN Register – toàn bộ việc còn lại (cập nhật 19/11 23:11)

Quy tắc: báo việc còn lại là liệt kê HẾT, tự khai đủ, đối chiếu mọi yêu cầu đã giao. Mỗi lần xong/thêm việc thì sửa file này.

## 1. Đã viết code + test, CHƯA CHẠY THẬT trên DEV (chờ nạp cap_nhat_dev.bat, commit 9dfc7b9)
1. Tải file thật lên Drive, mở/tải về được
2. Tự điền dòng vào sổ RFI khi Tạo/Trình/Trả lời
3. Tự đặt tên file mã_revNN, trùng tên khác nội dung thêm đuôi hash
4. ĐANG ĐỒNG BỘ / Thử lại / báo xung đột CONFLICT
5. Cache theo dv + tự phát hiện người khác vừa sửa
6. Nút Xuất báo cáo (In/PDF, CSV)
7. Đo median/p95 thật bằng nút Đo tốc độ

## 2. Lỗi / thiếu trên Register PROD
8. AI đọc phiếu tự điền: lỗi lặp từng dòng, chưa tìm nguyên nhân gốc
9. Nút Quản lý hồ sơ ~18 giây
10. Tạo mã còn chậm; 4 nút còn lại chưa đo
11. Trigger thời tiết nhân 4
12. Manifest PROD chưa bật Sheets advanced service
13. Ô E7: chờ anh xác nhận trả lại "Không duyệt"
14. V6 chưa nối với sổ PROD (đang ghi file DEV)

## 3. Chưa làm
15. V6 cho RFA, RFC, SDS, MSS, MAS (thiếu bảng ô trả lời + số ngày hạn)
16. AI đọc nội dung phiếu trong V6 (hiện chỉ đoán theo tên file)
17. Môi trường STAGING
18. Đưa V6 lên PROD (sao lưu, test, tag)
19. Danh sách người dùng thật + vai trò (hiện chỉ anh ADMIN)
20. Giao diện Operations Desk / Technical Register / Field Companion (mới có màn RFI + báo cáo)
21. Dọn job upload cũ – có code, chưa kiểm thật
22. Thời tiết, nhắc hạn, Trang chủ bản cũ chưa chuyển sang V6

## 4. Theo yêu cầu gốc V6/V7/kế hoạch 4 giai đoạn, chưa làm
23. Đo 6 nút PROD median/p95 (V7)
24. Kiểm các nút PROD: bấm 2 lần, 2 người cùng lúc, hạn mức email/trigger (giai đoạn 2)
25. Smart Intake: thả file → tự nhận loại, đề xuất mã/nội dung
26. Notification Engine: thông báo trong web / Trang chủ
27. Weather Engine + Automation Monitor trong V6
28. Màn cấu hình nhiều dự án (mẫu mã, ô trả lời, hạn)
29. Màn quản lý quyền người dùng
30. Màn xem nhật ký/audit cho ADMIN
31. Lint, CI GitHub Actions, diễn tập rollback thật
32. Adapter PostgreSQL / ASP.NET Core

## Đã bỏ theo ý anh
- Email nhắc 7h (19/11 23:03)
