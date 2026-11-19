# ADR-0002: Cache máy chủ theo phiên bản dữ liệu (dv) + hỏi nhẹ apiPeek
- Trạng thái: Chấp nhận trong code (commit 9dfc7b9), **CHƯA ĐO trên DEV thật**.

## Quyết định
- Mỗi lệnh ghi thành công đổi `TQN6_DV` (Script Properties). Khóa CacheService = dự án + dv → không bao giờ trả dữ liệu cũ sau khi có ghi qua API.
- Web hỏi `apiPeek` (chỉ đọc 1 thuộc tính) mỗi 20 s và khi quay lại tab.
- Sửa tay thẳng vào tab TQN6_* không đổi dv → dựa vào TTL 300 s.

## Bằng chứng
Test `gas-adapter.spec.ts` mục 7: lần mở thứ 2 đọc Sheets 0 lần; sau khi người khác ghi, dv đổi và mở lại thấy bản mới.

## Rủi ro mở
- CacheService giới hạn 100 KB/khóa → danh sách lớn sẽ không được cache (đã bỏ qua an toàn, chỉ mất tốc độ). Cần đo khi > 500 hồ sơ.
- Properties có hạn mức ghi/ngày – mỗi lệnh ghi 1 lần đổi dv. Cần ước lượng số thao tác/ngày.
