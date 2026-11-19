# ADR-0001: V6 tách lõi nghiệp vụ / adapter lưu trữ / giao diện
- Trạng thái: Chấp nhận (Claude, 19/11). **Mời ChatGPT phản biện.**
- Issue: (tạo kèm)

## Bối cảnh
Register cũ là 1 file Mã.gs ~627 KB, hàm giao diện đọc/ghi thẳng 32 tab, khó test, khó chuyển PostgreSQL.

## Quyết định
- `v6/src/core`: nghiệp vụ thuần TypeScript (đánh số, revision, workflow, quyền, idempotency, audit), không gọi API Google.
- `Repo` interface: `MemoryRepo` (test), `SheetsRepo` (Apps Script, tab `TQN6_*`), sau này `PostgresRepo`.
- `v6/src/backend/main.ts`: hàm toàn cục cho `google.script.run`, chỉ là adapter.
- Giao diện React một trang (SPA), 1 lần gọi máy chủ cho mỗi thao tác (`apiDo` trả cả danh sách + hồ sơ 360).
- Mọi lệnh ghi: LockService + idem_key + version (CONFLICT) + nhật ký.

## Bằng chứng
35 test PASS (`v6/tests`), gồm 10 kịch bản an toàn ghi (`an-toan.spec.ts`).

## Điểm yếu đã biết (cần phản biện)
1. `Repo` đang **đồng bộ** → Postgres (async) phải đổi interface. Phương án: chuyển lõi sang async ngay khi ít code, hay giữ sync cho Apps Script và viết lớp async riêng?
2. `SheetsRepo` đọc cả tab mỗi request (cache trong request + CacheService theo dv). Khi TQN6_EVENTS lớn (>10k dòng) sẽ chậm → cần tách đọc theo documentId hoặc chỉ đọc N dòng cuối.
3. V6 chưa nối với sổ PROD (chỉ ghi file DEV) – kế hoạch chuyển đổi dữ liệu (mục 33) chưa có ADR.
