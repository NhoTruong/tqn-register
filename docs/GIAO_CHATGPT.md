# TQN Register – toàn bộ code cần viết (giao ChatGPT + phần nâng cao)

Cập nhật 19/11 23:15. Repo: NhoTruong/tqn-register (riêng tư). Lõi V6: `v6/src/core` (TypeScript), adapter Apps Script `v6/src/backend`, giao diện React `v6/src/frontend`, test vitest `v6/tests`. Register PROD đang chạy: 1 file `Mã.gs` (bản sao lưu `prod-baseline/Ma.js`).

## LUẬT CHUNG cho ChatGPT (dán kèm mọi việc)
1. Không sửa các file Claude đang giữ: `v6/src/core/service.ts`, `repo.ts`, `types.ts`, `upload.ts`, `v6/src/backend/*.ts`, `v6/src/frontend/App.tsx`, `api.ts`. Chỉ tạo **file mới** theo tên ghi ở từng việc.
2. TypeScript strict, không `any` thừa, không thư viện ngoài nếu không ghi rõ. Chữ hiển thị và chú thích bằng tiếng Việt.
3. Mỗi việc kèm file test vitest `v6/tests/<tên>.spec.ts`, chạy `npm test` phải PASS.
4. Không chứa mật khẩu, API key, email thật. Key chỉ đọc từ Script Properties.
5. Giao diện không ghi thẳng vào Sheets; mọi ghi đi qua API có sẵn (`api.act`, `api.upload`) hoặc hàm adapter mới do Claude nối.
6. AI chỉ được **đề xuất**, không tự phê duyệt hay trả lời hồ sơ.
7. Code Apps Script PROD: viết thành hàm mới tên bắt đầu `tqn`, có chế độ `xemTruoc` (chỉ báo, không đổi) chạy trước. Gửi về Claude soát trước khi nạp. Không push đè PROD.
8. Trả về: file code + file test + 3 dòng mô tả "đã làm / chưa làm / giả định".

---

## NHÓM A – Apps Script PROD (file Mã.gs), ChatGPT viết được ngay

### A1. Dọn trigger trùng (mục 11)
- Hàm `tqnDonTrigger(xemTruoc)`: `ScriptApp.getProjectTriggers()`, nhóm theo `getHandlerFunction()` + `getEventType()` (+ nguồn), giữ 1 cái mỗi nhóm, xóa phần dư. `xemTruoc = true` chỉ trả bảng (hàm, loại, số lượng, sẽ xóa mấy).
- Ghi kết quả vào nhật ký `console.log` + trả object.
- Đạt: chạy 2 lần liên tiếp, lần 2 xóa 0.

### A2. Hộp Workflow AI báo lỗi lặp từng dòng (mục 8)
- Cần anh dán cho ChatGPT các hàm AI trong `Ai.gs` / phần AI của Mã.gs.
- Viết `tqnAiGoiAnToan_(fn)`: lỗi lần đầu do key/quota (HTTP 401/403/429, chữ "API key", "quota") → dừng cả vòng, báo **1 lần** kèm nguyên nhân, không gọi tiếp các dòng sau. Lỗi riêng 1 dòng → ghi vào danh sách, cuối vòng báo gộp "N dòng lỗi".
- Hàm `tqnAiKiemKetNoi()`: kiểm có key trong Script Properties không (chỉ báo có/không, không in key), gọi thử 1 lần, trả mã HTTP.

### A3. Bật Sheets advanced service (mục 12)
- Sửa `appsscript.json`: thêm `"dependencies": { "enabledAdvancedServices": [{ "userSymbol": "Sheets", "serviceId": "sheets", "version": "v4" }] }`. Giữ nguyên các khóa khác. Việc nhỏ, ChatGPT chỉ cần đưa đoạn JSON.

### A4. Bộ đo tốc độ 6 nút PROD (mục 23, 24)
- `tqnDoTocDo(tenNut, soLan)`: chạy hàm của nút ở chế độ không ghi (hoặc trên bản sao), đo `Date.now()`, tính median, p95, ghi tab `TQN_DO_TOC_DO` (giờ, nút, n, median, p95, lỗi).
- `tqnThuBam2Lan(tenNut)`: gọi 2 lần liên tiếp với cùng dữ liệu, báo có sinh trùng mã/email không.
- Danh sách 6 nút anh đưa kèm (tên hàm trong menu).

---

## NHÓM B – V6, file mới độc lập, ChatGPT viết được

### B1. Cấu hình RFC, SDS, MSS, MAS (mục 15) – chỉ dữ liệu
- File mới `v6/src/core/config.types.ts`, xuất hằng `RFC_OPTIONS_COASTAL`, `SDS_OPTIONS_COASTAL`, `MSS_OPTIONS_COASTAL`, `MAS_OPTIONS_COASTAL` cùng dạng:
  `{ code: string; text: string; result: 'CLOSE'|'NEED_INFO'|'FORWARD'|'RESUBMIT'|'REJECT' }[]`
- Chữ trong `text` phải **đúng y chữ trên phiếu** của dự án (anh đưa ảnh phiếu). Kèm `slaDays` từng loại.
- Mẫu đã có (RFI): `1 Xem trả lời bên dưới → CLOSE`, `2 Vui lòng cung cấp thêm thông tin → NEED_INFO`, `3 Chuyển tiếp tới → FORWARD`. RFA: A/B → CLOSE, C → RESUBMIT, D → REJECT.

### B2. Lint + CI GitHub Actions (mục 31)
- `v6/eslint.config.js` (flat config, `typescript-eslint`, `eslint-plugin-react-hooks`), script `"lint"` trong package.json (đưa đoạn cần thêm, không sửa thẳng).
- `.github/workflows/v6.yml`: Node 20, `cd v6 && npm ci && npm run typecheck && npm test && npm run build`, chạy khi push/PR vào `main`. Không dùng secret.
- `v6/scripts/rollback-drill.md`: các bước diễn tập rollback DEV (clasp deploy lại version cũ, kiểm `tqn6Version`).

### B3. Smart Intake (mục 25) – hàm thuần
- File `v6/src/core/intake.ts`: `export function intake(fileName: string, text?: string): { type: DocType | null; confidence: number; titleGuess: string; refDocNo: string | null; revGuess: number | null; reasons: string[] }`
- Đọc mã dạng `\d{6}-[A-Z]{2,5}-[A-Z]{2,5}-(RFI|RFA|RFC|SDS|MSS|MAS)-\d{3}`, rev dạng `rev\s*0?\d+|R\d+`, từ khóa tiếng Việt có/không dấu (yêu cầu thông tin, phê duyệt vật liệu, biện pháp thi công, shop drawing…).
- Test ≥ 15 tên file thật dạng dự án (anh đưa danh sách tên file).

### B4. Màn nhật ký / audit (mục 30)
- `v6/src/frontend/AuditView.tsx`: props `{ events: EventRec[] }` (kiểu trong `core/types.ts`: eventId, at, actor, actorKind, action, documentId, revNo, idemKey, detail, ok). Lọc theo người, hành động, ngày, chỉ lỗi; xuất CSV.

### B5. Màn quản lý quyền (mục 29)
- `v6/src/frontend/MembersEditor.tsx`: props `{ members: Record<string, Role[]>; onSave(m): Promise<void> }`, Role = ADMIN | DC | CONTRACTOR | REVIEWER | VIEWER. Kiểm email hợp lệ, không cho xóa ADMIN cuối cùng, hiện thay đổi trước khi lưu.

### B6. Màn cấu hình dự án (mục 28)
- `v6/src/frontend/ProjectConfigEditor.tsx` + `v6/src/core/configCheck.ts`: kiểm mẫu mã chỉ dùng token `{yymmdd} {project} {contractor} {type} {seq:N}`, slaDays 0–60, mỗi loại ≥ 1 ô CLOSE. Xem trước mã mẫu bằng `formatNo` trong `core/numbering.ts`.

### B7. Chuông thông báo trong web (mục 26)
- `v6/src/frontend/Bell.tsx`: props `{ items: { to: string; docNo: string; title: string; why: string }[]; me: string; roles: string[] }` (đúng dạng `service.reminders()`), lọc việc của tôi, số đỏ, bấm mở hồ sơ qua `onOpen(docNo)`.

### B8. Trang chủ V6 + Field Companion (mục 20, 22)
- `v6/src/frontend/Home.tsx`: ô KPI (chờ CĐT, quá hạn, cần bổ sung, mới 7 ngày), danh sách việc hôm nay, khung thời tiết (props, không gọi API).
- `v6/src/frontend/Field.tsx`: bản điện thoại: tìm mã, xem 360, chụp ảnh tải lên (dùng `api.upload`), nút to, chạy tốt màn 380 px.

### B9. AI đọc phiếu cho V6 (mục 16) – cần anh đồng ý dùng API key
- `v6/src/backend/aiExtract.ts`: `aiExtract(text: string): { title; type; questions: string[]; refDrawings: string[]; confidence }` gọi Claude API qua `UrlFetchApp`, key ở Script Properties `TQN_AI_KEY`, timeout, 1 lần thử lại khi 429/5xx, lỗi trả `{ error, code }` 1 lần. Kết quả chỉ để **điền gợi ý**, người dùng bấm lưu.
- Test bằng giả `UrlFetchApp`.

### B10. Weather Engine + Automation Monitor (mục 27)
- Cần anh dán hàm thời tiết cũ. Viết `v6/src/core/weather.ts` (thuần: đổi dữ liệu API → dòng nhật ký công trường) và `v6/src/backend/monitor.ts`: liệt kê trigger, lần chạy cuối, lỗi cuối (đọc tab log), trả JSON cho màn theo dõi.

---

## NHÓM C – Nâng cao (làm sau khi RFI chạy thật)

| # | Việc | Ai viết | Ghi chú |
|---|---|---|---|
| C1 | **Nhập dữ liệu RFI/RFA hiện có từ sổ PROD vào V6** (mục 33 – mới khai thêm) | ChatGPT viết bộ đọc + Claude nối | Đọc tab RFI cũ → DocumentRec/RevisionRec, giữ mã cũ, chạy thử trên bản sao, báo dòng lỗi |
| C2 | PostgresRepo / ASP.NET Core (mục 32) | ChatGPT | Bảng theo `db/schema_v0.sql`; Repo hiện **đồng bộ** – bản Postgres cần interface async, Claude đổi lõi trước rồi mới viết |
| C3 | Phiếu RFI/RFA xuất PDF đúng mẫu dự án, có QR mã hồ sơ | ChatGPT | HTML in A4 + QR, dữ liệu từ get360 |
| C4 | Thống kê thời gian phản hồi CĐT (SLA) theo tuần, biểu đồ | ChatGPT | Từ EventRec SUBMIT/REPLY |
| C5 | Hàng đợi offline cho Field Companion (IndexedDB) | ChatGPT | Gửi lại bằng cùng idem khi có mạng |
| C6 | Phân quyền theo nhà thầu phụ / hạng mục | Claude (đụng lõi) | |
| C7 | Thông báo Zalo/Telegram (thay email) | ChatGPT | Chỉ khi anh muốn |
| C8 | Tìm kiếm trong nội dung PDF đã tải (Drive OCR) | ChatGPT | Cần quyền Drive đầy đủ |

---

## Claude giữ (không giao, vì đụng lõi / máy anh / PROD)
Mục 1–7 (nạp + chạy thật DEV), 9–10 (tăng tốc Quản lý hồ sơ, Tạo mã – sửa `apps-script/QuanLy.gs`, `Manager.gs`), 13 (E7), 14 + 18 (nối PROD), 17 (STAGING), 19 (người dùng thật), 21 (kiểm dọn job), và **soát + nối** mọi file ChatGPT gửi về.
