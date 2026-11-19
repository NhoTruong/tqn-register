# Phối hợp Claude + ChatGPT trên TQN Register (1 sản phẩm, 1 repo)

## Mã nguồn ở đâu
| Thư mục | Nội dung |
|---|---|
| `prod-baseline/Ma.js` | Bản sao lưu code Register PROD đang chạy (clasp clone 19/11 22:18, SHA trong `SHA256SUMS`). Chỉ đọc. |
| `apps-script/*.gs` | Mã nguồn Register bản cũ, tách theo file (QuanLy, Manager, TuDongMa, Workflow, Weather…). Ghép thành 1 file Mã.gs khi nạp. |
| `v6/` | Register V6: lõi TypeScript `src/core`, adapter Apps Script `src/backend`, web React `src/frontend`, test `tests`. Build bằng `npm run build` → `gas/dist`. |
| `docs/VIEC_CON_LAI.md` | Toàn bộ việc còn lại (đánh số). |
| `docs/GIAO_CHATGPT.md` | Đặc tả từng việc giao ChatGPT + luật chung. |

## Quy trình
1. **Không ai push thẳng `main`** trừ Claude khi gộp. Mỗi bên làm trên nhánh riêng:
   - ChatGPT: `gpt/<mục>-<tên-ngắn>` (vd `gpt/11-don-trigger`)
   - Claude: `claude/<mục>-<tên-ngắn>` hoặc commit nhỏ lên `main` sau khi test PASS
2. Mở **Pull Request vào `main`**, ghi: mục số mấy trong VIEC_CON_LAI, đã làm / chưa làm / giả định, kết quả `npm test`.
3. Bên kia review trên PR. Claude chạy test + nạp **DEV** để kiểm thật, báo tốc độ, rồi mới gộp.
4. **PROD chỉ nạp khi anh Nho duyệt**, sau khi đã sao lưu (`prod-baseline` cập nhật) và test.
5. Trước khi bắt đầu một mục: ghi tên mình vào cột "Ai làm" trong `docs/VIEC_CON_LAI.md` (qua PR nhỏ) để tránh 2 bên làm trùng.

## Chia file để không đụng nhau
- **Claude giữ:** `v6/src/core/service.ts`, `repo.ts`, `types.ts`, `upload.ts`, `v6/src/backend/*`, `v6/src/frontend/App.tsx`, `api.ts`, `apps-script/QuanLy.gs`, `Manager.gs` (đang tăng tốc nút Quản lý hồ sơ – mục 9).
- **ChatGPT nhận trước:** mục 8 (AI lỗi lặp – `apps-script/Ai.gs`, `Workflow.gs`), 11 (dọn trigger – file mới `apps-script/DonTrigger.gs`), 12 (manifest), 15 (cấu hình phiếu – `v6/src/core/config.types.ts`), 31 (CI/lint – `.github/workflows/`, `v6/eslint.config.js`), các file mới nhóm B trong GIAO_CHATGPT.md.
- Cần sửa file của bên kia → nêu trong PR, không sửa thẳng.

## Kiểm trước khi mở PR
```
cd v6 && npm ci && npm run typecheck && npm test && npm run build
```
Code Apps Script cũ (`apps-script/`): test ở `apps-script/tests` / `test/` (node), chạy `node test/run_quanly_test.js`.

## Luật an toàn (cả hai bên)
- Không đưa mật khẩu, API key, token, email người khác vào repo. Key chỉ ở Script Properties.
- AI không tự phê duyệt/trả lời hồ sơ thay người có thẩm quyền.
- Không sửa dữ liệu gốc Register, không xóa file Drive, không triển khai Web App công khai.
