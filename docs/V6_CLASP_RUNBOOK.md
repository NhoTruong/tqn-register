# Quy trình clasp – TQN Register (19/11/2026)

## Hiện trạng đã kiểm (21:30–21:45)
- PROD: Script ID `143dMX2iKTeSjODi2n2xeAOunMNLRcxiDFsjH7Xeqm8pjozKfyK40wOvu` (gắn sổ HAUS QUANG NGAI- Register). 2 file: `appsscript.json`, `Mã.gs` (627.517 ký tự, SHA-256 bắt đầu 6cd5ba87).
- Sao lưu phía Google: **phiên bản 1** lúc 21:38 (triển khai loại Thư viện, chỉ để giữ bản, không mở cho ai).
- Không có Web App deployment nào trước đó.
- 10 trình kích hoạt; **tqnWeatherJob bị cài 4 lần** (3 cái chưa chạy lần nào) → cần dọn còn 1.
- `appsscript.json` đang chạy KHÔNG bật dịch vụ Sheets nâng cao → nút "Bộ lọc sẵn" sẽ lỗi cho tới khi bật.
- Cài đặt người dùng "API Google Apps Script": **Tắt** → clasp chưa dùng được.
- Máy anh (Windows): **chưa có Node.js** (lệnh `npx` không nhận) → chưa cài được clasp.

## Anh làm 1 lần (không đưa mật khẩu/token cho Claude)
1. Mở https://script.google.com/home/usersettings → bật **API Google Apps Script**.
2. Cài Node.js LTS từ https://nodejs.org (bản Windows, Next đến hết).
3. Mở PowerShell: `npm install -g @google/clasp` rồi `clasp login` → trình duyệt mở, anh chọn tài khoản nhotruong562@gmail.com, bấm Cho phép. Thông tin đăng nhập nằm trong máy anh (`~/.clasprc.json`), không gửi đi đâu.
4. Tạo repo **riêng tư** `NhoTruong/tqn-register` trên GitHub (để trống).
5. Tạo 2 bản sao sổ Register làm DEV và STAGING (Tệp → Tạo bản sao); gửi Claude 2 đường link.

## Quy trình sau khi có clasp
- Trước mọi lần push: `node scripts/pull-backup.mjs prod` → kéo nguyên trạng về `backup/…` kèm SHA-256.
- Thử: `node scripts/deploy.mjs dev` → `staging` → `prod`.
  - PROD chỉ chạy khi: git sạch, nhánh `main`, có tag `v*`, test PASS, và đặt `TQN_CONFIRM_PROD=<commit>`.
- Mỗi lần deploy ghi `RELEASES.md`: thời điểm, môi trường, commit, version Apps Script, deployment.
- Kiểm bản đang chạy: chạy hàm `tqn6Version` hoặc xem chân trang Web App → ra commit; đối chiếu `RELEASES.md`.
- Rollback: `node scripts/rollback.mjs prod <version cũ>` → trỏ deployment về version cũ, không sửa code.
