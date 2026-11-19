"""Đóng gói bản cài: dist/TQN_Register_1_FILE.gs (code + HTML gộp), dist/TQN_Code.gs (chỉ code) và các file HTML."""
import json, os, shutil
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A, D = os.path.join(ROOT, "apps-script"), os.path.join(ROOT, "dist")
GS = ["Config", "Read", "Import", "QC", "Remind", "Weather", "Workflow", "Ai", "Theme", "Docs", "Manager", "Auto", "Menu"]
HT = ["Sidebar", "Workflow", "NewDoc", "Manager"]
code = "".join(f"\n// ===== {g}.gs =====\n" + open(os.path.join(A, g + ".gs"), encoding="utf-8").read() for g in GS)
emb = {h: open(os.path.join(A, h + ".html"), encoding="utf-8").read() for h in HT}
open(os.path.join(D, "TQN_Code.gs"), "w", encoding="utf-8").write("// TQN Register — code (dùng kèm 4 file HTML: " + ", ".join(HT) + ")\n" + code)
open(os.path.join(D, "TQN_Register_1_FILE.gs"), "w", encoding="utf-8").write(
    "// TQN Register — BẢN 1 FILE: chỉ cần dán toàn bộ file này vào Code.gs (không cần tạo file HTML).\n" + code
    + "\n// ===== Giao diện HTML gộp sẵn =====\nvar TQN_HTML_EMBED = " + json.dumps(emb, ensure_ascii=False) + ";\n")
for h in HT: shutil.copy(os.path.join(A, h + ".html"), os.path.join(D, h + ".html"))
print("1 file:", os.path.getsize(os.path.join(D, "TQN_Register_1_FILE.gs")) // 1024, "KB;", len(GS), "file .gs +", len(HT), "HTML")
