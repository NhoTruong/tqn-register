# Tạo MAU_TRN.docx (phiếu giao nhận gộp nhiều hồ sơ) từ khung mẫu RFA của dự án
import sys, copy, docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

src, dst = sys.argv[1], sys.argv[2]
d = docx.Document(src)          # src = MAU_RFA.docx (đã có ô {{...}})
t = d.tables
body = d.element.body

def dat(p, text):
    if not p.runs: p.add_run(text); return
    p.runs[0].text = text
    for r in p.runs[1:]: r._r.getparent().remove(r._r)

# 1. Bảng thông tin: đổi tiêu đề + ô "Hy vọng phản hồi vào"
for c in t[1].rows[0].cells: dat(c.paragraphs[0], "PHIẾU GIAO NHẬN HỒ SƠ")
dat(t[1].rows[5].cells[2].paragraphs[0], "Số lượng hồ sơ")
dat(t[1].rows[5].cells[3].paragraphs[0], "{{TRN_SL}}")

# 2. Bỏ bảng mô tả + 2 bảng phản hồi + 2 tiêu đề BÊN BAN HÀNH / PHẢN HỒI
for tb in t[2:]: body.remove(tb._tbl)
for p in list(d.paragraphs):
    if p.text.strip().startswith(("BÊN BAN HÀNH", "PHẢN HỒI TỪ BÊN NHẬN")):
        body.remove(p._p)

# chèn sau đoạn "Về việc"
anchor = next(p for p in d.paragraphs if p.text.startswith("Về việc"))
font_name = "Arial"

def run(p, text, bold=False, size=10):
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size); r.font.name = font_name
    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font_name)
    return r

def borders(tbl):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}"); x.set(qn("w:val"), "single"); x.set(qn("w:sz"), "4"); x.set(qn("w:color"), "000000")
        b.append(x)
    tblPr.append(b)

def fix(tbl, widths):
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(int(sum(widths) * 567))); w.set(qn("w:type"), "dxa"); tblPr.append(w)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"), "-639"); ind.set(qn("w:type"), "dxa"); tblPr.append(ind)
    for gc, wd in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths): gc.set(qn("w:w"), str(int(wd * 567)))
    for row in tbl.rows:
        for c, wd in zip(row.cells, widths): c.width = Cm(wd)
        for c in row.cells:
            for p in c.paragraphs:
                if p.alignment is None: p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:fill"), hexc); tcPr.append(s)

def new_par_after(el, text="", bold=False, size=11):
    p = d.add_paragraph(); run(p, text, bold, size) if text else None
    el.addnext(p._p); return p

h1 = new_par_after(anchor._p, "DANH MỤC HỒ SƠ GIAO", True, 11); h1.paragraph_format.left_indent = Cm(-1.25)
h1.paragraph_format.space_before = Pt(10)

cols = [("STT", 1.0), ("Số tham chiếu", 5.0), ("Nội dung hồ sơ", 8.12), ("Rev", 1.0), ("Bản cứng", 1.4), ("Ghi chú", 2.1)]
tb = d.add_table(rows=2, cols=len(cols)); borders(tb)
keys = ["{{TRN_STT}}", "{{TRN_MA}}", "{{TRN_ND}}", "{{TRN_REV}}", "{{TRN_BC}}", "{{TRN_GC}}"]
for i, (name, w) in enumerate(cols):
    for r in (0, 1):
        c = tb.rows[r].cells[i]; c.width = Cm(w)
    c0 = tb.rows[0].cells[i]; shade(c0, "FCE4D6")
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; run(p, name, True, 9)
    p = tb.rows[1].cells[i].paragraphs[0]; run(p, keys[i], False, 9)
    if i in (0, 3, 4): p.alignment = WD_ALIGN_PARAGRAPH.CENTER
fix(tb, [w for _, w in cols]); h1._p.addnext(tb._tbl)

note = new_par_after(tb._tbl, "Đính kèm: {{DINH_KEM}}", False, 10); note.paragraph_format.left_indent = Cm(-1.25)
note.paragraph_format.space_before = Pt(6)

sg = d.add_table(rows=3, cols=2); borders(sg)
for i, (a, b) in enumerate([("BÊN GIAO", "BÊN NHẬN"), ("Ký, ghi rõ họ tên", "Ký, ghi rõ họ tên"), ("Ngày: …/…/20…   Giờ: …:…", "Ngày: …/…/20…   Giờ: …:…")]):
    for j, txt in enumerate((a, b)):
        c = sg.rows[i].cells[j]; c.width = Cm(9.6)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, txt, i == 0, 10 if i == 0 else 9)
        if i == 0: shade(c, "D9D9D9")
sg.rows[1].height = Cm(2.6); fix(sg, [9.31, 9.31])
gap = new_par_after(note._p, "", False, 6)
gap._p.addnext(sg._tbl)
d.save(dst)
print("ok")
