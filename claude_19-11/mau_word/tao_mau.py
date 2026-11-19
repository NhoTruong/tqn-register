# Tạo mẫu phiếu có ô giữ chỗ {{...}} từ mẫu gốc RFI/RFC/RFA của dự án
import sys, docx, copy

def dat(p, text):
    """Đặt chữ cho đoạn, giữ định dạng run đầu, xoá các run còn lại."""
    if not p.runs:
        p.add_run(text); return
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)

def bang1(t):
    c = lambda r, k: t.rows[r].cells[k].paragraphs[0]
    dat(c(1, 1), "Công ty: {{BEN_GUI}}")
    dat(c(1, 3), "Công ty: {{GUI_TOI}}")
    dat(c(2, 1), "Ông/ Bà : {{NGUOI_LAP}}")
    dat(c(2, 3), "Ông/ Bà : {{NGUOI_NHAN}}")
    # ô trống: chép định dạng từ ô nhãn cùng dòng
    for (r, k, key) in [(4, 1, "{{NGAY_BAN_HANH}}"), (4, 3, "{{CC}}")]:
        p = c(r, k)
        if p.runs: dat(p, key)
        else:
            lab = c(r, 0).runs[0]
            nr = copy.deepcopy(lab._r); p._p.append(nr)
            p.runs[0].text = key
    dat(c(5, 1), "{{SO_THAM_CHIEU}}")
    dat(c(5, 3), "{{HAN_PHAN_HOI}}")

def ve_viec(d):
    for p in d.paragraphs:
        if p.text.startswith("Về việc"):
            dat(p, "Về việc: {{VE_VIEC}}"); return
    raise SystemExit("không thấy Về việc")

def mo_ta(t, rfc=False):
    cell = t.rows[1].cells[0] if not rfc else t.rows[2].cells[0]
    for p in cell.paragraphs:
        if p.text.startswith("Căn cứ Hợp đồng"):
            dat(p, "Căn cứ Hợp đồng số {{SO_HOP_DONG}}")
        elif p.text.startswith("Nhà thầu") or p.text.startswith("Mô tả các yêu cầu"):
            dat(p, "{{NOI_DUNG}}")

def main(src, dst, loai):
    d = docx.Document(src)
    t = d.tables
    bang1(t[1]); ve_viec(d)
    if loai == "RFC":
        mo_ta(t[2], rfc=True)
        for p in t[2].rows[4].cells[0].paragraphs:
            for k, key in [("Tiến độ", "AH_TIEN_DO"), ("Thiết kế", "AH_THIET_KE"), ("Chi phí", "AH_CHI_PHI")]:
                if p.text.startswith(k):
                    dat(p, f"{k} : {{{{{key}}}}}")
        for p in t[2].rows[5].cells[0].paragraphs:
            if p.text.startswith("1."):
                dat(p, "{{DINH_KEM}}")
    else:
        mo_ta(t[2])
        ps = t[2].rows[2].cells[0].paragraphs
        for i, p in enumerate(ps):
            if p.text.startswith("Đính kèm"):
                dat(p, "Đính kèm (nếu có): {{DINH_KEM}}")
    d.save(dst)

if __name__ == "__main__":
    main(*sys.argv[1:4])
