Attribute VB_Name = "TQN_TrinhDuyet"
'==============================================================================
' TQN - QUAN LY HO SO TRINH DUYET - Du an Coastal Quang Ngai (TBCONS)
' Ban 2.0 - 19/11/2026
' Register tren Google Sheets giu nguyen: file nay CHI DOC Register, khong ghi nguoc.
'
' Chuoi U("...") chua tieng Viet co dau; khi dong goi duoc doi sang ma \uXXXX (file .bas thuan ASCII).
' Nut bam (tao bang TD_CaiDat):
'   TD_NhapRegister  TD_KiemMa  TD_MaMoi  TD_LenRev  TD_TaoPhieu  TD_GiaoNhan
'   TD_InPhieu  TD_InDanhSach  TD_InNhanGay  TD_LamMoi  TD_MoThuMuc  TD_BoChon
'   TD_LocTrinhDuyet  TD_LocHoiDoi  TD_LocGiaoNhan  TD_LocHienTruong  TD_BoLoc
'==============================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const TEN_BANG As String = "tblPhieu"
Private Const TEN_BANG_VB As String = "tblVanBan"
Private Const TEN_BANG_GN As String = "tblGiaoNhan"
Private Const SH_DS As String = "DANH_SACH"
Private Const SH_VB As String = "VAN_BAN"
Private Const SH_GN As String = "GIAO_NHAN"
Private Const SH_BDK As String = "BANG_DIEU_KHIEN"
Private Const SH_LOG As String = "NHAT_KY"
Private Const SH_NHAN As String = "NHAN_GAY"

' Cot bang tblPhieu (cot hay xem de dau)
Private Const C_CHON As Long = 1
Private Const C_NHOMMAU As Long = 2
Private Const C_LOAI As Long = 3
Private Const C_MA As Long = 4
Private Const C_REV As Long = 5
Private Const C_ND As Long = 6
Private Const C_TT As Long = 7
Private Const C_BUOC As Long = 8
Private Const C_NGAY As Long = 9
Private Const C_HAN As Long = 10
Private Const C_TVGS As Long = 11
Private Const C_CDT As Long = 12
Private Const C_TRN As Long = 13
Private Const C_KIEM As Long = 14
Private Const C_TAO As Long = 15
Private Const C_FILE As Long = 16
Private Const C_IN As Long = 17
Private Const C_NHAN As Long = 18
Private Const C_CC As Long = 19
Private Const C_DK As Long = 20
Private Const C_NGUOI As Long = 21
Private Const C_NGUON As Long = 22
Private Const C_NHOM As Long = 23
Private Const C_MAU As Long = 24
Private Const SO_COT As Long = 24

Private Const wdFormatXMLDocument As Long = 16
Private Const wdExportFormatPDF As Long = 17
Private Const wdFindStop As Long = 0
Private Const wdCollapseEnd As Long = 0
Private Const LOI_TQN As Long = vbObjectError + 700

'------------------------------------------------------------------------------
' CHU TIENG VIET + HOP THOAI
'------------------------------------------------------------------------------
Public Function U(ByVal s As String) As String
    ' doi \uXXXX thanh ky tu Unicode
    Dim i As Long, r As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            r = r & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = r
End Function

Private Function Hop(ByVal noiDung As String, Optional ByVal kieu As Long = vbInformation) As VbMsgBoxResult
    ' MsgBox thuong khong hien duoc tieng Viet co dau -> dung MessageBoxW
    Dim tieuDe As String: tieuDe = "TQN"
    Hop = MessageBoxW(Application.hWnd, StrPtr(noiDung), StrPtr(tieuDe), kieu)
End Function

'------------------------------------------------------------------------------
' TIEN ICH CHUNG
'------------------------------------------------------------------------------
Private Function Bang() As ListObject
    Set Bang = ThisWorkbook.Worksheets(SH_DS).ListObjects(TEN_BANG)
End Function

Private Function CH(ByVal ten As String) As String
    On Error Resume Next
    CH = Trim$(CStr(ThisWorkbook.Names(ten).RefersToRange.Value))
    On Error GoTo 0
End Function

Private Function CoBat(ByVal ten As String) As Boolean
    Dim s As String: s = LCase$(Left$(CH(ten), 1))
    CoBat = (s = "c" Or s = "y" Or s = "1" Or s = "x")
End Function

Private Function Chu(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Chu = "" Else Chu = CStr(v)
End Function

Private Function ThuMuc(ByVal ten As String, ByVal macDinh As String) As String
    Dim s As String: s = CH(ten)
    If Len(s) = 0 Then s = ThisWorkbook.Path & "\" & macDinh
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    ThuMuc = s
End Function

Private Sub TaoThuMuc(ByVal p As String)
    Dim parts() As String, i As Long, cur As String
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

Public Function LamSach(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(s, " -", "-")
    s = Replace(s, "- ", "-")
    LamSach = Trim$(s)
End Function

Private Function TenFile(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "_")
    Next c
    TenFile = Trim$(s)
End Function

Private Function DocNgay(ByVal v As Variant) As Variant
    Dim s As String, p() As String
    DocNgay = Empty
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbDate Then DocNgay = CDate(v): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then DocNgay = CDate(CDbl(v))
        Exit Function
    End If
    s = Trim$(CStr(v))
    If s Like "#*/#*/####*" Then
        p = Split(Left$(s, InStr(s & " ", " ") - 1), "/")
        On Error Resume Next
        DocNgay = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
        On Error GoTo 0
    End If
End Function

Private Function NgayVN(ByVal v As Variant) As String
    Dim d As Variant: d = DocNgay(v)
    If IsEmpty(d) Then NgayVN = "" Else NgayVN = Format$(d, "dd/mm/yyyy")
End Function

Public Function LoaiTuMa(ByVal ma As String, ByVal nguon As String) As String
    Dim m As String, k As Variant
    m = "-" & UCase$(Replace(LamSach(ma), " ", "-")) & "-"
    For Each k In Array("RFI", "RFC")
        If InStr(m, "-" & k & "-") > 0 Then LoaiTuMa = k: Exit Function
    Next k
    For Each k In Array("SDS", "MSS", "MAS", "OTS", "COS", "SCH", "TML", "REP", "DAR", "WER", "MOR", "HSE", "RFP", "RFO")
        If InStr(m, "-" & k & "-") > 0 Then LoaiTuMa = k: Exit Function
    Next k
    If InStr(m, "-RFA-") > 0 Then LoaiTuMa = "RFA": Exit Function
    Select Case UCase$(Left$(nguon, 3))
        Case "SD": LoaiTuMa = "SDS"
        Case "M0S", "MOS": LoaiTuMa = "MSS"
        Case "MAT": LoaiTuMa = "MAS"
        Case "OTS": LoaiTuMa = "OTS"
        Case "SCH": LoaiTuMa = "SCH"
        Case "RFI": LoaiTuMa = "RFI"
        Case "RFC": LoaiTuMa = "RFC"
        Case "TRN": LoaiTuMa = "TML"
        Case "CRP": LoaiTuMa = "REP"
        Case Else
            If LCase$(Left$(nguon, 2)) = "h" & ChrW(&H1ED3) Then LoaiTuMa = "HSE" Else LoaiTuMa = "RFA"
    End Select
End Function

Private Function TachMa(ByVal ma As String, ByRef ngay As String, ByRef nhom As String, _
                        ByRef so As Long, ByRef duoi As String) As Boolean
    Dim re As Object, mc As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(\d{6})-(.+?)-(\d{3})(.*)$"
    ma = LamSach(ma)
    If Not re.Test(ma) Then Exit Function
    Set mc = re.Execute(ma)(0).SubMatches
    ngay = mc(0): nhom = mc(1): so = CLng(mc(2)): duoi = mc(3)
    TachMa = True
End Function

Private Sub GhiLog(ByVal viec As String, ByVal ma As String, ByVal kq As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = viec
    ws.Cells(r, 3).Value = ma
    ws.Cells(r, 4).Value = kq
    ws.Cells(r, 5).Value = Environ$("USERNAME")
End Sub

Private Sub BaoLoi(ByVal viec As String)
    GhiLog viec, "", "LOI " & Err.Number & ": " & Err.Description
    Hop U("Lỗi ") & Err.Number & ": " & Err.Description & vbLf & "(" & viec & ")", vbExclamation
End Sub

Private Sub DatCongThucDong(ByVal lr As ListRow)
    ' cong thuc mau luu o DANH_MUC (ten CT_...), de khong phai viet tieng Viet trong code
    Dim c As Variant, ten As Variant, s As String, i As Long
    c = Array(C_NHOMMAU, C_BUOC, C_HAN, C_MAU)
    ten = Array("CT_NHOM", "CT_BUOC", "CT_HAN", "CT_MAU")
    For i = 0 To 3
        s = CH(CStr(ten(i)))
        If Len(s) > 0 Then lr.Range.Cells(1, c(i)).FormulaR1C1 = "=" & s
    Next i
End Sub

Private Function TimCot(ByVal v As Variant, ByVal soDongDau As Long, ParamArray tuKhoa() As Variant) As Long
    ' tim cot theo chu tieu de (khong phan biet hoa thuong) trong soDongDau dong dau
    Dim r As Long, c As Long, k As Variant, s As String
    For Each k In tuKhoa
        For r = 1 To Application.Min(soDongDau, UBound(v, 1))
            For c = 1 To UBound(v, 2)
                s = LCase$(LamSach(Chu(v(r, c))))
                If Len(s) > 0 Then
                    If InStr(s, LCase$(CStr(k))) > 0 Then TimCot = c: Exit Function
                End If
            Next c
        Next r
    Next k
End Function

'------------------------------------------------------------------------------
' 1. NHAP TU FILE REGISTER (chi doc)
'------------------------------------------------------------------------------
Public Sub TD_NhapRegister()
    On Error GoTo LOI
    Dim duongDan As String, wbR As Workbook, v As Variant, daMo As Boolean
    Dim nThem As Long, nCapNhat As Long, nVB As Long, t As cTatManHinh

    duongDan = CH("FILE_REGISTER")
    If Len(duongDan) = 0 Or Len(Dir$(duongDan)) = 0 Then
        v = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Chon file Register")
        If VarType(v) = vbBoolean Then Exit Sub
        duongDan = CStr(v)
        ThisWorkbook.Names("FILE_REGISTER").RefersToRange.Value = duongDan
    End If

    Set t = New cTatManHinh
    For Each wbR In Application.Workbooks
        If LCase$(wbR.FullName) = LCase$(duongDan) Then daMo = True: Exit For
    Next wbR
    If Not daMo Then Set wbR = Workbooks.Open(duongDan, UpdateLinks:=0, ReadOnly:=True)

    NhapHoSo wbR, nThem, nCapNhat
    nVB = NhapVanBan(wbR)

    If Not daMo Then wbR.Close SaveChanges:=False
    Set wbR = Nothing
    Set t = Nothing
    KiemMaNoiBo False
    LamMoiNoiBo
    ThisWorkbook.Names("LAN_NHAP_CUOI").RefersToRange.Value = Now
    GhiLog "NhapRegister", duongDan, "Ho so: them " & nThem & ", cap nhat " & nCapNhat & "; van ban: " & nVB
    Hop U("Đã nhập từ Register (chỉ đọc, không sửa Register):") & vbLf & _
        U("  Hồ sơ thêm mới: ") & nThem & vbLf & U("  Hồ sơ cập nhật: ") & nCapNhat & vbLf & _
        U("  Văn bản / chỉ thị / NCR: ") & nVB & vbLf & U("Cột Kiểm mã và bảng Việc hôm nay đã làm mới.")
    Exit Sub
LOI:
    If Not wbR Is Nothing And Not daMo Then wbR.Close SaveChanges:=False
    Set t = Nothing
    BaoLoi "NhapRegister"
End Sub

Private Sub NhapHoSo(ByVal wbR As Workbook, ByRef nThem As Long, ByRef nCapNhat As Long)
    ' Sheet kieu trinh duyet: B ma, C noi dung, D nguoi, E tinh trang, F rev, H ke hoach / ngay trinh, I thuc te
    Dim ds() As String, i As Long, r As Long, ws As Worksheet, v As Variant, cuoi As Long
    Dim lo As ListObject, dict As Object, khoa As String, lr As ListRow
    Dim ma As String, nd As String, ngay As Variant, cTV As Long, cCD As Long
    Dim ng As String, nh As String, so As Long, du As String

    Set lo = Bang()
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            khoa = UCase$(LamSach(Chu(lr.Range.Cells(1, C_MA).Value))) & "|" & LCase$(LamSach(Chu(lr.Range.Cells(1, C_ND).Value)))
            If Not dict.Exists(khoa) Then dict.Add khoa, lr.Index
        Next lr
    End If

    ds = Split(CH("SHEET_NGUON"), ",")
    For i = 0 To UBound(ds)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbR.Worksheets(Trim$(ds(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.StatusBar = "TQN: " & ws.Name
            cuoi = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            If cuoi < 2 Then cuoi = 2
            v = ws.Range(ws.Cells(1, 1), ws.Cells(cuoi, 30)).Value
            ' cot TVGS / CDT tim theo tieu de
            cTV = TimCot(v, 8, U("tư vấn giám sát"), "tvgs")
            cCD = TimCot(v, 8, U("chủ đầu tư"), U("cđt"))
            For r = 1 To UBound(v, 1)
                ma = LamSach(Chu(v(r, 2)))
                If ma Like "######-*" Then
                    nd = LamSach(Chu(v(r, 3)))
                    khoa = UCase$(ma) & "|" & LCase$(nd)
                    If dict.Exists(khoa) Then
                        Set lr = lo.ListRows(dict(khoa))
                        nCapNhat = nCapNhat + 1
                    Else
                        Set lr = lo.ListRows.Add
                        DatCongThucDong lr
                        dict.Add khoa, lr.Index
                        nThem = nThem + 1
                    End If
                    With lr.Range
                        .Cells(1, C_NGUON).Value = ws.Name
                        .Cells(1, C_LOAI).Value = LoaiTuMa(ma, ws.Name)
                        .Cells(1, C_MA).Value = ma
                        .Cells(1, C_ND).Value = nd
                        If Len(Trim$(Chu(v(r, 6)))) > 0 Then .Cells(1, C_REV).Value = "'" & Format$(Val(Chu(v(r, 6))), "00")
                        If Len(Trim$(Chu(v(r, 4)))) > 0 Then .Cells(1, C_NGUOI).Value = Trim$(Chu(v(r, 4)))
                        If Len(Trim$(Chu(v(r, 5)))) > 0 Then .Cells(1, C_TT).Value = Trim$(Chu(v(r, 5)))
                        ngay = DocNgay(v(r, 9))
                        If IsEmpty(ngay) Then ngay = DocNgay(v(r, 8))
                        If Not IsEmpty(ngay) And IsEmpty(.Cells(1, C_NGAY).Value) Then .Cells(1, C_NGAY).Value = ngay
                        If cTV > 0 Then .Cells(1, C_TVGS).Value = GopO(v, r, cTV)
                        If cCD > 0 Then .Cells(1, C_CDT).Value = GopO(v, r, cCD)
                        If TachMa(ma, ng, nh, so, du) Then .Cells(1, C_NHOM).Value = nh
                    End With
                End If
            Next r
        End If
    Next i
    Application.StatusBar = False
End Sub

Private Function GopO(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As String
    ' gop ngay + ghi chu cua khoi TVGS/CDT (o dau va o ke ben) thanh 1 chuoi ngan
    Dim a As String, b As String, d As Variant
    d = DocNgay(v(r, c))
    If Not IsEmpty(d) Then a = Format$(d, "dd/mm/yyyy") Else a = LamSach(Chu(v(r, c)))
    If c + 1 <= UBound(v, 2) Then
        d = DocNgay(v(r, c + 1))
        If Not IsEmpty(d) Then
            If Len(a) = 0 Then a = Format$(d, "dd/mm/yyyy")
        Else
            b = LamSach(Chu(v(r, c + 1)))
            If IsNumeric(b) Then b = ""
        End If
    End If
    If Len(a) > 0 And Len(b) > 0 Then GopO = a & " - " & b Else GopO = a & b
End Function

Private Function NhapVanBan(ByVal wbR As Workbook) As Long
    ' So van ban (CV di/den, chi thi, NCR): tu do cot theo tieu de
    Dim ds() As String, i As Long, r As Long, ws As Worksheet, v As Variant, cuoi As Long
    Dim lo As ListObject, lr As ListRow, dict As Object, khoa As String
    Dim cSo As Long, cND As Long, cNg As Long, cNguoi As Long, cTT As Long, cPH As Long, hdr As Long
    Dim so As String, nd As String, n As Long

    Set lo = ThisWorkbook.Worksheets(SH_VB).ListObjects(TEN_BANG_VB)
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            khoa = LCase$(Chu(lr.Range.Cells(1, 2).Value)) & "|" & UCase$(LamSach(Chu(lr.Range.Cells(1, 3).Value))) & "|" & LCase$(LamSach(Chu(lr.Range.Cells(1, 4).Value)))
            If Not dict.Exists(khoa) Then dict.Add khoa, lr.Index
        Next lr
    End If

    ds = Split(CH("SHEET_VAN_BAN"), ",")
    For i = 0 To UBound(ds)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbR.Worksheets(Trim$(ds(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            cuoi = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If cuoi < 2 Then cuoi = 2
            v = ws.Range(ws.Cells(1, 1), ws.Cells(cuoi, 30)).Value
            cSo = TimCot(v, 8, U("số văn bản"), U("số chỉ thị"), U("mã hồ sơ"), U("số ncr"), U("số"))
            cND = TimCot(v, 8, U("nội dung"), U("danh mục"))
            cNg = TimCot(v, 8, U("ngày phát hành"), U("ngày nhận"), U("ngày"))
            cNguoi = TimCot(v, 8, U("người ký"), U("người gửi"), U("người phát hành"), U("phụ trách"))
            cTT = TimCot(v, 8, U("tình trạng"))
            cPH = TimCot(v, 8, U("ngày phản hồi"), U("phản hồi"))
            hdr = 3
            If cSo = 0 Or cND = 0 Then GoTo TIEP_SHEET
            For r = hdr + 1 To UBound(v, 1)
                so = LamSach(Chu(v(r, cSo)))
                nd = LamSach(Chu(v(r, cND)))
                If Len(nd) > 0 And so <> "..." And Len(so) > 0 And LCase$(so) <> LCase$(LamSach(Chu(v(1, cSo)))) Then
                    khoa = LCase$(ws.Name) & "|" & UCase$(so) & "|" & LCase$(nd)
                    If dict.Exists(khoa) Then
                        Set lr = lo.ListRows(dict(khoa))
                    Else
                        Set lr = lo.ListRows.Add
                        dict.Add khoa, lr.Index
                    End If
                    With lr.Range
                        .Cells(1, 2).Value = ws.Name
                        .Cells(1, 3).Value = so
                        .Cells(1, 4).Value = nd
                        If cNg > 0 Then If Not IsEmpty(DocNgay(v(r, cNg))) Then .Cells(1, 5).Value = DocNgay(v(r, cNg))
                        If cNguoi > 0 Then .Cells(1, 6).Value = LamSach(Chu(v(r, cNguoi)))
                        If cTT > 0 Then .Cells(1, 7).Value = LamSach(Chu(v(r, cTT)))
                        If cPH > 0 Then .Cells(1, 8).Value = GopO(v, r, cPH)
                    End With
                    n = n + 1
                End If
            Next r
        End If
TIEP_SHEET:
    Next i
    NhapVanBan = n
End Function

'------------------------------------------------------------------------------
' 2. KIEM MA SO
'------------------------------------------------------------------------------
Public Sub TD_KiemMa()
    On Error GoTo LOI
    Dim n As Long
    n = KiemMaNoiBo(True)
    Hop U("Đã kiểm ") & Bang().ListRows.Count & U(" dòng. Có ") & n & U(" dòng cần xem lại (tô đỏ ở cột Kiểm mã)."), _
        IIf(n > 0, vbExclamation, vbInformation)
    Exit Sub
LOI:
    BaoLoi "KiemMa"
End Sub

Private Function KiemMaNoiBo(ByVal ghiNhatKy As Boolean) As Long
    Dim lo As ListObject, v As Variant, r As Long, n As Long
    Dim dem As Object, dongDau As Object, k As String, s As String
    Dim ma As String, goc As String, nd As String, loai As String
    Dim ng As String, nh As String, so As Long, du As String, d As Date
    Dim t As cTatManHinh
    Set t = New cTatManHinh

    Set lo = Bang()
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value
    Set dem = CreateObject("Scripting.Dictionary")
    Set dongDau = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(v, 1)
        ma = UCase$(LamSach(Chu(v(r, C_MA))))
        If Len(ma) > 0 Then
            k = ma & "|" & Chu(v(r, C_REV))
            dem(k) = dem(k) + 1
            If Not dongDau.Exists(k) Then dongDau.Add k, r
        End If
    Next r

    For r = 1 To UBound(v, 1)
        s = ""
        goc = Chu(v(r, C_MA))
        ma = LamSach(goc)
        nd = UCase$(Chu(v(r, C_ND)))
        loai = UCase$(Chu(v(r, C_LOAI)))
        If Len(ma) = 0 Then
            s = U("Chưa có mã")
        Else
            If ma <> goc Then v(r, C_MA) = ma: s = s & U("Đã bỏ khoảng trắng/tab thừa; ")
            If Not TachMa(ma, ng, nh, so, du) Then
                s = s & U("Sai cấu trúc yymmdd-...-nnn; ")
            Else
                On Error Resume Next
                d = DateSerial(2000 + CLng(Left$(ng, 2)), CLng(Mid$(ng, 3, 2)), CLng(Right$(ng, 2)))
                If Err.Number <> 0 Or Format$(d, "yymmdd") <> ng Then s = s & U("6 số đầu không phải ngày; ")
                On Error GoTo 0
                If InStr(UCase$(nh), "COA") = 0 Then s = s & U("Thiếu mã dự án COA; ")
                If Len(loai) > 0 And loai <> "RFA" And loai <> "HSE" And loai <> "REP" Then
                    If InStr("-" & UCase$(Replace(nh, ".", "-")) & "-", "-" & loai & "-") = 0 Then _
                        s = s & U("Mã không chứa loại ") & loai & "; "
                End If
                If InStr(UCase$(nh), "HTKT") > 0 And (InStr(nd, "REHAUS") > 0 Or nd Like "*RE[45]*") Then _
                    s = s & U("Nội dung nhà phố (Rehaus) nhưng mã là HTKT; ")
                If InStr(UCase$(nh), ".NP") > 0 And nd Like "*(LD-*" Then _
                    s = s & U("Nội dung công viên nhưng mã là NP; ")
            End If
            k = UCase$(ma) & "|" & Chu(v(r, C_REV))
            If dem(k) > 1 And dongDau(k) <> r Then
                s = s & U("TRÙNG MÃ với dòng ") & (dongDau(k) + 1) & "; "
            ElseIf dem(k) > 1 Then
                s = s & U("TRÙNG MÃ (") & dem(k) & U(" dòng); ")
            End If
        End If
        If Len(s) = 0 Then s = "OK" Else n = n + 1
        If Right$(s, 2) = "; " Then s = Left$(s, Len(s) - 2)
        v(r, C_KIEM) = s
    Next r

    Dim vMa As Variant, vK As Variant
    ReDim vMa(1 To UBound(v, 1), 1 To 1): ReDim vK(1 To UBound(v, 1), 1 To 1)
    For r = 1 To UBound(v, 1): vMa(r, 1) = v(r, C_MA): vK(r, 1) = v(r, C_KIEM): Next r
    lo.ListColumns(C_MA).DataBodyRange.Value = vMa
    lo.ListColumns(C_KIEM).DataBodyRange.Value = vK
    If ghiNhatKy Then GhiLog "KiemMa", "", n & " dong can xem"
    KiemMaNoiBo = n
End Function

'------------------------------------------------------------------------------
' 3. MA MOI / TRINH LAI
'------------------------------------------------------------------------------
Private Function DongDangChon() As ListRow
    Dim lo As ListObject, c As Range
    Set lo = Bang()
    If ActiveSheet.Name <> SH_DS Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set c = ActiveCell
    If Intersect(c, lo.DataBodyRange) Is Nothing Then Exit Function
    Set DongDangChon = lo.ListRows(c.Row - lo.HeaderRowRange.Row)
End Function

Private Function SoTiepTheo(ByVal nhom As String, ByRef duoiMau As String) As Long
    Dim v As Variant, r As Long, ng As String, nh As String, so As Long, du As String, maxSo As Long
    duoiMau = ""
    If Bang().DataBodyRange Is Nothing Then SoTiepTheo = 1: Exit Function
    v = Bang().ListColumns(C_MA).DataBodyRange.Value
    If Not IsArray(v) Then
        Dim v1 As Variant: v1 = v: ReDim v(1 To 1, 1 To 1): v(1, 1) = v1
    End If
    For r = 1 To UBound(v, 1)
        If TachMa(Chu(v(r, 1)), ng, nh, so, du) Then
            If UCase$(nh) = UCase$(nhom) Then
                If so > maxSo Then maxSo = so
                If Len(du) > 0 Then duoiMau = du
            End If
        End If
    Next r
    SoTiepTheo = maxSo + 1
End Function

Public Sub TD_MaMoi()
    On Error GoTo LOI
    Dim lr As ListRow, nhom As String, duoiMau As String, ma As String, so As Long
    Set lr = DongDangChon()
    If lr Is Nothing Then Hop U("Chọn 1 ô trong dòng cần cấp mã (sheet DANH_SACH)."), vbExclamation: Exit Sub
    nhom = LamSach(Chu(lr.Range.Cells(1, C_NHOM).Value))
    If Len(nhom) = 0 Then
        nhom = LamSach(InputBox("Nhap nhom ma (phan giua ngay va so thu tu)." & vbLf & _
                       "Vi du: COA.NP-NT-SDS | COA.HTKT-NT-SDS | COA.NP-NT-MSS | COA-TBC-RFI", "TQN - Ma moi"))
        If Len(nhom) = 0 Then Exit Sub
    End If
    If Len(LamSach(Chu(lr.Range.Cells(1, C_MA).Value))) > 0 Then
        If Hop(U("Dòng này đã có mã ") & lr.Range.Cells(1, C_MA).Value & U(". Cấp mã mới đè lên?"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    so = SoTiepTheo(nhom, duoiMau)
    If Len(duoiMau) > 0 Then duoiMau = DoiRev(duoiMau, 1)
    ma = Format$(Date, "yymmdd") & "-" & nhom & "-" & Format$(so, "000") & duoiMau
    With lr.Range
        .Cells(1, C_MA).Value = ma
        .Cells(1, C_NHOM).Value = nhom
        If Len(Chu(.Cells(1, C_LOAI).Value)) = 0 Then .Cells(1, C_LOAI).Value = LoaiTuMa(ma, "")
        If Len(Chu(.Cells(1, C_REV).Value)) = 0 Then .Cells(1, C_REV).Value = IIf(Len(duoiMau) > 0, "'01", "'00")
        If IsEmpty(.Cells(1, C_NGAY).Value) Then .Cells(1, C_NGAY).Value = Date
    End With
    DatCongThucDong lr
    GhiLog "MaMoi", ma, "OK"
    KiemMaNoiBo False
    Exit Sub
LOI:
    BaoLoi "MaMoi"
End Sub

Private Function DoiRev(ByVal duoi As String, ByVal rev As Long) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(rev ?)(\d+)"
    If re.Test(duoi) Then DoiRev = re.Replace(duoi, "$1" & Format$(rev, "00")) Else DoiRev = duoi
End Function

Public Sub TD_LenRev()
    On Error GoTo LOI
    Dim lr As ListRow, moi As ListRow, c As Long, rev As Long
    Dim ng As String, nh As String, so As Long, du As String, ma As String
    Set lr = DongDangChon()
    If lr Is Nothing Then Hop U("Chọn 1 ô trong dòng cần trình lại (sheet DANH_SACH)."), vbExclamation: Exit Sub
    ma = LamSach(Chu(lr.Range.Cells(1, C_MA).Value))
    If Not TachMa(ma, ng, nh, so, du) Then Hop U("Mã của dòng này sai cấu trúc, không lên rev được."), vbExclamation: Exit Sub
    rev = Val(Chu(lr.Range.Cells(1, C_REV).Value)) + 1
    Set moi = Bang().ListRows.Add(lr.Index + 1)
    For c = 1 To SO_COT
        Select Case c
            Case C_NHOMMAU, C_BUOC, C_HAN, C_MAU
            Case Else: moi.Range.Cells(1, c).Value = lr.Range.Cells(1, c).Value
        End Select
    Next c
    DatCongThucDong moi
    With moi.Range
        .Cells(1, C_MA).Value = Format$(Date, "yymmdd") & "-" & nh & "-" & Format$(so, "000") & DoiRev(du, rev)
        .Cells(1, C_REV).Value = "'" & Format$(rev, "00")
        .Cells(1, C_NGAY).Value = Date
        Dim cc As Variant
        For Each cc In Array(C_CHON, C_TT, C_TVGS, C_CDT, C_TRN, C_TAO, C_FILE, C_IN)
            .Cells(1, cc).Value = ""
        Next cc
        .Cells(1, C_MA).Select
    End With
    GhiLog "LenRev", ma, "rev " & Format$(rev, "00")
    KiemMaNoiBo False
    Exit Sub
LOI:
    BaoLoi "LenRev"
End Sub

'------------------------------------------------------------------------------
' 4. TAO PHIEU WORD (+PDF, +IN)
'------------------------------------------------------------------------------
Private Function DongDanhDau() As Collection
    Dim lr As ListRow, col As Collection, s As String
    Set col = New Collection
    If Not Bang().DataBodyRange Is Nothing Then
        For Each lr In Bang().ListRows
            If Not lr.Range.EntireRow.Hidden Then
                s = LCase$(Trim$(Chu(lr.Range.Cells(1, C_CHON).Value)))
                If Len(s) > 0 And s <> "0" And s <> "false" Then col.Add lr
            End If
        Next lr
    End If
    Set DongDanhDau = col
End Function

Private Function GiaTri(ByVal s As String, ByVal macDinh As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then s = macDinh
    s = Replace(s, vbCrLf, Chr$(11))
    s = Replace(s, vbLf, Chr$(11))
    GiaTri = s
End Function

Private Sub ThayChu(ByVal doc As Object, ByVal khoa As String, ByVal giaTriMoi As String)
    Dim st As Object, rg As Object, dem As Long
    For Each st In doc.StoryRanges
        Do While Not st Is Nothing
            Set rg = st.Duplicate
            With rg.Find
                .ClearFormatting
                .Text = khoa
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = True
                .MatchWildcards = False
            End With
            dem = 0
            Do While rg.Find.Execute
                rg.Text = giaTriMoi
                rg.Collapse wdCollapseEnd
                dem = dem + 1
                If dem > 50 Then Exit Do
            Loop
            Set st = st.NextStoryRange
        Loop
    Next st
End Sub

Private Sub DienChung(ByVal doc As Object)
    Dim dots As String: dots = String$(10, ChrW(8230))
    ThayChu doc, "{{BEN_GUI}}", GiaTri(CH("BEN_GUI"), dots)
    ThayChu doc, "{{GUI_TOI}}", GiaTri(CH("GUI_TOI"), dots)
    ThayChu doc, "{{NGUOI_LAP}}", GiaTri(CH("NGUOI_LAP"), dots)
    ThayChu doc, "{{SO_HOP_DONG}}", GiaTri(CH("SO_HOP_DONG"), ChrW(8230))
End Sub

Private Function MoWord(ByRef taoMoi As Boolean) As Object
    Dim wd As Object
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then Set wd = CreateObject("Word.Application"): taoMoi = True
    wd.Visible = False
    wd.DisplayAlerts = 0
    Set MoWord = wd
End Function

Public Sub TD_TaoPhieu()
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, wd As Object, doc As Object, taoWord As Boolean
    Dim thuMucMau As String, thuMucXuat As String, mau As String, loai As String
    Dim fDocx As String, fPdf As String, ma As String, nd As String, tienTo As String
    Dim nOK As Long, nBo As Long, ghiDe As VbMsgBoxResult, xuatPdf As Boolean, inNgay As Boolean, dem As Long
    Dim dots As String: dots = String$(10, ChrW(8230))

    Set ds = DongDanhDau()
    If ds.Count = 0 Then Hop U("Chưa đánh dấu ""x"" ở cột Chọn cho dòng nào."), vbExclamation: Exit Sub
    thuMucMau = ThuMuc("THU_MUC_MAU", "MAU")
    thuMucXuat = ThuMuc("THU_MUC_XUAT", "PHIEU")
    xuatPdf = CoBat("XUAT_PDF")
    inNgay = CoBat("IN_NGAY")
    If Hop(U("Tạo ") & ds.Count & U(" phiếu Word") & IIf(xuatPdf, " + PDF", "") & IIf(inNgay, U(" và IN NGAY"), "") & _
           vbLf & U("Lưu vào: ") & thuMucXuat, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub

    For Each lr In ds
        mau = thuMucMau & "\MAU_" & MauCua(lr) & ".docx"
        If Len(Dir$(mau)) = 0 Then Err.Raise LOI_TQN + 1, , U("Không thấy file mẫu: ") & mau
    Next lr

    Set wd = MoWord(taoWord)
    ghiDe = vbNo
    For Each lr In ds
        dem = dem + 1
        Application.StatusBar = "TQN: phieu " & dem & "/" & ds.Count
        DoEvents
        With lr.Range
            ma = LamSach(Chu(.Cells(1, C_MA).Value))
            nd = LamSach(Chu(.Cells(1, C_ND).Value))
            loai = UCase$(Chu(.Cells(1, C_LOAI).Value))
            If Len(loai) = 0 Then loai = "KHAC"
            If Len(ma) = 0 Then
                nBo = nBo + 1: GhiLog "TaoPhieu", "(dong " & lr.Index & ")", "Bo qua: chua co ma"
                GoTo TIEP
            End If
            mau = MauCua(lr)
            fDocx = thuMucXuat & "\" & loai & "\" & TenFile(ma) & ".docx"
            fPdf = Left$(fDocx, Len(fDocx) - 5) & ".pdf"
            TaoThuMuc thuMucXuat & "\" & loai
            If Len(Dir$(fDocx)) > 0 And ghiDe <> vbYes Then
                ghiDe = Hop(U("Đã có phiếu ") & ma & "." & vbLf & U("Yes = ghi đè tất cả, No = bỏ qua file đã có, Cancel = dừng"), vbYesNoCancel + vbQuestion)
                If ghiDe = vbCancel Then Exit For
                If ghiDe = vbNo Then nBo = nBo + 1: GoTo TIEP
            End If
            Select Case mau
                Case "RFI": tienTo = CH("TIEN_TO_RFI")
                Case "RFC": tienTo = CH("TIEN_TO_RFC")
                Case Else: tienTo = CH("TIEN_TO_RFA")
            End Select

            Set doc = wd.Documents.Add(Template:=thuMucMau & "\MAU_" & mau & ".docx", Visible:=False)
            DienChung doc
            ThayChu doc, "{{SO_THAM_CHIEU}}", ma
            ThayChu doc, "{{VE_VIEC}}", GiaTri(nd, ChrW(8230))
            ThayChu doc, "{{NOI_DUNG}}", GiaTri(tienTo & nd, ChrW(8230))
            ThayChu doc, "{{NGUOI_NHAN}}", GiaTri(Chu(.Cells(1, C_NHAN).Value), GiaTri(CH("NGUOI_NHAN_MD"), dots))
            ThayChu doc, "{{NGAY_BAN_HANH}}", GiaTri(NgayVN(.Cells(1, C_NGAY).Value), "")
            ThayChu doc, "{{HAN_PHAN_HOI}}", GiaTri(NgayVN(.Cells(1, C_HAN).Value), dots)
            ThayChu doc, "{{CC}}", GiaTri(Chu(.Cells(1, C_CC).Value), CH("CC_MAC_DINH"))
            ThayChu doc, "{{DINH_KEM}}", GiaTri(Chu(.Cells(1, C_DK).Value), "")
            ThayChu doc, "{{AH_TIEN_DO}}", ""
            ThayChu doc, "{{AH_THIET_KE}}", ""
            ThayChu doc, "{{AH_CHI_PHI}}", ""
            doc.SaveAs2 Filename:=fDocx, FileFormat:=wdFormatXMLDocument
            If xuatPdf Then doc.ExportAsFixedFormat OutputFileName:=fPdf, ExportFormat:=wdExportFormatPDF
            If inNgay Then doc.PrintOut Background:=False: .Cells(1, C_IN).Value = Now
            doc.Close SaveChanges:=False
            Set doc = Nothing
            .Cells(1, C_TAO).Value = Now
            .Parent.Hyperlinks.Add Anchor:=.Cells(1, C_FILE), Address:=fDocx, TextToDisplay:=TenFile(ma) & ".docx"
            .Cells(1, C_CHON).Value = ""
            nOK = nOK + 1
            GhiLog "TaoPhieu", ma, fDocx
        End With
TIEP:
    Next lr
    If taoWord Then wd.Quit SaveChanges:=False
    Set wd = Nothing
    Application.StatusBar = False
    LamMoiNoiBo
    Hop U("Đã tạo ") & nOK & U(" phiếu.") & IIf(nBo > 0, vbLf & U("Bỏ qua ") & nBo & U(" dòng (xem NHAT_KY)."), "")
    Exit Sub
LOI:
    Application.StatusBar = False
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If taoWord And Not wd Is Nothing Then wd.Quit SaveChanges:=False
    On Error GoTo 0
    BaoLoi "TaoPhieu"
End Sub

Private Function MauCua(ByVal lr As ListRow) As String
    Dim m As String: m = UCase$(Chu(lr.Range.Cells(1, C_MAU).Value))
    If m <> "RFI" And m <> "RFC" Then m = "RFA"
    MauCua = m
End Function

'------------------------------------------------------------------------------
' 5. PHIEU GIAO NHAN GOP (TRN) cho cac dong danh "x"
'------------------------------------------------------------------------------
Public Sub TD_GiaoNhan()
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, wd As Object, doc As Object, taoWord As Boolean
    Dim tbl As Object, tb As Object, rDong As Long, i As Long, maTRN As String, so As Long, duoi As String
    Dim nhom As String, f As String, fPdf As String, dsMa As String, nguoiNhan As String
    Dim mau As String, thuMucXuat As String
    Dim dots As String: dots = String$(10, ChrW(8230))

    Set ds = DongDanhDau()
    If ds.Count = 0 Then Hop U("Đánh dấu ""x"" các hồ sơ cần giao rồi bấm lại."), vbExclamation: Exit Sub
    mau = ThuMuc("THU_MUC_MAU", "MAU") & "\MAU_TRN.docx"
    If Len(Dir$(mau)) = 0 Then Err.Raise LOI_TQN + 2, , U("Không thấy file mẫu: ") & mau
    nhom = CH("NHOM_MA_TRN")
    If Len(nhom) = 0 Then nhom = "COA-TBC-TML"
    so = SoTRNTiepTheo(nhom)
    maTRN = Format$(Date, "yymmdd") & "-" & nhom & "-" & Format$(so, "000")
    nguoiNhan = Chu(ds(1).Range.Cells(1, C_NHAN).Value)
    If Hop(U("Lập phiếu giao nhận ") & maTRN & U(" cho ") & ds.Count & U(" hồ sơ?"), vbOKCancel + vbQuestion) <> vbOK Then Exit Sub

    Set wd = MoWord(taoWord)
    Set doc = wd.Documents.Add(Template:=mau, Visible:=False)
    DienChung doc
    ThayChu doc, "{{SO_THAM_CHIEU}}", maTRN
    ThayChu doc, "{{VE_VIEC}}", U("Giao nhận ") & ds.Count & U(" hồ sơ")
    ThayChu doc, "{{NGUOI_NHAN}}", GiaTri(nguoiNhan, GiaTri(CH("NGUOI_NHAN_MD"), dots))
    ThayChu doc, "{{NGAY_BAN_HANH}}", Format$(Date, "dd/mm/yyyy")
    ThayChu doc, "{{CC}}", GiaTri(CH("CC_MAC_DINH"), "")
    ThayChu doc, "{{TRN_SL}}", CStr(ds.Count)
    ThayChu doc, "{{DINH_KEM}}", ""

    ' tim bang co dong mau {{TRN_STT}}
    For Each tb In doc.Tables
        If InStr(tb.Range.Text, "{{TRN_STT}}") > 0 Then Set tbl = tb: Exit For
    Next tb
    If tbl Is Nothing Then Err.Raise LOI_TQN + 3, , U("Mẫu TRN thiếu dòng {{TRN_STT}}")
    rDong = tbl.Rows.Count
    For i = 2 To ds.Count: tbl.Rows.Add: Next i
    i = 0
    For Each lr In ds
        i = i + 1
        DatO tbl, rDong + i - 1, 1, CStr(i)
        DatO tbl, rDong + i - 1, 2, LamSach(Chu(lr.Range.Cells(1, C_MA).Value))
        DatO tbl, rDong + i - 1, 3, LamSach(Chu(lr.Range.Cells(1, C_ND).Value))
        DatO tbl, rDong + i - 1, 4, Chu(lr.Range.Cells(1, C_REV).Value)
        DatO tbl, rDong + i - 1, 5, "1"
        DatO tbl, rDong + i - 1, 6, ""
        dsMa = dsMa & IIf(Len(dsMa) > 0, "; ", "") & LamSach(Chu(lr.Range.Cells(1, C_MA).Value))
    Next lr

    thuMucXuat = ThuMuc("THU_MUC_XUAT", "PHIEU") & "\TRN"
    TaoThuMuc thuMucXuat
    f = thuMucXuat & "\" & TenFile(maTRN) & ".docx"
    fPdf = Left$(f, Len(f) - 5) & ".pdf"
    doc.SaveAs2 Filename:=f, FileFormat:=wdFormatXMLDocument
    If CoBat("XUAT_PDF") Then doc.ExportAsFixedFormat OutputFileName:=fPdf, ExportFormat:=wdExportFormatPDF
    If CoBat("IN_NGAY") Then doc.PrintOut Background:=False
    doc.Close SaveChanges:=False: Set doc = Nothing
    If taoWord Then wd.Quit SaveChanges:=False
    Set wd = Nothing

    For Each lr In ds
        lr.Range.Cells(1, C_TRN).Value = maTRN
        lr.Range.Cells(1, C_CHON).Value = ""
    Next lr
    Dim loGN As ListObject, lrGN As ListRow
    Set loGN = ThisWorkbook.Worksheets(SH_GN).ListObjects(TEN_BANG_GN)
    Set lrGN = loGN.ListRows.Add
    With lrGN.Range
        .Cells(1, 1).Value = maTRN
        .Cells(1, 2).Value = Date
        .Cells(1, 3).Value = CH("GUI_TOI")
        .Cells(1, 4).Value = ds.Count
        .Cells(1, 5).Value = dsMa
        .Parent.Hyperlinks.Add Anchor:=.Cells(1, 6), Address:=f, TextToDisplay:=TenFile(maTRN) & ".docx"
    End With
    GhiLog "GiaoNhan", maTRN, ds.Count & " ho so"
    LamMoiNoiBo
    Hop U("Đã lập phiếu giao nhận ") & maTRN & U(" (") & ds.Count & U(" hồ sơ).") & vbLf & f
    Exit Sub
LOI:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If taoWord And Not wd Is Nothing Then wd.Quit SaveChanges:=False
    On Error GoTo 0
    BaoLoi "GiaoNhan"
End Sub

Private Sub DatO(ByVal tbl As Object, ByVal r As Long, ByVal c As Long, ByVal s As String)
    tbl.Cell(r, c).Range.Text = s
End Sub

Private Function SoTRNTiepTheo(ByVal nhom As String) As Long
    Dim lo As ListObject, lr As ListRow, ng As String, nh As String, so As Long, du As String, m As Long
    Set lo = ThisWorkbook.Worksheets(SH_GN).ListObjects(TEN_BANG_GN)
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            If TachMa(Chu(lr.Range.Cells(1, 1).Value), ng, nh, so, du) Then
                If UCase$(nh) = UCase$(nhom) And so > m Then m = so
            End If
        Next lr
    End If
    SoTRNTiepTheo = m + 1
End Function

'------------------------------------------------------------------------------
' 6. IN
'------------------------------------------------------------------------------
Public Sub TD_InPhieu()
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, wd As Object, doc As Object, taoWord As Boolean
    Dim f As String, nIn As Long, nThieu As Long
    Set ds = DongDanhDau()
    If ds.Count = 0 Then Hop U("Chưa đánh dấu ""x"" ở cột Chọn cho dòng nào."), vbExclamation: Exit Sub
    If Hop(U("In ") & ds.Count & U(" phiếu ra máy in mặc định?"), vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    Set wd = MoWord(taoWord)
    For Each lr In ds
        f = ""
        If lr.Range.Cells(1, C_FILE).Hyperlinks.Count > 0 Then f = lr.Range.Cells(1, C_FILE).Hyperlinks(1).Address
        If Len(f) = 0 Or Len(Dir$(f)) = 0 Then
            nThieu = nThieu + 1
            GhiLog "InPhieu", Chu(lr.Range.Cells(1, C_MA).Value), "Chua co file phieu"
        Else
            Set doc = wd.Documents.Open(Filename:=f, ReadOnly:=True, Visible:=False)
            doc.PrintOut Background:=False
            doc.Close SaveChanges:=False
            Set doc = Nothing
            lr.Range.Cells(1, C_IN).Value = Now
            lr.Range.Cells(1, C_CHON).Value = ""
            nIn = nIn + 1
            GhiLog "InPhieu", Chu(lr.Range.Cells(1, C_MA).Value), "Da gui lenh in"
        End If
    Next lr
    If taoWord Then wd.Quit SaveChanges:=False
    LamMoiNoiBo
    Hop U("Đã gửi in ") & nIn & U(" phiếu.") & IIf(nThieu > 0, vbLf & nThieu & U(" dòng chưa có file phiếu: bấm Tạo phiếu trước."), "")
    Exit Sub
LOI:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If taoWord And Not wd Is Nothing Then wd.Quit SaveChanges:=False
    On Error GoTo 0
    BaoLoi "InPhieu"
End Sub

Public Sub TD_InDanhSach()
    On Error GoTo LOI
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets(SH_DS)
    Set lo = Bang()
    With ws.PageSetup
        .PrintArea = ws.Range(lo.HeaderRowRange.Cells(1, 1), lo.Range.Cells(lo.Range.Rows.Count, C_KIEM)).Address
        .PrintTitleRows = lo.HeaderRowRange.EntireRow.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "&P/&N"
        .LeftFooter = "TQN - &D"
    End With
    ws.PrintPreview
    Exit Sub
LOI:
    BaoLoi "InDanhSach"
End Sub

Public Sub TD_InNhanGay()
    ' Nhan gay / bia cho cac dong danh "x": 2 nhan moi hang, xem truoc khi in
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, ws As Worksheet, i As Long, r As Long, c As Long
    Set ds = DongDanhDau()
    If ds.Count = 0 Then Hop U("Đánh dấu ""x"" các hồ sơ cần in nhãn."), vbExclamation: Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_NHAN)
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 48: ws.Columns("B").ColumnWidth = 3: ws.Columns("C").ColumnWidth = 48
    For Each lr In ds
        r = (i \ 2) * 5 + 1: c = IIf(i Mod 2 = 0, 1, 3)
        With ws.Cells(r, c)
            .Value = U("DỰ ÁN COASTAL QUẢNG NGÃI")
            .Font.Size = 9: .Font.Color = RGB(31, 78, 121)
        End With
        With ws.Cells(r + 1, c)
            .Value = LamSach(Chu(lr.Range.Cells(1, C_MA).Value))
            .Font.Size = 14: .Font.Bold = True
        End With
        With ws.Cells(r + 2, c)
            .Value = LamSach(Chu(lr.Range.Cells(1, C_ND).Value))
            .Font.Size = 10: .WrapText = True
        End With
        ws.Rows(r + 2).RowHeight = 42
        ws.Cells(r + 3, c).Value = "Rev " & Chu(lr.Range.Cells(1, C_REV).Value) & "   |   " & NgayVN(lr.Range.Cells(1, C_NGAY).Value)
        ws.Cells(r + 3, c).Font.Size = 9
        With ws.Range(ws.Cells(r, c), ws.Cells(r + 3, c))
            .BorderAround xlContinuous, xlMedium, , RGB(31, 78, 121)
            .IndentLevel = 1
        End With
        i = i + 1
    Next lr
    With ws.PageSetup
        .PrintArea = ws.Range("A1", ws.Cells(r + 3, 3)).Address
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    ws.Activate
    ws.PrintPreview
    Exit Sub
LOI:
    BaoLoi "InNhanGay"
End Sub

Public Sub TD_MoThuMuc()
    Dim p As String
    p = ThuMuc("THU_MUC_XUAT", "PHIEU")
    If Len(Dir$(p, vbDirectory)) = 0 Then TaoThuMuc p
    Shell "explorer.exe """ & p & """", vbNormalFocus
End Sub

Public Sub TD_BoChon()
    If Bang().DataBodyRange Is Nothing Then Exit Sub
    Bang().ListColumns(C_CHON).DataBodyRange.ClearContents
End Sub

'------------------------------------------------------------------------------
' 7. LOC THEO NHOM MAU
'------------------------------------------------------------------------------
Private Sub LocNhom(ByVal tenTen As String)
    Dim lo As ListObject: Set lo = Bang()
    ThisWorkbook.Worksheets(SH_DS).Activate
    If lo.ShowAutoFilter = False Then lo.ShowAutoFilter = True
    lo.Range.AutoFilter Field:=C_NHOMMAU, Criteria1:=CH(tenTen)
End Sub
Public Sub TD_LocTrinhDuyet(): LocNhom "TEN_NHOM_1": End Sub
Public Sub TD_LocHoiDoi(): LocNhom "TEN_NHOM_2": End Sub
Public Sub TD_LocGiaoNhan(): LocNhom "TEN_NHOM_3": End Sub
Public Sub TD_LocHienTruong(): LocNhom "TEN_NHOM_4": End Sub
Public Sub TD_BoLoc()
    Dim lo As ListObject: Set lo = Bang()
    ThisWorkbook.Worksheets(SH_DS).Activate
    If lo.ShowAutoFilter Then lo.AutoFilter.ShowAllData
End Sub

'------------------------------------------------------------------------------
' 8. LAM MOI BANG DIEU KHIEN: "Viec hom nay"
'------------------------------------------------------------------------------
Public Sub TD_LamMoi()
    On Error GoTo LOI
    LamMoiNoiBo
    ThisWorkbook.Worksheets(SH_BDK).Activate
    Exit Sub
LOI:
    BaoLoi "LamMoi"
End Sub

Private Sub LamMoiNoiBo()
    ' liet ke ho so qua han / den han trong 2 ngay toi (chua co chu "duyet" o Tinh trang), sap theo han
    Dim ws As Worksheet, lo As ListObject, v As Variant, r As Long, n As Long, i As Long, j As Long
    Dim han As Variant, tt As String, buoc As String, kq() As Variant, tmp As Variant, o As Range
    Dim duyet As String: duyet = U("duyệt")
    Set ws = ThisWorkbook.Worksheets(SH_BDK)
    Set o = ws.Range("VIEC_HOM_NAY")          ' o dau vung (tieu de nam ngay tren)
    ws.Range(o, o.Offset(29, 5)).ClearContents
    Set lo = Bang()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    v = lo.DataBodyRange.Value
    ReDim kq(1 To UBound(v, 1), 1 To 6)
    For r = 1 To UBound(v, 1)
        han = DocNgay(v(r, C_HAN))
        tt = LCase$(Chu(v(r, C_TT)))
        buoc = Chu(v(r, C_BUOC))
        If Not IsEmpty(han) And InStr(tt, duyet) = 0 Then
            If han <= Date + 2 Then
                n = n + 1
                kq(n, 1) = han
                kq(n, 2) = IIf(han < Date, U("Quá hạn ") & (Date - han) & U(" ngày"), IIf(han = Date, U("Hạn hôm nay"), U("Còn ") & (han - Date) & U(" ngày")))
                kq(n, 3) = Chu(v(r, C_MA))
                kq(n, 4) = Chu(v(r, C_ND))
                kq(n, 5) = Chu(v(r, C_TT))
                kq(n, 6) = buoc
            End If
        End If
    Next r
    ' sap tang theo han (it dong, sap noi bot la du)
    For i = 1 To n - 1
        For j = i + 1 To n
            If kq(j, 1) < kq(i, 1) Then
                For r = 1 To 6: tmp = kq(i, r): kq(i, r) = kq(j, r): kq(j, r) = tmp: Next r
            End If
        Next j
    Next i
    If n > 30 Then n = 30
    For i = 1 To n
        For r = 1 To 6: o.Offset(i - 1, r - 1).Value = kq(i, r): Next r
    Next i
    If n = 0 Then o.Offset(0, 1).Value = U("Không có hồ sơ quá hạn hoặc sắp đến hạn.")
    ws.Range("CAP_NHAT_LUC").Value = Now
End Sub

'------------------------------------------------------------------------------
' 9. CAI DAT NUT (chay 1 lan): nut mau theo nhom
'------------------------------------------------------------------------------
Private Sub Nut(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                ByVal chu As String, ByVal macro As String, ByVal mau As Long, Optional ByVal dam As Boolean = False)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 30)
    s.Name = "TQN_" & macro
    s.Adjustments.Item(1) = 0.18
    s.Fill.ForeColor.RGB = mau
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    With s.TextFrame2
        .TextRange.Text = chu
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = IIf(dam, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 4: .MarginRight = 4
    End With
    s.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
    s.Placement = xlFreeFloating
End Sub

Public Sub TD_CaiDat()
    On Error GoTo LOI
    Dim ws As Worksheet, s As Shape, x As Double, y As Double, y2 As Double
    Dim NAVY As Long, TEAL As Long, XANH As Long, VANG As Long, TIM As Long, XAM As Long, DO_ As Long
    NAVY = RGB(27, 63, 122): TEAL = RGB(15, 110, 86): XANH = RGB(24, 95, 165)
    VANG = RGB(133, 79, 11): TIM = RGB(83, 74, 183): XAM = RGB(95, 94, 90): DO_ = RGB(163, 45, 45)

    Set ws = ThisWorkbook.Worksheets(SH_BDK)
    For Each s In ws.Shapes
        If Left$(s.Name, 4) = "TQN_" Then s.Delete
    Next s
    x = ws.Range("NUT_BAT_DAU").Left: y = ws.Range("NUT_BAT_DAU").Top: y2 = y + 38

    ' hang 1: quy trinh chinh
    Nut ws, x, y, 130, U("1  Nhập Register"), "TD_NhapRegister", NAVY
    Nut ws, x + 136, y, 110, U("2  Kiểm mã"), "TD_KiemMa", NAVY
    Nut ws, x + 252, y, 120, U("3  Cấp mã mới"), "TD_MaMoi", NAVY
    Nut ws, x + 378, y, 120, U("4  Trình lại rev+1"), "TD_LenRev", DO_
    Nut ws, x + 504, y, 130, U("5  TẠO PHIẾU"), "TD_TaoPhieu", TEAL, True
    Nut ws, x + 640, y, 130, U("6  Phiếu giao nhận"), "TD_GiaoNhan", VANG, True
    ' hang 2: in + tien ich
    Nut ws, x, y2, 130, U("In phiếu (dòng x)"), "TD_InPhieu", XAM
    Nut ws, x + 136, y2, 110, U("In danh sách"), "TD_InDanhSach", XAM
    Nut ws, x + 252, y2, 120, U("In nhãn gáy"), "TD_InNhanGay", XAM
    Nut ws, x + 378, y2, 120, U("Làm mới"), "TD_LamMoi", XAM
    Nut ws, x + 504, y2, 130, U("Mở thư mục phiếu"), "TD_MoThuMuc", XAM
    Nut ws, x + 640, y2, 130, U("Bỏ hết dấu x"), "TD_BoChon", XAM

    ' nut loc nhom: dat tren 4 o nhom mau
    Nut ws, ws.Range("O_NHOM_1").Left, ws.Range("O_NHOM_1").Offset(3, 0).Top, ws.Range("O_NHOM_1").Resize(1, 2).Width - 6, U("Lọc nhóm này"), "TD_LocTrinhDuyet", TEAL
    Nut ws, ws.Range("O_NHOM_2").Left, ws.Range("O_NHOM_2").Offset(3, 0).Top, ws.Range("O_NHOM_2").Resize(1, 2).Width - 6, U("Lọc nhóm này"), "TD_LocHoiDoi", XANH
    Nut ws, ws.Range("O_NHOM_3").Left, ws.Range("O_NHOM_3").Offset(3, 0).Top, ws.Range("O_NHOM_3").Resize(1, 2).Width - 6, U("Lọc nhóm này"), "TD_LocGiaoNhan", VANG
    Nut ws, ws.Range("O_NHOM_4").Left, ws.Range("O_NHOM_4").Offset(3, 0).Top, ws.Range("O_NHOM_4").Resize(1, 2).Width - 6, U("Lọc nhóm này"), "TD_LocHienTruong", TIM

    ' nut tren DANH_SACH
    Set ws = ThisWorkbook.Worksheets(SH_DS)
    For Each s In ws.Shapes
        If Left$(s.Name, 4) = "TQN_" Then s.Delete
    Next s
    GhiLog "CaiDat", "", "Da tao nut"
    LamMoiNoiBo
    ThisWorkbook.Worksheets(SH_BDK).Activate
    Hop U("Đã tạo nút trên BANG_DIEU_KHIEN.") & vbLf & U("Bước tiếp: điền CAU_HINH rồi bấm 1 Nhập Register.")
    Exit Sub
LOI:
    BaoLoi "CaiDat"
End Sub
