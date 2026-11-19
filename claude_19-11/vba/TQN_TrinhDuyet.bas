Attribute VB_Name = "TQN_TrinhDuyet"
'==============================================================================
' TQN - QUAN LY HO SO TRINH DUYET (RFA / RFI / RFC / SDS / MSS ...)
' Du an Coastal Quang Ngai - TBCONS
' Ban 1.0 - 19/11/2026
'
' Nut tren sheet BANG_DIEU_KHIEN:
'   TD_NhapRegister   : lay danh sach ho so tu file Register (SD, M0S, MAT, OTS, RFA, RFI, RFC)
'   TD_KiemMa         : kiem ma so theo quy dinh ghi ma COA (trung ma, sai goi, sai loai, khoang trang)
'   TD_MaMoi          : tao ma so tiep theo cho dong dang chon
'   TD_LenRev         : tao dong trinh lai (rev + 1) tu dong dang chon
'   TD_TaoPhieu       : dien phieu Word dung mau cho cac dong danh "x" o cot Chon (+ PDF, + in)
'   TD_InPhieu        : in lai cac phieu da tao cua cac dong danh "x"
'   TD_InDanhSach     : xem truoc / in bang danh sach dang loc
'   TD_MoThuMuc       : mo thu muc chua phieu
'   TD_BoChon         : xoa het dau "x"
'   TD_CaiDat         : (chay 1 lan) tao cac nut bam
'
' Khong can bat "Trust access to VBA project". Word goi kieu late binding.
'==============================================================================
Option Explicit

Private Const TEN_BANG As String = "tblPhieu"
Private Const SH_DS As String = "DANH_SACH"
Private Const SH_BDK As String = "BANG_DIEU_KHIEN"
Private Const SH_LOG As String = "NHAT_KY"

' Cot trong bang tblPhieu
Private Const C_CHON As Long = 1
Private Const C_NGUON As Long = 2
Private Const C_LOAI As Long = 3
Private Const C_MAU As Long = 4
Private Const C_NHOM As Long = 5
Private Const C_MA As Long = 6
Private Const C_REV As Long = 7
Private Const C_ND As Long = 8
Private Const C_NGUOI As Long = 9
Private Const C_TT As Long = 10
Private Const C_NGAY As Long = 11
Private Const C_HAN As Long = 12
Private Const C_NHAN As Long = 13
Private Const C_CC As Long = 14
Private Const C_DK As Long = 15
Private Const C_KIEM As Long = 16
Private Const C_TAO As Long = 17
Private Const C_FILE As Long = 18
Private Const C_IN As Long = 19
Private Const SO_COT As Long = 19

' Word
Private Const wdFormatXMLDocument As Long = 16
Private Const wdExportFormatPDF As Long = 17
Private Const wdFindStop As Long = 0
Private Const wdCollapseEnd As Long = 0

Private Const LOI_TQN As Long = vbObjectError + 700

'------------------------------------------------------------------------------
' TIEN ICH CHUNG
'------------------------------------------------------------------------------
Private Function Bang() As ListObject
    Set Bang = ThisWorkbook.Worksheets(SH_DS).ListObjects(TEN_BANG)
End Function

Private Function CH(ByVal ten As String) As String
    ' doc o cau hinh theo ten (Name) tren sheet CAU_HINH
    On Error Resume Next
    CH = Trim$(CStr(ThisWorkbook.Names(ten).RefersToRange.Value))
    On Error GoTo 0
End Function

Private Function CoBat(ByVal ten As String) As Boolean
    ' "Co" / "Yes" / "1" / "x" = bat ; "Khong" / trong = tat
    Dim s As String: s = LCase$(Left$(CH(ten), 1))
    CoBat = (s = "c" Or s = "y" Or s = "1" Or s = "x")
End Function

Private Function Chu(ByVal v As Variant) As String
    ' doi gia tri o thanh chuoi, an toan voi o loi (#N/A...)
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Chu = "" Else Chu = CStr(v)
End Function

Private Function ThuMuc(ByVal ten As String, ByVal macDinh As String) As String
    Dim s As String: s = CH(ten)
    If Len(s) = 0 Then s = ThisWorkbook.Path & "\" & macDinh
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    ThuMuc = s
End Function

Private Sub TaoThuMuc(ByVal p As String)
    ' tao thu muc nhieu cap
    Dim parts() As String, i As Long, cur As String
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

Public Function LamSach(ByVal s As String) As String
    ' bo tab, khoang trang cung, khoang trang thua, khoang trang quanh dau "-"
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
    ' tra ve Date hoac Empty; nhan so serial hoac chuoi dd/mm/yyyy
    Dim s As String, p() As String
    DocNgay = Empty
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsDate(v) And VarType(v) = vbDate Then DocNgay = CDate(v): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then DocNgay = CDate(CDbl(v))
        Exit Function
    End If
    s = Trim$(CStr(v))
    If s Like "#*/#*/####" Then
        p = Split(s, "/")
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
    ' Loai ho so lay tu ma (quy dinh COA muc III.1.3 va III.1.6)
    Dim m As String, k As Variant
    m = "-" & UCase$(LamSach(ma)) & "-"
    m = Replace(m, " ", "-")
    For Each k In Array("RFI", "RFC")
        If InStr(m, "-" & k & "-") > 0 Then LoaiTuMa = k: Exit Function
    Next k
    For Each k In Array("SDS", "MSS", "MAS", "OTS", "COS", "SCH", "RFP", "RFO")
        If InStr(m, "-" & k & "-") > 0 Then LoaiTuMa = k: Exit Function
    Next k
    If InStr(m, "-RFA-") > 0 Then LoaiTuMa = "RFA": Exit Function
    Select Case UCase$(nguon)
        Case "SD": LoaiTuMa = "SDS"
        Case "M0S", "MOS", "MS": LoaiTuMa = "MSS"
        Case "MAT": LoaiTuMa = "MAS"
        Case "OTS": LoaiTuMa = "OTS"
        Case "RFI": LoaiTuMa = "RFI"
        Case "RFC": LoaiTuMa = "RFC"
        Case Else: LoaiTuMa = "RFA"
    End Select
End Function

Public Function MauTuLoai(ByVal loai As String) As String
    Select Case UCase$(loai)
        Case "RFI": MauTuLoai = "RFI"
        Case "RFC": MauTuLoai = "RFC"
        Case Else: MauTuLoai = "RFA"
    End Select
End Function

Private Function TachMa(ByVal ma As String, ByRef ngay As String, ByRef nhom As String, _
                        ByRef so As Long, ByRef duoi As String) As Boolean
    ' ma = yymmdd-<nhom>-nnn<duoi>, vi du 261117-COA.NP-NT-SDS-001 / 261116-COA-TBC-RFI-001 rev01
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
    MsgBox "Loi " & Err.Number & ": " & Err.Description, vbExclamation, "TQN - " & viec
End Sub

Private Sub DatCongThucDong(ByVal lr As ListRow)
    lr.Range.Cells(1, C_MAU).FormulaR1C1 = "=IF(RC[-1]=""RFI"",""RFI"",IF(RC[-1]=""RFC"",""RFC"",""RFA""))"
    lr.Range.Cells(1, C_HAN).FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]+SO_NGAY_PH)"
End Sub

'------------------------------------------------------------------------------
' 1. NHAP TU FILE REGISTER
'------------------------------------------------------------------------------
Public Sub TD_NhapRegister()
    On Error GoTo LOI
    Dim duongDan As String, wbR As Workbook, ws As Worksheet
    Dim dsSheet() As String, i As Long, r As Long, v As Variant
    Dim lo As ListObject, dict As Object, khoa As String, lr As ListRow
    Dim ma As String, nd As String, nThem As Long, nCapNhat As Long
    Dim daMo As Boolean, ngay As Variant, cuoi As Long
    Dim ng As String, nh As String, so As Long, du As String
    Dim t As cTatManHinh

    duongDan = CH("FILE_REGISTER")
    If Len(duongDan) = 0 Or Len(Dir$(duongDan)) = 0 Then
        v = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                        "Chon file Register (HAUS QUANG NGAI- Register.xlsx)")
        If VarType(v) = vbBoolean Then Exit Sub
        duongDan = CStr(v)
        ThisWorkbook.Names("FILE_REGISTER").RefersToRange.Value = duongDan
    End If

    Set t = New cTatManHinh
    Set lo = Bang()

    ' khoa = ma da lam sach | noi dung (ma bi trung trong Register nen phai kem noi dung)
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For Each lr In lo.ListRows
            khoa = UCase$(LamSach(CStr(lr.Range.Cells(1, C_MA).Value))) & "|" & _
                   LCase$(LamSach(CStr(lr.Range.Cells(1, C_ND).Value)))
            If Not dict.Exists(khoa) Then dict.Add khoa, lr.Index
        Next lr
    End If

    ' mo Register chi doc (neu dang mo thi dung luon)
    For Each wbR In Application.Workbooks
        If LCase$(wbR.FullName) = LCase$(duongDan) Then daMo = True: Exit For
    Next wbR
    If Not daMo Then Set wbR = Workbooks.Open(duongDan, UpdateLinks:=0, ReadOnly:=True)

    dsSheet = Split(CH("SHEET_NGUON"), ",")
    For i = 0 To UBound(dsSheet)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbR.Worksheets(Trim$(dsSheet(i)))
        On Error GoTo LOI
        If Not ws Is Nothing Then
            Application.StatusBar = "TQN: dang doc sheet " & ws.Name & "..."
            cuoi = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            If cuoi < 2 Then cuoi = 2
            v = ws.Range(ws.Cells(1, 1), ws.Cells(cuoi, 9)).Value
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
                        ' ngay trinh: cot I (thuc te) roi den cot H (ke hoach)
                        ngay = DocNgay(v(r, 9))
                        If IsEmpty(ngay) Then ngay = DocNgay(v(r, 8))
                        If Not IsEmpty(ngay) And IsEmpty(.Cells(1, C_NGAY).Value) Then .Cells(1, C_NGAY).Value = ngay
                        If TachMa(ma, ng, nh, so, du) Then .Cells(1, C_NHOM).Value = nh
                    End With
                End If
            Next r
        End If
    Next i

    If Not daMo Then wbR.Close SaveChanges:=False
    Set t = Nothing
    KiemMaNoiBo False
    GhiLog "NhapRegister", duongDan, "Them " & nThem & ", cap nhat " & nCapNhat
    MsgBox "Da nhap tu Register:" & vbLf & "  Them moi: " & nThem & " dong" & vbLf & _
           "  Cap nhat: " & nCapNhat & " dong" & vbLf & "Cot 'Kiem ma' da duoc kiem lai.", vbInformation, "TQN"
    Exit Sub
LOI:
    If Not wbR Is Nothing And Not daMo Then wbR.Close SaveChanges:=False
    BaoLoi "NhapRegister"
End Sub

'------------------------------------------------------------------------------
' 2. KIEM MA SO (theo Quy dinh ghi ma ho so COA)
'------------------------------------------------------------------------------
Public Sub TD_KiemMa()
    On Error GoTo LOI
    Dim n As Long
    n = KiemMaNoiBo(True)
    MsgBox "Da kiem " & Bang().ListRows.Count & " dong. Co " & n & " dong can xem lai (to mau o cot Kiem ma).", _
           vbInformation, "TQN"
    Exit Sub
LOI:
    BaoLoi "KiemMa"
End Sub

Private Function KiemMaNoiBo(ByVal baoTrangThai As Boolean) As Long
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

    ' dem trung: cung ma (bo phan rev) nhung khac rev thi la trinh lai, khong tinh trung
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
            s = "Chua co ma"
        Else
            If ma <> goc Then v(r, C_MA) = ma: s = s & "Da bo khoang trang/tab thua; "
            If Not TachMa(ma, ng, nh, so, du) Then
                s = s & "Sai cau truc yymmdd-...-nnn; "
            Else
                On Error Resume Next
                d = DateSerial(2000 + CLng(Left$(ng, 2)), CLng(Mid$(ng, 3, 2)), CLng(Right$(ng, 2)))
                If Err.Number <> 0 Or Format$(d, "yymmdd") <> ng Then s = s & "6 so dau khong phai ngay; "
                On Error GoTo 0
                If InStr(UCase$(nh), "COA") = 0 Then s = s & "Thieu ma du an COA; "
                If Len(loai) > 0 And loai <> "RFA" Then
                    If InStr("-" & UCase$(Replace(nh, ".", "-")) & "-", "-" & loai & "-") = 0 Then _
                        s = s & "Ma khong chua loai " & loai & "; "
                End If
                ' goi Nha pho (NP) / Ha tang ky thuat (HTKT) theo noi dung
                If InStr(UCase$(nh), "HTKT") > 0 And (InStr(nd, "REHAUS") > 0 Or nd Like "*RE[45]*") Then _
                    s = s & "Noi dung nha pho (Rehaus) nhung ma la HTKT; "
                If InStr(UCase$(nh), ".NP") > 0 And (nd Like "*(LD-*" Or InStr(nd, "CONG VIEN") > 0) Then _
                    s = s & "Noi dung cong vien nhung ma la NP; "
            End If
            k = UCase$(ma) & "|" & Chu(v(r, C_REV))
            If dem(k) > 1 And dongDau(k) <> r Then
                s = s & "TRUNG MA voi dong " & (dongDau(k) + 1) & "; "
            ElseIf dem(k) > 1 Then
                s = s & "TRUNG MA (" & dem(k) & " dong); "
            End If
        End If
        If Len(s) = 0 Then s = "OK" Else n = n + 1
        If Right$(s, 2) = "; " Then s = Left$(s, Len(s) - 2)
        v(r, C_KIEM) = s
    Next r

    ' chi ghi lai 2 cot Ma va Kiem ma (giu cong thuc cac cot khac)
    Dim vMa As Variant, vK As Variant
    ReDim vMa(1 To UBound(v, 1), 1 To 1): ReDim vK(1 To UBound(v, 1), 1 To 1)
    For r = 1 To UBound(v, 1): vMa(r, 1) = v(r, C_MA): vK(r, 1) = v(r, C_KIEM): Next r
    lo.ListColumns(C_MA).DataBodyRange.Value = vMa
    lo.ListColumns(C_KIEM).DataBodyRange.Value = vK
    If baoTrangThai Then GhiLog "KiemMa", "", n & " dong can xem"
    KiemMaNoiBo = n
End Function

'------------------------------------------------------------------------------
' 3. MA MOI / TRINH LAI
'------------------------------------------------------------------------------
Private Function DongDangChon() As ListRow
    Dim lo As ListObject, c As Range
    Set lo = Bang()
    If ActiveSheet.Name <> SH_DS Then Exit Function
    Set c = ActiveCell
    If Intersect(c, lo.DataBodyRange) Is Nothing Then Exit Function
    Set DongDangChon = lo.ListRows(c.Row - lo.HeaderRowRange.Row)
End Function

Public Sub TD_MaMoi()
    ' Dong dang chon: dien Loai + Nhom ma (vd COA.NP-NT-SDS). Macro cap so thu tu tiep theo, ngay hom nay.
    On Error GoTo LOI
    Dim lr As ListRow, lo As ListObject, nhom As String, maxSo As Long, duoiMau As String
    Dim v As Variant, r As Long, ng As String, nh As String, so As Long, du As String, ma As String
    Set lr = DongDangChon()
    If lr Is Nothing Then MsgBox "Chon 1 o trong dong can cap ma (sheet DANH_SACH).", vbExclamation, "TQN": Exit Sub
    nhom = LamSach(CStr(lr.Range.Cells(1, C_NHOM).Value))
    If Len(nhom) = 0 Then
        nhom = InputBox("Nhap nhom ma (phan giua ngay va so thu tu)." & vbLf & _
                        "Vi du: COA.NP-NT-SDS  |  COA.HTKT-NT-SDS  |  COA.NP-NT-MSS  |  COA-TBC-RFI", "TQN - Ma moi")
        nhom = LamSach(nhom)
        If Len(nhom) = 0 Then Exit Sub
    End If
    If Len(LamSach(CStr(lr.Range.Cells(1, C_MA).Value))) > 0 Then
        If MsgBox("Dong nay da co ma " & lr.Range.Cells(1, C_MA).Value & ". Cap ma moi de len?", _
                  vbYesNo + vbQuestion, "TQN") <> vbYes Then Exit Sub
    End If

    Set lo = Bang()
    v = lo.ListColumns(C_MA).DataBodyRange.Value
    duoiMau = ""
    For r = 1 To UBound(v, 1)
        If TachMa(CStr(v(r, 1)), ng, nh, so, du) Then
            If UCase$(nh) = UCase$(nhom) Then
                If so > maxSo Then maxSo = so
                If Len(du) > 0 Then duoiMau = du
            End If
        End If
    Next r
    ' duoi rev (neu nhom nay co dung) dat ve rev01
    If Len(duoiMau) > 0 Then duoiMau = DoiRev(duoiMau, 1)
    ma = Format$(Date, "yymmdd") & "-" & nhom & "-" & Format$(maxSo + 1, "000") & duoiMau

    With lr.Range
        .Cells(1, C_MA).Value = ma
        .Cells(1, C_NHOM).Value = nhom
        If Len(CStr(.Cells(1, C_LOAI).Value)) = 0 Then .Cells(1, C_LOAI).Value = LoaiTuMa(ma, "")
        If Len(CStr(.Cells(1, C_REV).Value)) = 0 Then .Cells(1, C_REV).Value = IIf(Len(duoiMau) > 0, "'01", "'00")
        If IsEmpty(.Cells(1, C_NGAY).Value) Then .Cells(1, C_NGAY).Value = Date
        If Len(.Cells(1, C_MAU).Formula) = 0 Then DatCongThucDong lr
    End With
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
    If re.Test(duoi) Then
        DoiRev = re.Replace(duoi, "$1" & Format$(rev, "00"))
    Else
        DoiRev = duoi
    End If
End Function

Public Sub TD_LenRev()
    ' Tao dong trinh lai tu dong dang chon: rev + 1, ngay hom nay, giu so thu tu
    On Error GoTo LOI
    Dim lr As ListRow, moi As ListRow, c As Long, rev As Long
    Dim ng As String, nh As String, so As Long, du As String, ma As String
    Set lr = DongDangChon()
    If lr Is Nothing Then MsgBox "Chon 1 o trong dong can trinh lai (sheet DANH_SACH).", vbExclamation, "TQN": Exit Sub
    ma = LamSach(CStr(lr.Range.Cells(1, C_MA).Value))
    If Not TachMa(ma, ng, nh, so, du) Then MsgBox "Ma cua dong nay sai cau truc, khong len rev duoc.", vbExclamation, "TQN": Exit Sub
    rev = Val(CStr(lr.Range.Cells(1, C_REV).Value)) + 1

    Set moi = Bang().ListRows.Add(lr.Index + 1)
    For c = 1 To SO_COT
        If c <> C_MAU And c <> C_HAN Then moi.Range.Cells(1, c).Value = lr.Range.Cells(1, c).Value
    Next c
    DatCongThucDong moi
    With moi.Range
        .Cells(1, C_CHON).Value = ""
        .Cells(1, C_MA).Value = Format$(Date, "yymmdd") & "-" & nh & "-" & Format$(so, "000") & DoiRev(du, rev)
        .Cells(1, C_REV).Value = "'" & Format$(rev, "00")
        .Cells(1, C_NGAY).Value = Date
        .Cells(1, C_TT).Value = ""
        .Cells(1, C_TAO).Value = ""
        .Cells(1, C_FILE).Value = ""
        .Cells(1, C_IN).Value = ""
        .Cells(1, C_CHON).Select
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
    Dim lr As ListRow, col As New Collection, s As String
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
    ' thay tat ca {{KHOA}} trong moi phan cua van ban (than, header, footer, khung chu)
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
                rg.Text = giaTriMoi   ' gan truc tiep: khong bi gioi han 255 ky tu nhu Replacement.Text
                rg.Collapse wdCollapseEnd
                dem = dem + 1
                If dem > 50 Then Exit Do
            Loop
            Set st = st.NextStoryRange
        Loop
    Next st
End Sub

Public Sub TD_TaoPhieu()
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, wd As Object, doc As Object, taoWord As Boolean
    Dim thuMucMau As String, thuMucXuat As String, mau As String, loai As String
    Dim fDocx As String, fPdf As String, ma As String, nd As String, tienTo As String
    Dim nOK As Long, nBo As Long, ghiDe As VbMsgBoxResult, xuatPdf As Boolean, inNgay As Boolean
    Dim dots As String: dots = String$(16, ChrW(8230))

    Set ds = DongDanhDau()
    If ds.Count = 0 Then MsgBox "Chua danh dau ""x"" o cot Chon cho dong nao.", vbExclamation, "TQN": Exit Sub
    thuMucMau = ThuMuc("THU_MUC_MAU", "MAU")
    thuMucXuat = ThuMuc("THU_MUC_XUAT", "PHIEU")
    xuatPdf = CoBat("XUAT_PDF")
    inNgay = CoBat("IN_NGAY")
    If MsgBox("Tao " & ds.Count & " phieu Word" & IIf(xuatPdf, " + PDF", "") & IIf(inNgay, " va IN NGAY", "") & _
              vbLf & "Luu vao: " & thuMucXuat, vbOKCancel + vbQuestion, "TQN - Tao phieu") <> vbOK Then Exit Sub

    ' kiem mau truoc khi mo Word
    For Each lr In ds
        mau = thuMucMau & "\MAU_" & MauTuLoai(Chu(lr.Range.Cells(1, C_LOAI).Value)) & ".docx"
        If Len(Dir$(mau)) = 0 Then Err.Raise LOI_TQN + 1, , "Khong thay file mau: " & mau
    Next lr

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo LOI
    If wd Is Nothing Then Set wd = CreateObject("Word.Application"): taoWord = True
    wd.Visible = False
    wd.DisplayAlerts = 0

    ghiDe = vbNo
    Dim dem As Long
    For Each lr In ds
        dem = dem + 1
        Application.StatusBar = "TQN: dang tao phieu " & dem & "/" & ds.Count & "..."
        DoEvents
        With lr.Range
            ma = LamSach(Chu(.Cells(1, C_MA).Value))
            nd = LamSach(Chu(.Cells(1, C_ND).Value))
            loai = UCase$(Chu(.Cells(1, C_LOAI).Value))
            If Len(ma) = 0 Then
                nBo = nBo + 1: GhiLog "TaoPhieu", "(dong " & lr.Index & ")", "Bo qua: chua co ma"
                GoTo TIEP
            End If
            mau = MauTuLoai(loai)
            fDocx = thuMucXuat & "\" & loai & "\" & TenFile(ma) & ".docx"
            fPdf = Left$(fDocx, Len(fDocx) - 5) & ".pdf"
            TaoThuMuc thuMucXuat & "\" & loai
            If Len(Dir$(fDocx)) > 0 And ghiDe <> vbYes Then
                ghiDe = MsgBox("Da co phieu " & ma & "." & vbLf & "Yes = ghi de tat ca, No = bo qua file da co, Cancel = dung", _
                               vbYesNoCancel + vbQuestion, "TQN")
                If ghiDe = vbCancel Then Exit For
                If ghiDe = vbNo Then nBo = nBo + 1: GoTo TIEP
            End If

            Select Case mau
                Case "RFI": tienTo = CH("TIEN_TO_RFI")
                Case "RFC": tienTo = CH("TIEN_TO_RFC")
                Case Else: tienTo = CH("TIEN_TO_RFA")
            End Select

            Set doc = wd.Documents.Add(Template:=thuMucMau & "\MAU_" & mau & ".docx", Visible:=False)
            ThayChu doc, "{{SO_THAM_CHIEU}}", ma
            ThayChu doc, "{{VE_VIEC}}", GiaTri(nd, "...")
            ThayChu doc, "{{NOI_DUNG}}", GiaTri(tienTo & nd, "...")
            ThayChu doc, "{{BEN_GUI}}", GiaTri(CH("BEN_GUI"), dots)
            ThayChu doc, "{{GUI_TOI}}", GiaTri(CH("GUI_TOI"), dots)
            ThayChu doc, "{{NGUOI_LAP}}", GiaTri(CH("NGUOI_LAP"), dots)
            ThayChu doc, "{{NGUOI_NHAN}}", GiaTri(Chu(.Cells(1, C_NHAN).Value), GiaTri(CH("NGUOI_NHAN_MD"), dots))
            ThayChu doc, "{{NGAY_BAN_HANH}}", GiaTri(NgayVN(.Cells(1, C_NGAY).Value), "")
            ThayChu doc, "{{HAN_PHAN_HOI}}", GiaTri(NgayVN(.Cells(1, C_HAN).Value), dots)
            ThayChu doc, "{{CC}}", GiaTri(Chu(.Cells(1, C_CC).Value), CH("CC_MAC_DINH"))
            ThayChu doc, "{{SO_HOP_DONG}}", GiaTri(CH("SO_HOP_DONG"), ChrW(8230))
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
    MsgBox "Da tao " & nOK & " phieu." & IIf(nBo > 0, vbLf & "Bo qua " & nBo & " dong (xem NHAT_KY).", ""), vbInformation, "TQN"
    Exit Sub
LOI:
    Application.StatusBar = False
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If taoWord And Not wd Is Nothing Then wd.Quit SaveChanges:=False
    On Error GoTo 0
    BaoLoi "TaoPhieu"
End Sub

'------------------------------------------------------------------------------
' 5. IN LAI PHIEU DA TAO
'------------------------------------------------------------------------------
Public Sub TD_InPhieu()
    On Error GoTo LOI
    Dim ds As Collection, lr As ListRow, wd As Object, doc As Object, taoWord As Boolean
    Dim f As String, nIn As Long, nThieu As Long

    Set ds = DongDanhDau()
    If ds.Count = 0 Then MsgBox "Chua danh dau ""x"" o cot Chon cho dong nao.", vbExclamation, "TQN": Exit Sub
    If MsgBox("In " & ds.Count & " phieu ra may in mac dinh?", vbOKCancel + vbQuestion, "TQN - In phieu") <> vbOK Then Exit Sub

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo LOI
    If wd Is Nothing Then Set wd = CreateObject("Word.Application"): taoWord = True
    wd.Visible = False

    For Each lr In ds
        f = ""
        If lr.Range.Cells(1, C_FILE).Hyperlinks.Count > 0 Then f = lr.Range.Cells(1, C_FILE).Hyperlinks(1).Address
        If Len(f) = 0 Or Len(Dir$(f)) = 0 Then
            nThieu = nThieu + 1
            GhiLog "InPhieu", CStr(lr.Range.Cells(1, C_MA).Value), "Chua co file phieu - bam Tao phieu truoc"
        Else
            Set doc = wd.Documents.Open(Filename:=f, ReadOnly:=True, Visible:=False)
            doc.PrintOut Background:=False
            doc.Close SaveChanges:=False
            Set doc = Nothing
            lr.Range.Cells(1, C_IN).Value = Now
            lr.Range.Cells(1, C_CHON).Value = ""
            nIn = nIn + 1
            GhiLog "InPhieu", CStr(lr.Range.Cells(1, C_MA).Value), "Da gui lenh in"
        End If
    Next lr
    If taoWord Then wd.Quit SaveChanges:=False
    MsgBox "Da gui in " & nIn & " phieu." & IIf(nThieu > 0, vbLf & nThieu & " dong chua co file phieu.", ""), vbInformation, "TQN"
    Exit Sub
LOI:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If taoWord And Not wd Is Nothing Then wd.Quit SaveChanges:=False
    On Error GoTo 0
    BaoLoi "InPhieu"
End Sub

'------------------------------------------------------------------------------
' 6. IN DANH SACH / MO THU MUC / BO CHON
'------------------------------------------------------------------------------
Public Sub TD_InDanhSach()
    On Error GoTo LOI
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets(SH_DS)
    Set lo = Bang()
    With ws.PageSetup
        .PrintArea = lo.Range.Address
        .PrintTitleRows = lo.HeaderRowRange.EntireRow.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Trang &P/&N"
        .LeftFooter = "TQN - So trinh duyet - in &D"
    End With
    ws.PrintPreview
    Exit Sub
LOI:
    BaoLoi "InDanhSach"
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
' 7. CAI DAT NUT BAM (chay 1 lan sau khi import module)
'------------------------------------------------------------------------------
Public Sub TD_CaiDat()
    On Error GoTo LOI
    Dim ws As Worksheet, b As Object, i As Long, x As Double, y As Double
    Dim ds As Variant
    ds = Array( _
        Array("1. Nhap tu Register", "TD_NhapRegister"), _
        Array("2. Kiem ma so", "TD_KiemMa"), _
        Array("3. Cap ma moi (dong chon)", "TD_MaMoi"), _
        Array("4. Trinh lai: rev + 1", "TD_LenRev"), _
        Array("5. TAO PHIEU (dong x)", "TD_TaoPhieu"), _
        Array("6. In phieu (dong x)", "TD_InPhieu"), _
        Array("7. In danh sach", "TD_InDanhSach"), _
        Array("8. Mo thu muc phieu", "TD_MoThuMuc"), _
        Array("Bo het dau x", "TD_BoChon"))

    Set ws = ThisWorkbook.Worksheets(SH_BDK)
    For Each b In ws.Buttons: b.Delete: Next b
    x = ws.Range("B4").Left: y = ws.Range("B4").Top
    For i = 0 To UBound(ds)
        Set b = ws.Buttons.Add(x + (i Mod 3) * 190, y + (i \ 3) * 40, 180, 32)
        b.Caption = ds(i)(0)
        b.OnAction = "'" & ThisWorkbook.Name & "'!" & ds(i)(1)
        b.Font.Bold = (ds(i)(1) = "TD_TaoPhieu")
    Next i
    ' them nut tren sheet DANH_SACH (o hang tren bang)
    Set ws = ThisWorkbook.Worksheets(SH_DS)
    For Each b In ws.Buttons: b.Delete: Next b
    GhiLog "CaiDat", "", "Da tao nut"
    MsgBox "Da tao nut tren sheet BANG_DIEU_KHIEN.", vbInformation, "TQN"
    Exit Sub
LOI:
    BaoLoi "CaiDat"
End Sub
