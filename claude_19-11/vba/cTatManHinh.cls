VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTatManHinh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TQN - tat man hinh/su kien/tinh toan khi chay, tu khoi phuc khi ket thuc (ke ca khi loi)
Option Explicit

Private mSU As Boolean, mEE As Boolean, mCalc As Long

Private Sub Class_Initialize()
    mSU = Application.ScreenUpdating
    mEE = Application.EnableEvents
    mCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub Class_Terminate()
    Application.Calculation = mCalc
    Application.EnableEvents = mEE
    Application.ScreenUpdating = mSU
    Application.StatusBar = False
End Sub
