Attribute VB_Name = "modDeclarations"
Public IncomingServer As String, Password As String, Username As String
Public mStrData As String, mErrorString As String
Public aCommand As Long, TotalMessages As Long, TotalSize As Long, mInterval As Long
Public mAlert As Boolean
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_MODIFY = &H1
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Sub mConnect()
    If frmPOP.Winsock1.State <> sckClosed Then frmPOP.Winsock1.Close
    frmPOP.Winsock1.Connect IncomingServer, frmPOP.txtPort.Text
End Sub
Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Call mConnect
    If Ret <> sOld Then
        sOld = Ret
        sSave = sSave + sOld
    End If
End Sub
