VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POP Mail Client"
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkServerInfo 
      Caption         =   "Save Server Info"
      Height          =   285
      Left            =   4950
      TabIndex        =   14
      Top             =   90
      Width           =   2085
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Save Password"
      Height          =   330
      Left            =   3780
      TabIndex        =   13
      Top             =   900
      Width           =   1770
   End
   Begin VB.CheckBox chkUsername 
      Caption         =   "Save Username"
      Height          =   330
      Left            =   3780
      TabIndex        =   12
      Top             =   495
      Width           =   1770
   End
   Begin MSComctlLib.ListView lstMessages 
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr No."
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From:"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   5733
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2925
      Top             =   630
   End
   Begin VB.TextBox txtPort 
      Height          =   330
      Left            =   4230
      TabIndex        =   9
      Top             =   45
      Width           =   645
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5940
      TabIndex        =   7
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   450
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1755
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   900
      Width           =   1950
   End
   Begin VB.TextBox txtUsername 
      Height          =   330
      Left            =   1755
      TabIndex        =   3
      Top             =   495
      Width           =   1950
   End
   Begin VB.TextBox txtIncomingServer 
      Height          =   330
      Left            =   1755
      TabIndex        =   1
      Top             =   90
      Width           =   1995
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   180
      Top             =   405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   4590
      Width           =   6900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   3825
      TabIndex        =   8
      Top             =   135
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password :"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   945
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Username :"
      Height          =   195
      Left            =   705
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Incoming Server :"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   1575
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mNotify As NOTIFYICONDATA, CurMessage As Long, lstMessage As Long
Private Sub cmdClose_Click()
    Me.Hide
End Sub
Private Sub cmdLogin_Click()
    If Len(Trim(txtIncomingServer.Text)) <= 0 Then
        MsgBox "Invalid Incoming Server", vbInformation, "Error"
        txtIncomingServer.SetFocus
        Exit Sub
    Else
        If Len(Trim(txtUsername.Text)) <= 0 Then
            MsgBox "Invalid Username", vbInformation, "Error"
            txtUsername.SetFocus
            Exit Sub
        Else
            If Len(Trim(txtPassword.Text)) <= 0 Then
                MsgBox "Invalid Password", vbInformation, "Error"
                txtPassword.SetFocus
                Exit Sub
            End If
        End If
    End If
    IncomingServer = txtIncomingServer.Text
    Password = txtPassword.Text
    Username = txtUsername.Text
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Connect txtIncomingServer.Text, txtPort.Text
    cmdLogin.Enabled = False
    SetTimer Me.hwnd, 0, mInterval, AddressOf TimerProc
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "One instance of POP already running...", vbSystemModal
        End
    End If
    CurMessage = 1
    Dim lng As Long
    mInterval = 300000
    mAlert = True
    aCommand = 0
    mNotify.cbSize = Len(mNotify)
    mNotify.hwnd = Me.hwnd
    mNotify.uID = 1&
    mNotify.uCallbackMessage = WM_LBUTTONDOWN
    mNotify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mNotify.hIcon = Me.Icon
    mNotify.szTip = "Total Message(s) on the Server is: " & TotalMessages & " Total bytes Consumed: " & TotalSize & Chr(0)
    lng = Shell_NotifyIcon(NIM_ADD, mNotify)
    Call GetSettings
    txtPort.Text = 110
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        Me.Visible = True
    ElseIf Msg = WM_RBUTTONUP Then
        Me.PopupMenu mnuView
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    lng = Shell_NotifyIcon(NIM_ADD, mNotify)
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    ElseIf UnloadMode = 1 Then
        If MsgBox("Are you sure you want to Exit?", vbYesNo) = vbYes Then
            lng = Shell_NotifyIcon(NIM_DELETE, mNotify)
            Unload Me
            Call SaveSettings
            Cancel = False
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub mnuViewExit_Click()
    Dim lng As Long
    lng = Shell_NotifyIcon(NIM_DELETE, mNotify)
    Unload Me
End Sub
Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal
End Sub
Private Sub Timer1_Timer()
    mNotify.szTip = "Total Message(s) on the Server is: " & TotalMessages & " Total bytes Consumed: " & TotalSize & Chr(0)
    lng = Shell_NotifyIcon(NIM_MODIFY, mNotify)
End Sub
Private Sub txtIncomingServer_GotFocus()
    If Len(txtIncomingServer.Text) > 0 Then
        txtIncomingServer.SelStart = 0
        txtIncomingServer.SelLength = Len(txtIncomingServer.Text)
    End If
End Sub
Private Sub txtPassword_GotFocus()
    If Len(txtPassword.Text) > 0 Then
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub
Private Sub txtPort_GotFocus()
    If Len(txtPort.Text) > 0 Then
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort.Text)
    End If
End Sub
Private Sub txtUsername_GotFocus()
    If Len(txtUsername.Text) > 0 Then
        txtUsername.SelStart = 0
        txtUsername.SelLength = Len(txtUsername.Text)
    End If
End Sub
Private Sub Winsock1_Connect()
    frmPOP.lblStatus.Caption = "Connected to " & frmPOP.Winsock1.RemoteHost
    frmPOP.Winsock1.SendData "USER " & Username & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TempString As String, mArray() As String, lstItem As ListItem, From As String, Subject As String
    Winsock1.GetData mStrData, vbString
    If Left(mStrData, 3) = "+OK" Then
        If aCommand = 0 Then
            aCommand = 1
            lblStatus.Caption = "Sending Username...."
        
        ElseIf aCommand = 1 Then
            SendCommand "PASS " & Password & vbCrLf
            aCommand = 2
            lblStatus.Caption = "Sending Password...."
        
        ElseIf aCommand = 2 Then
            SendCommand "STAT" & vbCrLf
            aCommand = 3
            lblStatus.Caption = "Getting Total Number of Messages from the server...."
        
        ElseIf aCommand = 3 Then
            TempString = Mid(mStrData, 4)
            mArray = Split(TempString, " ")
            If mAlert = True Then
                If TotalMessages <> mArray(1) Then
                    SendCommand "STAT" & vbCrLf
                    aCommand = 4
                End If
            End If
            TotalMessages = mArray(1)
            TotalSize = mArray(2)
            If mAlert = True Then
                If TotalMessages <> mArray(1) Then
                    MsgBox "You got a new Mail" & vbCrLf & vbCrLf & Mid(mStrData, InStr(1, mStrData, "From:"), InStr(InStr(1, mStrData, "From:"), mStrData, vbCr) - InStr(1, mStrData, "From:")) & vbCrLf & vbCrLf & Mid(mStrData, InStr(1, mStrData, "Subject:"), InStr(InStr(1, mStrData, "Subject:"), mStrData, vbCr) - InStr(1, mStrData, "Subject:")), vbSystemModal + vbInformation, "New Mail"
                End If
            End If
            lblStatus.Caption = "Total Message(s) on the Server is: " & TotalMessages & " Total bytes Consumed: " & TotalSize
        
        ElseIf aCommand = 4 Then
            mArray = Split(mStrData, " ")
            lstMessage = mArray(1)
            If lstMessage > 0 Then
                SendCommand "TOP " & CurMessage & " 0" & vbCrLf
                aCommand = 5
            Else
                aCommand = 0
            End If
        
        ElseIf aCommand = 5 Then
            CurMessage = CurMessage + 1
            If CurMessage > lstMessage Then
                aCommand = 0
            Else
                aCommand = 4
            End If
        
        Else
            aCommand = 0
        End If
    
    ElseIf Left(mStrData, 3) = "-ER" Then
        mErrorString = mStrData
        mErrorString = Mid(mErrorString, 6)
        TempString = UCase(Left(mErrorString, 1))
        mErrorString = TempString & Right(mErrorString, Len(mErrorString) - 1)
        lblStatus.Caption = mErrorString
        aCommand = 0
        cmdLogin.Enabled = True
        txtIncomingServer.SetFocus
        Winsock1.Close
    Else
        Set lstItem = lstMessages.ListItems.Add(, , CurMessage - 1)
        From = Mid(mStrData, InStr(1, mStrData, "From:"), InStr(InStr(1, mStrData, "From:"), mStrData, vbCr) - InStr(1, mStrData, "From:"))
        From = Mid(From, InStr(1, From, ":") + 2)
        Subject = Mid(mStrData, InStr(1, mStrData, "Subject:"), InStr(InStr(1, mStrData, "Subject:"), mStrData, vbCr) - InStr(1, mStrData, "Subject:"))
        Subject = Mid(Subject, InStr(1, Subject, ":") + 2)
        With lstItem
            .SubItems(1) = From
            .SubItems(2) = Subject
        End With
        
        If CurMessage > lstMessage Then
            SendCommand "QUIT" & vbCrLf
            Winsock1.Close
            aCommand = 0
        Else
            'send some dummy commands to go to aCommand = 4
            SendCommand "STAT" & vbCrLf
            aCommand = 4
        End If
    End If
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Description
End Sub
Private Sub SendCommand(mCommand As String)
    Winsock1.SendData mCommand
End Sub
Private Sub SaveSettings()
    If chkUsername.Value = 1 Then
        SaveSetting App.EXEName, "UserInfo", "Username", Username
    Else
        SaveSetting App.EXEName, "UserInfo", "Username", ""
    End If
    If chkPassword.Value = 1 Then
        SaveSetting App.EXEName, "UserInfo", "Password", Password
    Else
        SaveSetting App.EXEName, "UserInfo", "Password", ""
    End If
    If chkServerInfo.Value = 1 Then
        SaveSetting App.EXEName, "UserInfo", "Server", IncomingServer
        SaveSetting App.EXEName, "UserInfo", "Port", txtPort
    Else
        SaveSetting App.EXEName, "UserInfo", "Server", ""
        SaveSetting App.EXEName, "UserInfo", "Port", ""
    End If
    SaveSetting App.EXEName, "UserInfo", "SaveServer", chkServerInfo.Value
    SaveSetting App.EXEName, "UserInfo", "SaveUsername", chkUsername.Value
    SaveSetting App.EXEName, "UserInfo", "SavePassword", chkPassword.Value
End Sub
Private Sub GetSettings()
    txtUsername = GetSetting(App.EXEName, "UserInfo", "Username", "")
    txtPassword.Text = GetSetting(App.EXEName, "UserInfo", "Password", "")
    txtIncomingServer.Text = GetSetting(App.EXEName, "UserInfo", "Server", "")
    txtPort.Text = GetSetting(App.EXEName, "UserInfo", "Port", "")
    chkServerInfo.Value = GetSetting(App.EXEName, "UserInfo", "SaveServer", 0)
    chkUsername.Value = GetSetting(App.EXEName, "UserInfo", "SaveUsername", 0)
    chkPassword.Value = GetSetting(App.EXEName, "UserInfo", "SavePassword", 0)
End Sub
