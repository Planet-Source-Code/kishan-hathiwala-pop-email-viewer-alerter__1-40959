VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4770
      TabIndex        =   5
      Top             =   585
      Width           =   1005
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4770
      TabIndex        =   4
      Top             =   135
      Width           =   1005
   End
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   2925
      TabIndex        =   2
      Text            =   "5"
      Top             =   585
      Width           =   465
   End
   Begin VB.CheckBox chkAlert 
      Caption         =   "Alert Me when new Message arrives at my Inbox"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4560
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes"
      Height          =   240
      Left            =   3465
      TabIndex        =   3
      Top             =   630
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Check for New messages every"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   2805
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
    If Len(Trim(txtMinutes.Text)) <= 0 Then
        MsgBox "Error in Minutes!!", vbInformation, "Error"
        txtMinutes.Text = 5
        txtMinutes.SetFocus
        Exit Sub
    Else
        If Val(txtMinutes.Text) <= 4 Then
            MsgBox "Cannot be Less than 4 Minutes. Plese Enter the value between 5 to 15", vbInformation, "Error"
            txtMinutes.Text = 5
            txtMinutes.SetFocus
            Exit Sub
        Else
            If Val(txtMinutes.Text) > 15 Then
                MsgBox "Cannot be Greater than 15 Minutes. Plese Enter the value between 5 to 15", vbInformation, "Error"
                txtMinutes.Text = 5
                txtMinutes.SetFocus
                Exit Sub
            End If
        End If
    End If
    mInterval = Val(txtMinutes.Text) * 60000
    mAlert = CBool(chkAlert.Value)
    SetTimer frmPOP.hwnd, 0, mInterval, AddressOf TimerProc
    cmdClose_Click
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    If mAlert = True Then
        chkAlert.Value = 1
    Else
        chkAlert.Value = 0
    End If
End Sub
Private Sub txtMinutes_GotFocus()
    If Len(txtMinutes.Text) > 0 Then
        txtMinutes.SelStart = 0
        txtMinutes.SelLength = Len(txtMinutes.Text)
    End If
End Sub
