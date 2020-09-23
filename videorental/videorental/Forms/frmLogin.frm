VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2325
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1620
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   720
      Width           =   1620
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblLogUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   945
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub login_search()
  Dim rs As New ADODB.Recordset
  rs.Open "Select * From users Where UserName = '" & txtUserName.Text & "'", cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
      MsgBox "User Name is invalid.", vbInformation, "Login"
      txtUserName.SetFocus
      Exit Sub
    Else
      If txtPassword.Text = rs!Password Then
        Unload Me
        Load frmMain
        frmMain.Show
        frmMain.Enabled = True
        frmMain.mnuLogOff.Caption = "&Log Off " & rs!UserName & "..."
        frmMain.StatusBar1.Panels(2).Text = rs!Level
        frmMain.StatusBar1.Panels(5).Text = rs!UserName
        frmMain.StatusBar1.Panels(8).Text = Time
        frmMain.StatusBar1.Panels(11).Text = Date
        Call User_LogIn
        Exit Sub
      Else
        MsgBox "Password is invalid.", vbInformation, "Login"
        txtPassword.SetFocus
        Exit Sub
      End If
    End If
    Set rs = Nothing
End Sub

Private Sub cmdEnter_Click()
  If txtUserName.Text = "" Then
    txtUserName.SetFocus
    Exit Sub
  ElseIf txtPassword.Text = "" Then
    txtPassword.SetFocus
    Exit Sub
  Else
    Call login_search
  End If
End Sub

Private Sub cmdExit_Click()
  If MsgBox("This will terminate the application. Proceed?", vbOKCancel + vbQuestion, "Video Rental System") = vbOK Then
    cnn.Close
    End
  Else
    txtUserName.SetFocus
    Exit Sub
  End If
End Sub

Private Sub txtUserName_GotFocus()
  SendKeys hl
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
    txtPassword.SetFocus
    SendKeys hl
  End If
End Sub

Private Sub txtPassword_GotFocus()
  SendKeys hl
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
    Call cmdEnter_Click
  End If
End Sub


