VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmLocked 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter &Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1380
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   210
      Left            =   4440
      TabIndex        =   3
      Top             =   4680
      Width           =   180
      ExtentX         =   317
      ExtentY         =   370
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  On Error Resume Next
  Kill App.Path & "\tmpfile.swf"
  LoadDataIntoFile 102, App.Path & "\tmpfile.swf"
  WebBrowser1.Navigate App.Path & "\tmpfile.swf"
End Sub

Private Sub Form_Resize()
  WebBrowser1.Left = Me.ScaleLeft
  WebBrowser1.Top = Me.ScaleTop
  WebBrowser1.Width = Me.ScaleWidth
  WebBrowser1.Height = Me.ScaleHeight
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
    Dim rs As New ADODB.Recordset
      rs.Open "Select * From users Where UserName = '" & frmMain.StatusBar1.Panels(5).Text & "'", cnn, adOpenStatic, adLockReadOnly
        If txtPassword.Text = rs!Password Then
          Unload Me
          Kill App.Path & "\tmpfile.swf"
          Load frmMain
          frmMain.Show
          frmMain.Enabled = True
        Else
          MsgBox "Password is invalid.", vbCritical, "Application Locked"
          txtPassword.Text = ""
          txtPassword.SetFocus
          Exit Sub
        End If
  End If
  Set rs = Nothing
End Sub
