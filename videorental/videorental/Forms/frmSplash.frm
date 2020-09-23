VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   720
      Top             =   4920
   End
   Begin VB.Image Image1 
      Height          =   4245
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call getconnected
    Dim rs As New ADODB.Recordset
    rs.Open "Select * From users", cnn, adOpenStatic, adLockReadOnly
      If rs.RecordCount < 1 Then
        MsgBox "Warning! No User Exist, it is a must" & vbCrLf & "that you established a User Account", vbInformation, "Information"
        LogIn = True
        Load frmUserConfig
        frmUserConfig.Show
      Else
        Load frmLogin
        frmLogin.Show
      End If
      Set rs = Nothing
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub
