VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Penalty Price Rate"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmRate.frx":0442
   ScaleHeight     =   2970
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCCancel 
      BackColor       =   &H00FF8080&
      Height          =   760
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCSave 
      BackColor       =   &H00FF8080&
      Height          =   760
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCClose 
      BackColor       =   &H00FF8080&
      Height          =   760
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   760
      Left            =   1800
      Picture         =   "frmRate.frx":1E2E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FF8080&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   760
      Left            =   2640
      Picture         =   "frmRate.frx":2AF8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "Canc&el"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   760
      Left            =   960
      Picture         =   "frmRate.frx":2F3A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCChange 
      BackColor       =   &H00FF8080&
      Height          =   760
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Width           =   755
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FF8080&
      Caption         =   "C&hange"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   760
      Left            =   120
      Picture         =   "frmRate.frx":3804
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   755
   End
   Begin VB.TextBox txtNewRate 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1320
      Width           =   1020
   End
   Begin VB.TextBox txtPRate 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Penalty Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblNewRate 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New &Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCChange.Visible = strAry(0)
  cmdCCancel.Visible = strAry(1)
  cmdCSave.Visible = strAry(2)
  cmdCClose.Visible = strAry(3)
  cmdChange.Enabled = strAry(4)
  cmdCancel.Enabled = strAry(5)
  cmdsave.Enabled = strAry(6)
  cmdClose.Enabled = strAry(7)
  lblNewRate.Visible = strAry(8)
  txtNewRate.Visible = strAry(9)
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  Load frmItemList
  frmItemList.Show
  frmItemList.Enabled = True
  Call clear_opt_txtbox_itemlist
End Sub

Private Sub Form_Load()
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open "Select * from penaltyrateperday", cnn, adOpenStatic, adLockPessimistic
  txtPRate.Text = adopanerio!penaltyrateperday
  cmdButton "0 1 1 0 1 0 0 1 0 0"
End Sub
 
Private Sub cmdChange_Click()
  cmdButton "1 0 0 1 0 1 1 0 1 1"
  txtNewRate.SetFocus
  txtNewRate.Text = ""
End Sub

Private Sub cmdCancel_Click()
  txtPRate.Text = adopanerio!penaltyrateperday
  cmdButton "0 1 1 0 1 0 0 1 0 0"
End Sub

Private Sub cmdSave_Click()
  If txtNewRate.Text = "" Then
    MsgBox "Data Missing! Enter New Penalty Rate in the textfield", vbInformation, "Penalty Rate"
    txtNewRate.SetFocus
    Exit Sub
  Else
    Dim res As VbMsgBoxResult
    txtNewRate.Text = Format(txtNewRate.Text, "###.00")
    res = MsgBox("Save " & txtNewRate.Text & " as new Penalty Rate?", vbYesNo + vbQuestion, "Confirmation")
      If res = vbYes Then
        cmdButton "0 1 1 0 1 0 0 1 0 0"
        adopanerio!penaltyrateperday = txtNewRate.Text
        adopanerio.UpdateBatch adAffectCurrent
        txtPRate.Text = adopanerio!penaltyrateperday
      Else
        txtNewRate.SetFocus
        SendKeys hl
      End If
  End If
End Sub

Private Sub txtNewRate_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
    KeyAscii = 0
  End If
End Sub
