VERSION 5.00
Begin VB.Form frmVideoRentalSetup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmVideoRentalSetup.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FF8080&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtaddr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox txtcontactnum 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   4575
   End
   Begin VB.TextBox txtemail 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      MaxLength       =   34
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-&mail Address"
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
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Con&tact #"
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
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Address"
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
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Business &Name"
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
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1440
   End
End
Attribute VB_Name = "frmVideoRentalSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  frmMain.Show
  frmMain.Enabled = True
End Sub

Private Sub cmdSave_Click()
  On Error Resume Next
  If txtname.Text = "" And txtaddr.Text = "" Then
    GoTo nxt:
  ElseIf txtname.Text = "" And txtaddr.Text <> "" Then
    txtname.Text = "No Name Video Rental"
    GoTo nx:
  ElseIf txtname.Text <> "" And txtaddr.Text = "" Then
    txtaddr.Text = "No Address Specified"
    GoTo nx:
nxt:
    txtname.Text = "No Name Video Rental"
    txtaddr.Text = "No Address Specified"
nx:
    adopanerio!nname = txtname.Text
    adopanerio!address = txtaddr.Text
    adopanerio!email = txtemail.Text
    adopanerio!contactnum = txtcontactnum.Text
    adopanerio.Update
    MsgBox "Changes has been successfully saved!", vbInformation, "Setup"
  Else
    GoTo nx:
  End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Call setup_connected
  txtname.Text = adopanerio!nname
  txtaddr.Text = adopanerio!address
  txtemail.Text = adopanerio!email
  txtcontactnum.Text = adopanerio!contactnum
End Sub

Private Sub txtaddr_GotFocus()
  SendKeys hl
End Sub

Private Sub txtaddr_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtemail.SetFocus
  End If
End Sub

Private Sub txtcontactnum_GotFocus()
  SendKeys hl
End Sub

Private Sub txtcontactnum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdsave.SetFocus
  End If
End Sub

Private Sub txtemail_GotFocus()
  SendKeys hl
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtcontactnum.SetFocus
  End If
End Sub

Private Sub txtname_GotFocus()
  SendKeys hl
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtaddr.SetFocus
  End If
End Sub
