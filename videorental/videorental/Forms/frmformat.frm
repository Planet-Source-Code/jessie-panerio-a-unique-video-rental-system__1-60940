VERSION 5.00
Begin VB.Form frmformat 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Format"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   Icon            =   "frmformat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmformat.frx":08CA
   ScaleHeight     =   3795
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      Left            =   2880
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtFormat 
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCCancel 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCAdd 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCSave 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCCLose 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCDelete 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FF8080&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      Picture         =   "frmformat.frx":2552
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
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
      Height          =   855
      Left            =   3960
      Picture         =   "frmformat.frx":2E1C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FF8080&
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmformat.frx":325E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   855
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
      Height          =   855
      Left            =   1080
      Picture         =   "frmformat.frx":3B28
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   855
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
      Height          =   855
      Left            =   2040
      Picture         =   "frmformat.frx":43F2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblExistFormat 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "List of Existing &Format"
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
      Left            =   600
      TabIndex        =   15
      Top             =   1080
      Width           =   2115
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter &New Format"
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
      Left            =   600
      TabIndex        =   14
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Label lblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Price"
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
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modeval As Boolean

Private Sub Form_Load()
  Call connectcombo
  Call fillcombo
  cmdButton "0 1 1 0 0 1 0 0 1 1 1 0 0 1 1"
  lblPrice.Caption = "&Price (Rentals)"
End Sub

Private Sub cmdButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCAdd.Visible = strAry(0)
  cmdCCancel.Visible = strAry(1)
  cmdCSave.Visible = strAry(2)
  cmdCDelete.Visible = strAry(3)
  cmdCClose.Visible = strAry(4)
  cmdAdd.Enabled = strAry(5)
  cmdCancel.Enabled = strAry(6)
  cmdsave.Enabled = strAry(7)
  cmdDelete.Enabled = strAry(8)
  cmdClose.Enabled = strAry(9)
  Combo1.Enabled = strAry(10)
  lblFormat.Visible = strAry(11)
  txtFormat.Visible = strAry(12)
  txtPrice.Locked = strAry(13)
  txtFormat.Locked = strAry(14)
End Sub

Private Sub cmdAdd_Click()
  cmdButton "1 0 0 1 1 0 1 1 0 0 0 1 1 0 0"
  lblPrice.Caption = "Enter &Price (Rentals)"
  Combo1.Text = ""
  txtFormat.Text = ""
  txtPrice.Text = ""
  txtPrice.SetFocus
End Sub

Private Sub cmdCancel_Click()
  adopanerio.CancelUpdate
  adopanerio.Requery
  cmdButton "0 1 1 0 0 1 0 0 1 1 1 0 0 1 1"
  lblPrice.Caption = "&Price (Rentals)"
  txtPrice.Text = ""
  txtFormat.Text = ""
  Combo1.SetFocus
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  Load frmItemList
  frmItemList.Show
  frmItemList.Enabled = True
  Call clear_opt_txtbox_itemlist
End Sub

Private Sub cmdDelete_Click()
  Dim res As VbMsgBoxResult
  On Error Resume Next
    If adopanerio.BOF = True And adopanerio.EOF = True Then
      Exit Sub
      MsgBox "Empty Database", vbCritical, "Format"
    ElseIf Combo1.Text = "" Then
      MsgBox "Select a Format", vbInformation, "Format Configuration"
      Combo1.SetFocus
      Exit Sub
    Else
      res = MsgBox("Are you sure you want to Delete " & adopanerio!itemformat & "?", vbYesNo + vbQuestion, "Confirmation")
        If res = vbYes Then
          adopanerio.Delete
          adopanerio.Requery
          Combo1.clear
          txtPrice.Text = ""
          Call fillcombo
          Combo1.SetFocus
        Else
          Exit Sub
        End If
    End If
End Sub

Private Sub cmdSave_Click()
  On Error Resume Next
    With frmformat
      If .txtFormat.Text = "" Or .txtPrice.Text = "" Then
        MsgBox "Missing Data! Do not leave a blank textfield", vbInformation, "Format Configuration"
        txtFormat.SetFocus
        Exit Sub
      Else
        Call validate
          If modeval = False Then
            Dim res As VbMsgBoxResult
            res = MsgBox("Save " & txtFormat.Text & " format?", vbYesNo + vbQuestion, "Confirmation")
              If res = vbYes Then
                txtPrice.Text = Format(txtPrice.Text, "###.00")
                adopanerio.AddNew
                adopanerio!itemformat = txtFormat.Text
                adopanerio!price = txtPrice.Text
                adopanerio.Update
                Call fillcombo
                cmdButton "0 1 1 0 0 1 0 0 1 1 1 0 0 1 1"
                lblPrice.Caption = "&Price (Rentals)"
                txtPrice.Text = ""
                Combo1.SetFocus
              Else
                Exit Sub
              End If
          Else
            MsgBox "Warning! Duplication of entries is not allowed in this application." & vbCrLf & vbCrLf & "                    ''" & txtFormat.Text & "'' format already exist.", vbExclamation, "Format Configuration"
            txtFormat.SetFocus
            SendKeys hl
          End If
      End If
    End With
End Sub

Private Sub validate()
  Dim rs As New ADODB.Recordset
  rs.Open "Select * From format Where itemformat = '" & txtFormat.Text & "'", cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
      modeval = False
      Exit Sub
    Else
      modeval = True
    End If
      Set rs = Nothing
End Sub

Private Sub Combo1_Click()
  On Error Resume Next
    If adopanerio.BOF = True And adopanerio.EOF = True Then
      Exit Sub
    Else
      adopanerio.MoveFirst
      adopanerio.Move Combo1.ListIndex
      txtPrice.Text = adopanerio!price
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
End Sub

Private Sub txtFormat_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8) Then
    KeyAscii = 0
    End If
End Sub
