VERSION 5.00
Begin VB.Form frmCashier 
   BorderStyle     =   0  'None
   Caption         =   "Cashier"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCashier.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   1020
   End
   Begin VB.TextBox txtCash 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   1020
   End
   Begin VB.TextBox txtChange 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Compute"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Total Amount"
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
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Ca&sh"
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
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Change"
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
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompute_Click()
  If txtCash.Text = "" Then
    MsgBox "Enter Cash Amount", vbInformation, "Cashier"
    txtCash.SetFocus
  Else
    txtCash.Text = Format(txtCash.Text, "####.00")
    txtChange.Text = Val(txtCash.Text) - Val(txtTotalAmount.Text)
    txtChange.Text = Format(txtChange.Text, "####.00")
      If Val(txtChange.Text) < 0 Then
        txtChange.Text = ""
        MsgBox "Warning! Insufficient Cash Amount", vbCritical, "Cashier"
        txtCash.SetFocus
        SendKeys hl
        Exit Sub
      Else
        If Val(txtChange.Text) = 0 Then
          txtChange.Text = 0
        End If
      End If
  End If
 End Sub

Private Sub cmdExit_Click()
  Dim a
  If txtCash.Text = "" Then
    MsgBox "Enter Cash Amount then click the Compute button", vbInformation, "Cashier"
    txtCash.SetFocus
  Else
    If txtChange.Text = "" Then
      MsgBox "Click first the Compute button", vbInformation, "Cashier"
    Else
      a = Val(txtCash.Text) - Val(txtTotalAmount.Text)
        If Val(a) = Val(txtChange) Then
          Dim res As VbMsgBoxResult
          res = MsgBox("Change is:  " & txtChange.Text & vbCrLf & vbCrLf & "Do you want to Print the Official Receipt?", vbOKCancel + vbQuestion, "Cashier")
            If res = vbOK Then
              frmRent.Picture1.Enabled = False
              frmRent.Picture3.Enabled = False
              frmRent.Enabled = True
              Call setup_connected
                With DataReport1.Sections("Section2").Controls
                  .Item("lblName").Caption = adopanerio!nname
                  .Item("lblAddr").Caption = adopanerio!address
                  .Item("lblCashier").Caption = "Cashier " & frmMain.StatusBar1.Panels(5).Text
                End With
              Set adopanerio = Nothing
              print_or (frmRent.Text6.Text)
              Call clearall
              Unload Me
            Else
              Exit Sub
            End If
        Else
          MsgBox "Warning! Change is miscalculated." & vbCrLf & vbCrLf & "Click the Compute button to update computation", vbInformation, "Cashier"
          Exit Sub
        End If
    End If
  End If
End Sub

Private Sub Form_Load()
  txtTotalAmount.Text = frmRent.txtAmount.Text
End Sub

Private Sub txtCash_GotFocus()
  SendKeys hl
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then
    KeyAscii = 0
  ElseIf KeyAscii = 13 Then
    cmdCompute_Click
  End If
End Sub
