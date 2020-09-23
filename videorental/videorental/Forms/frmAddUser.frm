VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   758
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmAddUser.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdsave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdlcancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdlcancel 
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
         Left            =   1200
         Picture         =   "frmAddUser.frx":171C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
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
         Left            =   120
         Picture         =   "frmAddUser.frx":1FE6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C000&
         Height          =   2700
         Left            =   0
         ScaleHeight     =   2640
         ScaleWidth      =   2595
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
         Begin VB.TextBox txtUserName 
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   3
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtVal 
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtCPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   6
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbLevel 
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
            ForeColor       =   &H00000000&
            Height          =   360
            ItemData        =   "frmAddUser.frx":2CB0
            Left            =   840
            List            =   "frmAddUser.frx":2CBA
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C000&
            Caption         =   "&Level"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C000&
            Caption         =   "&UserName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C000&
            Caption         =   "&Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C000&
            Caption         =   "&Confirm Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modeval As Boolean

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
  Dim strvalid
    strvalid = ""
      If KeyAscii > 26 Then
        If InStr(strvalid, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
        End If
      End If
End Sub

Private Sub cmdlCancel_Click()
  frmUserConfig.Enabled = True
  Unload Me
End Sub

Private Sub validate_user()
  Dim rs As New ADODB.Recordset
  rs.Open "Select * From users Where UserName = '" & txtUserName.Text & "'", cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
      modeval = False
      Exit Sub
    Else
      modeval = True
    End If
      Set rs = Nothing
End Sub

Private Sub cmdSave_Click()
  Dim resp As VbMsgBoxResult
    If cmbLevel.Text = "" Or txtUserName.Text = "" Or txtPassword.Text = "" Or txtCPassword.Text = "" Then
      MsgBox "Missing Data! Do not leave a blank textfield.", vbInformation, "Information"
      Exit Sub
    Else
      If txtPassword.Text <> txtCPassword.Text Then
        MsgBox "Password does not match", vbInformation, "Password"
        Exit Sub
      Else
        Call validate_user
          If modeval = False Then
save:
            Dim res As VbMsgBoxResult
            res = MsgBox("Save this to Database?", vbYesNo + vbQuestion, "Confirmation")
              If res = vbYes Then
                If useradd = True And useredit = False Then
                  adopanerio.AddNew
edit:
                  adopanerio!Level = cmbLevel.Text
                  adopanerio!UserName = txtUserName.Text
                  adopanerio!Password = txtPassword.Text
                  adopanerio!CPassword = txtCPassword.Text
                    If useradd = True And useredit = False Then
                      adopanerio.Update
                      adopanerio.Requery
                    ElseIf useredit = True And useradd = False Then
                      adopanerio.UpdateBatch adAffectCurrent
                    End If
                      frmUserConfig.Enabled = True
                      Load frmUserConfig
                      frmUserConfig.Show
                      Unload Me
                      Call User_recno
                ElseIf useredit = True And useradd = False Then
                  GoTo edit:
                End If
              Else
                Exit Sub
              End If
          ElseIf modeval = True Then
            If txtVal.Text = txtUserName.Text Then
              GoTo save:
            Else
              MsgBox "Warning! Duplication of entries is not allowed in this application." & vbCrLf & vbCrLf & "UserName  " & "''" & txtUserName.Text & "''  already exist.", vbExclamation, "Users Configuration"
              txtUserName.SetFocus
              SendKeys hl
            End If
            
          End If
      End If
    End If
End Sub

Private Sub Form_Load()
  If useredit = True And useradd = False Then
    cmbLevel.Text = adopanerio!Level
    txtUserName.Text = adopanerio!UserName
    txtPassword.Text = adopanerio!Password
    txtCPassword.Text = adopanerio!CPassword
    txtVal.Text = adopanerio!UserName
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmUserConfig.Enabled = True
  Unload Me
End Sub

Private Sub txtCPassword_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
