VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmeditmembership 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Members Record"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmeditmembership.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1320
      Picture         =   "frmeditmembership.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5640
      Width           =   975
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
      Left            =   240
      Picture         =   "frmeditmembership.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   975
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
      Left            =   2400
      Picture         =   "frmeditmembership.frx":1E5E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   758
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "  Membership Form"
      TabPicture(0)   =   "frmeditmembership.frx":22A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C000&
         Height          =   4875
         Left            =   0
         ScaleHeight     =   4815
         ScaleWidth      =   6555
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdAddPicture 
            BackColor       =   &H00FF8080&
            Caption         =   "Change &Picture"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2520
            Width           =   1695
         End
         Begin VB.PictureBox picimage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   4680
            ScaleHeight     =   1425
            ScaleWidth      =   1740
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   960
            Width           =   1770
            Begin MSComDlg.CommonDialog cd2 
               Left            =   1920
               Top             =   840
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSComDlg.CommonDialog cd1 
               Left            =   1920
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Image imgpic 
               BorderStyle     =   1  'Fixed Single
               Height          =   1395
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1755
            End
         End
         Begin VB.TextBox txtFName 
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
            Height          =   390
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtLName 
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
            Height          =   390
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   1
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtMID 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtMName 
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
            Height          =   390
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   5
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtBDate 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   735
            Left            =   2040
            MaxLength       =   92
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   3480
            Width           =   4455
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
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
            Left            =   5040
            MaxLength       =   8
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtLandLine 
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
            Left            =   2040
            MaxLength       =   16
            TabIndex        =   13
            Top             =   4320
            Width           =   1815
         End
         Begin VB.TextBox txtMobile 
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
            Left            =   4680
            MaxLength       =   16
            TabIndex        =   15
            Top             =   4320
            Width           =   1815
         End
         Begin VB.ComboBox cmbGender 
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
            Height          =   390
            ItemData        =   "frmeditmembership.frx":2B7A
            Left            =   2040
            List            =   "frmeditmembership.frx":2B84
            TabIndex        =   9
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txtPictureName 
            Height          =   285
            Left            =   5040
            TabIndex        =   29
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "MM/DD/YYYY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3600
            TabIndex        =   31
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Format"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "&FirstName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "&LastName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   0
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Borrowers ID#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C000&
            Caption         =   "&MiddleName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
            Caption         =   "&BirthDate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   6
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C000&
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
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C000&
            Caption         =   "Mem. &Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   5160
            TabIndex        =   16
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C000&
            Caption         =   "Contact #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   4080
            Width           =   975
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C000&
            Caption         =   "LandL&ine"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   4440
            Width           =   975
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C000&
            Caption         =   "M&obile"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3960
            TabIndex        =   14
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            Caption         =   "&Gender"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   3120
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmeditmembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbGender_KeyPress(KeyAscii As Integer)
  Dim strvalid
    strvalid = ""
    If KeyAscii > 26 Then
      If InStr(strvalid, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
    End If
End Sub

Private Sub cmdAddPicture_Click()
  Call frmaddmembership.cmdAddPicture_Click
  cmdAddPicture.Visible = False
End Sub

Private Sub cmdCancel_Click()
  Call edit_display_data
  cmdAddPicture.Visible = True
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Load frmMembership
  frmMembership.Show
  Call clear_opt_txtbox_members
End Sub

Private Sub cmdSave_Click()
  Dim resp As VbMsgBoxResult
    With frmeditmembership
      If .txtLName.Text = "" Or .txtFName.Text = "" Or .txtMName.Text = "" Or .txtBDate.Text = "" Or .cmbGender.Text = "" Or .txtAddress.Text = "" Then
         MsgBox "Missing Data! Do not leave a blank textfield.", vbInformation, "Information"
         Exit Sub
      ElseIf .txtLandLine.Text = "" And .txtMobile.Text = "" Then
         resp = MsgBox("Member have no Contact No., Do you want to proceed?", vbYesNo, "Information")
         If resp = vbYes Then
           GoTo continue:
           Else
           Exit Sub
         End If
      End If
    End With
continue:
  Dim res As VbMsgBoxResult
  res = MsgBox("Save this to Database?", vbYesNo, "Confirmation")
    If res = vbYes Then
      txtBDate.Text = Format(txtBDate.Text, "MM/DD/YYYY")
      Call edit_save_data
      Call cmdClose_Click
    Else
      Exit Sub
    End If
End Sub

Private Sub Form_Activate()
  Call edit_display_data
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call cmdClose_Click
End Sub

Private Sub txtBDate_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtLandLine_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 43 Or KeyAscii = 45) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtLName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtMName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 43 Or KeyAscii = 45) Then
    KeyAscii = 0
  End If
End Sub
