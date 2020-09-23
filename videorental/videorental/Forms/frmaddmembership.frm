VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmaddmembership 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Entry"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmaddmembership.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCClose 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCCancel 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCSave 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCAdd 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FF8080&
      Caption         =   "&New"
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
      Picture         =   "frmaddmembership.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
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
      Left            =   3480
      Picture         =   "frmaddmembership.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   2400
      Picture         =   "frmaddmembership.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   975
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
      Left            =   1320
      Picture         =   "frmaddmembership.frx":22A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   23
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
      TabPicture(0)   =   "frmaddmembership.frx":2B6A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C000&
         Height          =   4935
         Left            =   0
         ScaleHeight     =   4875
         ScaleWidth      =   6555
         TabIndex        =   24
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdCRemovePicture 
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
            Height          =   375
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3000
            Width           =   1725
         End
         Begin VB.CommandButton cmdCAddPicture 
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
            Height          =   375
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   37
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
            TabIndex        =   35
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
         Begin VB.TextBox dateMem 
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
            Height          =   390
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemovePicture 
            BackColor       =   &H00FF8080&
            Caption         =   "&Remove Picture"
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
            TabIndex        =   21
            Top             =   3000
            Width           =   1725
         End
         Begin VB.CommandButton cmdAddPicture 
            BackColor       =   &H00FF8080&
            Caption         =   "Add &Picture"
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
            TabIndex        =   20
            Top             =   2520
            Width           =   1695
         End
         Begin VB.ComboBox cmbGender 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            ItemData        =   "frmaddmembership.frx":3444
            Left            =   2040
            List            =   "frmaddmembership.frx":344E
            TabIndex        =   13
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txtMobile 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   19
            Top             =   4320
            Width           =   1815
         End
         Begin VB.TextBox txtLandLine 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   17
            Top             =   4320
            Width           =   1815
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   615
            Left            =   2040
            MaxLength       =   92
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   3480
            Width           =   4455
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
            TabIndex        =   11
            Top             =   2520
            Width           =   1455
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
            TabIndex        =   9
            Top             =   2040
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
            TabIndex        =   26
            Top             =   360
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
            TabIndex        =   5
            Top             =   1080
            Width           =   2415
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
            TabIndex        =   7
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtautonum 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtPictureName 
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1800
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   2520
            Width           =   615
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
            TabIndex        =   12
            Top             =   3120
            Width           =   855
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
            TabIndex        =   18
            Top             =   4440
            Width           =   735
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
            TabIndex        =   16
            Top             =   4440
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
            TabIndex        =   29
            Top             =   4080
            Width           =   975
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C000&
            Caption         =   "Date"
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
            Left            =   4800
            TabIndex        =   28
            Top             =   480
            Width           =   495
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
            TabIndex        =   14
            Top             =   3600
            Width           =   855
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
            TabIndex        =   10
            Top             =   2640
            Width           =   975
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
            TabIndex        =   8
            Top             =   2160
            Width           =   1335
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
            TabIndex        =   27
            Top             =   480
            Width           =   1695
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
            TabIndex        =   4
            Top             =   1200
            Width           =   975
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
            TabIndex        =   6
            Top             =   1680
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmaddmembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCAdd.Visible = strAry(0)
  cmdCSave.Visible = strAry(1)
  cmdCCancel.Visible = strAry(2)
  cmdCClose.Visible = strAry(3)
  cmdAdd.Enabled = strAry(4)
  cmdsave.Enabled = strAry(5)
  cmdCancel.Enabled = strAry(6)
  cmdClose.Enabled = strAry(7)
  cmdCAddPicture.Visible = strAry(8)
  cmdCRemovePicture.Visible = strAry(9)
  cmdAddPicture.Enabled = strAry(10)
  cmdRemovePicture.Enabled = strAry(11)
End Sub

Private Sub cmdAdd_Click()
  Call clear_txtbox
  cmdRemovePicture_Click
  Call rs_act_autonum
  With adoautonum
     frmaddmembership.txtautonum.Text = !autonum
  End With
  txtautonum.Text = Val(txtautonum.Text) + 1
  txtautonum.Text = Format(txtautonum.Text, "0000000")
  txtMID.Text = txtautonum.Text
  cmdCButton "1 0 0 1 0 1 1 0 0 0 1 1"
  txtLName.SetFocus
  Call lock_textbox(False)
End Sub

Public Sub cmdAddPicture_Click()
  On Error Resume Next
     With cd1
       .InitDir = "C:\My Documents"
       .Filter = "JPEG image|*.jpg|GIF image|*.gif|BITMAP image|*.bmp|Icon image|*.ico|Cursor image|*.cur|Panerio image|*.pan"
       .ShowOpen
          If frmaddmembership.Visible = True And frmeditmembership.Visible = False Then
             If .FileName <> "" Then
                strImgN = .FileName
                txtPictureName.Text = .FileTitle
                imgpic.Picture = LoadPicture(.FileName)
             End If
          ElseIf frmaddmembership.Visible = False And frmeditmembership.Visible = True Then
             If .FileName <> "" Then
                strImgN = .FileName
                frmeditmembership.txtPictureName.Text = .FileTitle
                frmeditmembership.imgpic.Picture = LoadPicture(.FileName)
             End If
          End If
     End With
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Load frmMembership
  frmMembership.Show
  Call clear_opt_txtbox_members
  Call LoadImage
  Call recno
End Sub

Private Sub cmdRemovePicture_Click()
  Set imgpic.Picture = Nothing
  txtPictureName.Text = ""
End Sub

Private Sub cmdSave_Click()
  Dim resp As VbMsgBoxResult
     With frmaddmembership
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
        adomembership.AddNew
        Call WriteDataFromControls
        adomembership.Update
        adoautonum.UpdateBatch adAffectCurrent
        cmdCButton "0 1 1 0 1 0 0 1 0 0 1 1"
        MsgBox "Records successfully Saved to Database", vbInformation, "Membership"
        Call lock_textbox(True)
     Else
        Exit Sub
     End If
End Sub

Private Sub cmdCancel_Click()
  cmdCButton "0 1 1 0 1 0 0 1 1 1 0 0"
  txtautonum.Text = Val(txtautonum.Text) - 1
  txtautonum.Text = Format(txtautonum.Text, "0000000")
  Call clear_txtbox
  Call cmdRemovePicture_Click
  Call lock_textbox(True)
End Sub

Private Sub clear_txtbox()
  txtMID.Text = ""
  txtLName.Text = ""
  txtFName.Text = ""
  txtMName.Text = ""
  txtBDate.Text = ""
  cmbGender.Text = ""
  txtAddress.Text = ""
  txtLandLine.Text = ""
  txtMobile.Text = ""
End Sub

Private Sub Form_Load()
  cmdCButton "0 1 1 0 1 0 0 1 1 1 0 0"
  lock_textbox (True)
  dateMem.Text = Date
End Sub

Private Sub lock_textbox(answer As String)
  txtLName.Locked = answer
  txtFName.Locked = answer
  txtMName.Locked = answer
  txtBDate.Locked = answer
  cmbGender.Locked = answer
  txtAddress.Locked = answer
  txtLandLine.Locked = answer
  txtMobile.Locked = answer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdCAdd.Visible = True Then
     adomembership.CancelUpdate
     MsgBox "Add New Member Cancelled", vbInformation, "Membership"
     Call cmdClose_Click
  Else
     Call cmdClose_Click
  End If
End Sub

Private Sub cmbGender_KeyPress(KeyAscii As Integer)
  Dim strvalid
  strvalid = ""
    If KeyAscii > 26 Then
      If InStr(strvalid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If
    End If
End Sub

Private Sub txtBDate_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtLName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtMName_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
Private Sub txtLandLine_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 43 Or KeyAscii = 45) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 43 Or KeyAscii = 45) Then
     KeyAscii = 0
  End If
End Sub
