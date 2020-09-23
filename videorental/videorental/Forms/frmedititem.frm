VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmedititem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Item"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmedititem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10398
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
      TabCaption(0)   =   " Item "
      TabPicture(0)   =   "frmedititem.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C000&
         Height          =   5415
         Left            =   0
         ScaleHeight     =   5355
         ScaleWidth      =   5715
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox cmblformat 
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
            Height          =   390
            ItemData        =   "frmedititem.frx":15A4
            Left            =   1800
            List            =   "frmedititem.frx":15A6
            TabIndex        =   6
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtFormatRentPrice 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtlstatus 
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
            Height          =   405
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   4680
            Width           =   615
         End
         Begin VB.TextBox txtlnoofcd 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   20
            Top             =   4680
            Width           =   495
         End
         Begin VB.TextBox txtldate 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   16
            Top             =   3720
            Width           =   1095
         End
         Begin VB.TextBox txtlnoofdays 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   10
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtlprice 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   18
            Top             =   4200
            Width           =   1095
         End
         Begin VB.TextBox txtlsecondcast 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   14
            Top             =   3240
            Width           =   2415
         End
         Begin VB.TextBox txtlmaincast 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   12
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtlcategory 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtltitle 
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
            Height          =   375
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   4
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox txtlitemid 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   3120
            TabIndex        =   29
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " MM/DD/YY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   3000
            TabIndex        =   28
            Top             =   3960
            Width           =   705
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Status"
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
            Left            =   2640
            TabIndex        =   21
            Top             =   4800
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "N&o. of Disc/VHS"
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
            Height          =   240
            Left            =   240
            TabIndex        =   19
            Top             =   4800
            Width           =   1470
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C000&
            Caption         =   "Date Purc&hased"
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
            Left            =   240
            TabIndex        =   15
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C000&
            Caption         =   "No. of &Days"
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
            Left            =   240
            TabIndex        =   9
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Item &Price"
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
            Height          =   240
            Left            =   240
            TabIndex        =   17
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C000&
            Caption         =   "Second C&ast"
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
            Left            =   240
            TabIndex        =   13
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C000&
            Caption         =   "&Main Cast"
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
            Left            =   240
            TabIndex        =   11
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C000&
            Caption         =   "Catego&ry"
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
            Left            =   240
            TabIndex        =   7
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C000&
            Caption         =   "&Title"
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
            Left            =   240
            TabIndex        =   3
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C000&
            Caption         =   "&Format"
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
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C000&
            Caption         =   "Item ID#:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdlsave 
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
      Picture         =   "frmedititem.frx":15A8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   975
   End
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
      Picture         =   "frmedititem.frx":2272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
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
      Left            =   2280
      Picture         =   "frmedititem.frx":2B3C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5880
      Width           =   5775
   End
End
Attribute VB_Name = "frmedititem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmblformat_Click()
  On Error Resume Next
  adopanerio.MoveFirst
  adopanerio.Move cmblformat.ListIndex
  txtFormatRentPrice.Text = adopanerio!price
End Sub

Private Sub cmblformat_KeyPress(KeyAscii As Integer)
  Dim strvalid
  strvalid = ""
    If KeyAscii > 26 Then
      If InStr(strvalid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If
    End If
End Sub

Private Sub filformcombo()
    Dim z
    For z = 1 To adopanerio.RecordCount
      frmedititem.cmblformat.AddItem adopanerio.Fields(0)
      adopanerio.MoveNext
    Next z
End Sub

Private Sub cmdlCancel_Click()
  Call edit_display_item_data
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  Load frmItemList
  frmItemList.Show
  frmItemList.Enabled = True
  Call clear_opt_txtbox_itemlist
End Sub

Private Sub cmdlSave_Click()
  Dim resp As VbMsgBoxResult
  With frmedititem
    If .cmblformat.Text = "" Or .txtltitle.Text = "" Or .txtlcategory.Text = "" Or .txtlmaincast.Text = "" Or .txtlsecondcast.Text = "" Or .txtlprice.Text = "" Or .txtlnoofdays.Text = "" Or .txtldate.Text = "" Or .txtlstatus.Text = "" Or .txtlnoofcd.Text = "" Then
      MsgBox "Missing Data! Do not leave a blank textfield.", vbInformation, "Information"
      Exit Sub
    Else
      Dim res As VbMsgBoxResult
      res = MsgBox("Save this to Database?", vbYesNo, "Confirmation")
        If res = vbYes Then
          txtlprice.Text = Format(txtlprice.Text, "#####.00")
          Call edit_save_item_data
          Call cmdClose_Click
        Else
          Exit Sub
        End If
    End If
  End With
End Sub

Private Sub Form_Load()
  Call connectcombo
  Call filformcombo
  Call edit_display_item_data
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call cmdClose_Click
End Sub

Private Sub txtlcategory_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtldate_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtlmaincast_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtlnoofcd_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtlnoofdays_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtlprice_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtlsecondcast_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtlstatus_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtltitle_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
