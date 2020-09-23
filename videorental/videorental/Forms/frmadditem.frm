VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmadditem 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Entry"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmadditem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCNew 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCCancel 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCSave 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdlnew 
      BackColor       =   &H00FF8080&
      Caption         =   "Add &New"
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
      Picture         =   "frmadditem.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Picture         =   "frmadditem.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   975
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
      Left            =   2280
      Picture         =   "frmadditem.frx":23D6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCClose 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   4995
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   480
      Width           =   5055
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
         ForeColor       =   &H00000000&
         Height          =   390
         ItemData        =   "frmadditem.frx":30A0
         Left            =   1800
         List            =   "frmadditem.frx":30A2
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFormatRentPrice 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtnoofcd 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   15
         Top             =   4680
         Width           =   495
      End
      Begin VB.TextBox txtstatus 
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
         Height          =   405
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4680
         Width           =   615
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   7
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtlprice 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   13
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   11
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   9
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   5
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtListAutonum 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text7"
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   14
         Top             =   4800
         Width           =   1470
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
         TabIndex        =   31
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Date Purchased"
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
         TabIndex        =   24
         Top             =   3840
         Width           =   1500
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
         TabIndex        =   6
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
         TabIndex        =   12
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
         TabIndex        =   10
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
         TabIndex        =   8
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
         TabIndex        =   4
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
         TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
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
      TabPicture(0)   =   "frmadditem.frx":30A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
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
      Left            =   3360
      Picture         =   "frmadditem.frx":3D7E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
      Width           =   5655
   End
End
Attribute VB_Name = "frmadditem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCNew.Visible = strAry(0)
  cmdCCancel.Visible = strAry(1)
  cmdCSave.Visible = strAry(2)
  cmdCClose.Visible = strAry(3)
  cmdlnew.Enabled = strAry(4)
  cmdlcancel.Enabled = strAry(5)
  cmdlsave.Enabled = strAry(6)
  cmdClose.Enabled = strAry(7)
End Sub

Private Sub Form_Load()
  cmdButton "0 1 1 0 1 0 0 1"
  Call cleartxtbox
  Call locktxtbox(True)
End Sub

Private Sub cmblformat_Click()
  On Error Resume Next
  adopanerio.MoveFirst
  adopanerio.Move cmblformat.ListIndex
  txtFormatRentPrice.Text = adopanerio!price
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  Load frmItemList
  frmItemList.Show
  frmItemList.Enabled = True
  Call clear_opt_txtbox_itemlist
  Call recnoIL
End Sub

Private Sub cmdlnew_Click()
  Call cleartxtbox
  Call rs_act_autonum
  With adoautonum
    txtListAutonum.Text = !itemnum
  End With
  txtListAutonum.Text = Val(txtListAutonum.Text) + 1
  txtListAutonum.Text = Format(txtListAutonum.Text, "00000000")
  txtlitemid.Text = txtListAutonum.Text
  txtDate.Text = Date
  txtstatus.Text = "IN"
  cmblformat.clear
  Call connectcombo
  Call fillformcombo
  cmdButton "1 0 0 1 0 1 1 0"
  Call locktxtbox(False)
  txtstatus.Locked = True
End Sub

Private Sub cmdlSave_Click()
  Dim resp As VbMsgBoxResult
    With frmadditem
       If .cmblformat.Text = "" Or .txtltitle.Text = "" Or .txtlcategory.Text = "" Or .txtlmaincast.Text = "" Or .txtlsecondcast.Text = "" Or .txtlprice.Text = "" Or .txtlnoofdays.Text = "" Then
          MsgBox "Missing Data! Do not leave a blank textfield.", vbInformation, "Information"
          Exit Sub
       Else
          Dim res As VbMsgBoxResult
          res = MsgBox("Save this to Database?", vbYesNo + vbQuestion, "Confirmation")
           If res = vbYes Then
             txtlprice.Text = Format(txtlprice.Text, "#####.00")
             adoitemlist.AddNew
             Call WriteDataFromControlslist
             adoitemlist.Update
             adoautonum.UpdateBatch adAffectCurrent
             cmdButton "0 1 1 0 1 0 0 1"
             Call locktxtbox(True)
           Else
             Exit Sub
           End If
       End If
    End With
End Sub

Private Sub cmdlCancel_Click()
  With frmadditem
   .txtListAutonum.Text = Val(.txtListAutonum.Text) - 1
   .txtListAutonum.Text = Format(.txtListAutonum.Text, "00000000")
   .txtlitemid.Text = ""
   .txtstatus.Text = ""
  End With
   cmdButton "0 1 1 0 1 0 0 1"
   Call cleartxtbox
   Call locktxtbox(True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdCNew.Visible = True Then
    adoitemlist.CancelUpdate
    MsgBox "Add New Entry Cancelled", vbInformation, "Item List"
    Call cmdClose_Click
  Else
    Call cmdClose_Click
  End If
End Sub

Private Sub cleartxtbox()
  txtlitemid.Text = ""
  cmblformat.Text = ""
  txtFormatRentPrice.Text = ""
  txtltitle.Text = ""
  txtlcategory.Text = ""
  txtlmaincast.Text = ""
  txtlsecondcast.Text = ""
  txtDate.Text = ""
  txtnoofcd.Text = ""
  txtstatus.Text = ""
  txtlprice.Text = ""
  txtlnoofdays.Text = ""
End Sub

Private Sub locktxtbox(value As String)
  cmblformat.Locked = value
  txtltitle.Locked = value
  txtlcategory.Locked = value
  txtlmaincast.Locked = value
  txtlsecondcast.Locked = value
  txtnoofcd.Locked = value
  txtstatus.Locked = value
  txtlprice.Locked = value
  txtlnoofdays.Locked = value
End Sub

Private Sub txtlcategory_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 47 Or KeyAscii = 32 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtlmaincast_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
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

Private Sub txtltitle_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnoofcd_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
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
