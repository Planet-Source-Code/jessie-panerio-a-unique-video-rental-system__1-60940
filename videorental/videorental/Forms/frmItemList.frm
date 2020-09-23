VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmItemList 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmItemList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPriceRate 
      BackColor       =   &H00FF8080&
      Caption         =   "&Penalty Rate"
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
      Left            =   4560
      Picture         =   "frmItemList.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdFormat 
      BackColor       =   &H00FF8080&
      Caption         =   "&Format"
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
      Picture         =   "frmItemList.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   11355
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   240
      Width           =   11415
      Begin VB.CommandButton cmdMoveLast 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   10680
         Picture         =   "frmItemList.frx":1A16
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Last Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMoveNext 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   10080
         Picture         =   "frmItemList.frx":1DA0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Next Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMovePrevious 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   9480
         Picture         =   "frmItemList.frx":212A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Previous Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMoveFirst 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   8880
         Picture         =   "frmItemList.frx":24B4
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "First Record"
         Top             =   120
         Width           =   525
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   16777088
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   2
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "List of Items "
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "itemid"
            Caption         =   "Item ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "title"
            Caption         =   "Movie Title"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "status"
            Caption         =   "Status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "format"
            Caption         =   "Format"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "category"
            Caption         =   "Category"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "noofdays"
            Caption         =   "# of Days"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "maincast"
            Caption         =   "Maincast"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "secondcast"
            Caption         =   "Secondcast"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "datepurchase"
            Caption         =   "Date Purchased"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "price"
            Caption         =   "Price of Item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "noofcd"
            Caption         =   "No. of CD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   5
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5054.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1769.953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1305.071
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtISearchCategory 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4440
         Width           =   1815
      End
      Begin VB.OptionButton optSearchCategory 
         BackColor       =   &H00800000&
         Caption         =   "Search C&ategory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   8
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtISearchItem 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4440
         Width           =   1575
      End
      Begin VB.ComboBox cmbISort 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmItemList.frx":283E
         Left            =   8520
         List            =   "frmItemList.frx":284E
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4440
         Width           =   1695
      End
      Begin VB.OptionButton optSearchItem 
         BackColor       =   &H00800000&
         Caption         =   "Search &Item ID#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton optSearchTitle 
         BackColor       =   &H00800000&
         Caption         =   "Search Movie &Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtISearchMovie 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4440
         Width           =   3015
      End
      Begin VB.OptionButton optSort 
         BackColor       =   &H00800000&
         Caption         =   "&Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   10
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox cmbISortOrder 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmItemList.frx":2873
         Left            =   10320
         List            =   "frmItemList.frx":287D
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblRecordNoA 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      Height          =   5295
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   11655
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
      Left            =   2400
      Picture         =   "frmItemList.frx":288C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FF8080&
      Caption         =   "&Edit"
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
      Picture         =   "frmItemList.frx":3556
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Left            =   5640
      Picture         =   "frmItemList.frx":3E20
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
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
      Picture         =   "frmItemList.frx":4262
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Width           =   11655
   End
End
Attribute VB_Name = "frmItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
  Dim res As VbMsgBoxResult
    With adoitemlist
      If .BOF And .EOF = True Then
        MsgBox "Empty Database", vbInformation, "Item List"
        Exit Sub
      ElseIf adoitemlist!Status = "OUT" Then
        MsgBox "Cannot Delete. Item is Out", vbCritical, "Information"
        Exit Sub
      Else
        res = MsgBox("Are you sure you want to Delete  " & adoitemlist!Title & "?", vbYesNo + vbQuestion, "Confirmation")
          If res = vbYes Then
            .Delete
            .Requery
            Call recnoIL
            Call clear_opt_txtbox_itemlist
          Else
             Exit Sub
          End If
      End If
    End With
End Sub

Private Sub cmdEdit_Click()
  If adoitemlist.BOF = True Or adoitemlist.EOF = True Then
    MsgBox "Empty Database", vbInformation, "Edit Item"
  Else
    Load frmItemList
    frmedititem.Show
    frmItemList.Enabled = False
   End If
End Sub

Private Sub cmdFormat_Click()
  frmformat.Show
  frmItemList.Enabled = False
End Sub

Private Sub cmdPriceRate_Click()
  frmRate.Show
  frmItemList.Enabled = False
End Sub

Private Sub Form_Load()
  Call rs_act_itemlist
  Set DataGrid1.DataSource = adoitemlist
  Call recnoIL
End Sub

Private Sub cmdNew_Click()
  frmadditem.Show
  frmItemList.Enabled = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Load frmMain
  frmMain.Show
  frmMain.Enabled = True
  Call Due
End Sub

Private Sub cmdMoveFirst_Click()
  If adoitemlist.RecordCount <= 1 Then Exit Sub
    adoitemlist.MoveFirst
    Call recnoIL
End Sub
    
Private Sub cmdMoveLast_Click()
  If adoitemlist.RecordCount <= 1 Then Exit Sub
    adoitemlist.MoveLast
    Call recnoIL
End Sub
    
Private Sub cmdMoveNext_Click()
  If adoitemlist.AbsolutePosition >= adoitemlist.RecordCount Or adoitemlist.RecordCount <= 1 Then Exit Sub
    adoitemlist.MoveNext
    Call recnoIL
End Sub
    
Private Sub cmdMovePrevious_Click()
  If adoitemlist.AbsolutePosition <= 1 Then Exit Sub
    adoitemlist.MovePrevious
    Call recnoIL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call cmdClose_Click
End Sub

Private Sub txtISearchCategory_Change()
  Call SearchCategory(txtISearchCategory.Text)
End Sub
    
Private Sub txtISearchItem_Change()
  Call SearchItemID(txtISearchItem.Text)
End Sub
    
Private Sub txtISearchMovie_Change()
  Call SearchMovieTitle(txtISearchMovie.Text)
End Sub

Private Sub cmbISort_Click()
  Call SortIL
End Sub
   
Private Sub cmbISortOrder_Click()
  Call SortIL
End Sub

Private Sub optSearchItem_Click()
  If optSearchItem.value = True Then
    txtISearchItem.Locked = False
    txtISearchItem.SetFocus
    txtISearchMovie.Locked = True
    txtISearchMovie.Text = ""
    txtISearchCategory.Locked = True
    txtISearchCategory.Text = ""
    cmbISort.Locked = True
    cmbISort.Text = ""
    cmbISortOrder.Locked = True
    cmbISortOrder.Text = ""
  End If
End Sub

Private Sub optSearchTitle_Click()
  If optSearchTitle.value = True Then
    txtISearchItem.Locked = True
    txtISearchItem.Text = ""
    txtISearchMovie.Locked = False
    txtISearchMovie.SetFocus
    txtISearchCategory.Locked = True
    txtISearchCategory.Text = ""
    cmbISort.Locked = True
    cmbISort.Text = ""
    cmbISortOrder.Locked = True
    cmbISortOrder.Text = ""
  End If
End Sub
    
Private Sub optSearchCategory_Click()
  If optSearchCategory.value = True Then
    txtISearchItem.Locked = True
    txtISearchItem.Text = ""
    txtISearchMovie.Locked = True
    txtISearchMovie.Text = ""
    txtISearchCategory.Locked = False
    txtISearchCategory.SetFocus
    cmbISort.Locked = True
    cmbISort.Text = ""
    cmbISortOrder.Locked = True
    cmbISortOrder.Text = ""
  End If
End Sub
    
Private Sub optSort_Click()
  If optSort.value = True Then
    txtISearchItem.Locked = True
    txtISearchItem.Text = ""
    txtISearchMovie.Locked = True
    txtISearchMovie.Text = ""
    txtISearchCategory.Locked = True
    txtISearchCategory.Text = ""
    cmbISort.Locked = False
    cmbISort.SetFocus
    cmbISortOrder.Locked = False
  End If
End Sub

Private Sub txtISearchItem_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtISearchMovie_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtISearchCategory_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
