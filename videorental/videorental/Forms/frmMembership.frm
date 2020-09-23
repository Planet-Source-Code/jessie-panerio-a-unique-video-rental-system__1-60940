VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMembership 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   Icon            =   "frmMembership.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "frmMembership.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FF8080&
      Caption         =   "View / &Edit"
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
      Picture         =   "frmMembership.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5235
      ScaleWidth      =   10635
      TabIndex        =   16
      Top             =   240
      Width           =   10695
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3015
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Use the Navigational Button to select a record"
         Top             =   600
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777088
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
         Caption         =   "Borrowers Information"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "membershipid"
            Caption         =   "Borrowers ID#"
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
            DataField       =   "lastname"
            Caption         =   "LastName"
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
            DataField       =   "firstname"
            Caption         =   "FirstName"
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
            DataField       =   "middlename"
            Caption         =   "MiddleName"
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
            MarqueeStyle    =   1
            ScrollBars      =   0
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2534.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdMoveFirst 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   8160
         Picture         =   "frmMembership.frx":189E
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "First Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMoveNext 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   9360
         Picture         =   "frmMembership.frx":1C28
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Next Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMovePrevious 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   8760
         Picture         =   "frmMembership.frx":1FB2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Previous Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.CommandButton cmdMoveLast 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   9960
         Picture         =   "frmMembership.frx":233C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Last Record"
         Top             =   120
         Width           =   525
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   1635
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1695
         Begin VB.Image imgpic 
            Height          =   1275
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1665
         End
      End
      Begin VB.ComboBox cmbMSortOrder 
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
         ItemData        =   "frmMembership.frx":26C6
         Left            =   9480
         List            =   "frmMembership.frx":26D0
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4440
         Width           =   975
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
         Left            =   7200
         TabIndex        =   8
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtMSearchLN 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4440
         Width           =   1815
      End
      Begin VB.OptionButton optLastName 
         BackColor       =   &H00800000&
         Caption         =   "Search &Last Name"
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
         Left            =   4680
         TabIndex        =   6
         Top             =   4080
         Width           =   2175
      End
      Begin VB.OptionButton optSearchMem 
         BackColor       =   &H00800000&
         Caption         =   "Search &ID#"
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
         Left            =   2160
         TabIndex        =   4
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ComboBox cmbMSort 
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
         ItemData        =   "frmMembership.frx":26DF
         Left            =   7440
         List            =   "frmMembership.frx":26EC
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox txtMSearchMem 
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
         TabIndex        =   5
         Top             =   4440
         Width           =   1815
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      Height          =   5535
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   10935
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
      Picture         =   "frmMembership.frx":270E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
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
      Left            =   240
      Picture         =   "frmMembership.frx":2B50
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5760
      Width           =   10935
   End
End
Attribute VB_Name = "frmMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modeval As Boolean

Private Sub cmbMSort_KeyPress(KeyAscii As Integer)
  Dim strvalid
  strvalid = ""
    If KeyAscii > 26 Then
      If InStr(strvalid, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
    End If
End Sub

Private Sub cmbMSortOrder_KeyPress(KeyAscii As Integer)
  Dim strvalid
  strvalid = ""
    If KeyAscii > 26 Then
      If InStr(strvalid, Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
    End If
End Sub
    
Private Sub cmbMSort_Click()
  Call Sort
End Sub

Private Sub cmbMSortOrder_Click()
  Call Sort
End Sub

Private Sub cmdEdit_Click()
  With adomembership
    If .BOF = True And .EOF = True Then
      MsgBox "Empty Database", vbInformation, "Edit Members Record"
      Exit Sub
    Else
      frmMembership.Hide
      Load frmeditmembership
      frmeditmembership.Show
    End If
  End With
End Sub

Private Sub Form_Load()
  Call rs_act_membership
  Set frmMembership.DataGrid1.DataSource = adomembership
  Call recno
  Call LoadImage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call cmdClose_Click
End Sub

Private Sub optSearchMem_Click()
  If optSearchMem.value = True Then
    txtMSearchMem.Locked = False
    txtMSearchMem.SetFocus
    txtMSearchLN.Locked = True
    cmbMSort.Locked = True
    cmbMSortOrder.Locked = True
    txtMSearchLN.Text = ""
    cmbMSort.Text = ""
    cmbMSortOrder.Text = ""
  End If
End Sub

Private Sub optLastName_Click()
  If optLastName.value = True Then
    txtMSearchMem.Locked = True
    txtMSearchLN.Locked = False
    txtMSearchLN.SetFocus
    cmbMSort.Locked = True
    cmbMSortOrder.Locked = True
    txtMSearchMem.Text = ""
    cmbMSort.Text = ""
    cmbMSortOrder.Text = ""
  End If
End Sub
    
Private Sub optSort_Click()
  If optSort.value = True Then
    txtMSearchMem.Locked = True
    txtMSearchLN.Locked = True
    cmbMSort.Locked = False
    cmbMSort.SetFocus
    cmbMSortOrder.Locked = False
    txtMSearchMem.Text = ""
    txtMSearchLN.Text = ""
  End If
End Sub

Private Sub cmdMoveFirst_Click()
  If adomembership.RecordCount <= 1 Then Exit Sub
    adomembership.MoveFirst
    Call recno
    Call LoadImage
End Sub
    
Private Sub cmdMoveLast_Click()
  If adomembership.RecordCount <= 1 Then Exit Sub
    adomembership.MoveLast
    Call recno
    Call LoadImage
End Sub
    
Private Sub cmdMoveNext_Click()
  If adomembership.AbsolutePosition >= adomembership.RecordCount Or adomembership.RecordCount <= 1 Then Exit Sub
    adomembership.MoveNext
    Call recno
    Call LoadImage
End Sub
    
Private Sub cmdMovePrevious_Click()
  If adomembership.AbsolutePosition <= 1 Then Exit Sub
    adomembership.MovePrevious
    Call recno
    Call LoadImage
End Sub

Private Sub txtMSearchMem_Change()
  Call SearchMemID(txtMSearchMem.Text)
End Sub
    
Private Sub txtMSearchLN_Change()
  Call SearchLastName(txtMSearchLN.Text)
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Load frmMain
  frmMain.Show
  frmMain.Enabled = True
  Call Due
End Sub

Private Sub cmdNew_Click()
  frmaddmembership.Show
  frmMembership.Hide
End Sub

Private Sub cmdDelete_Click()
  Dim res As VbMsgBoxResult
    With adomembership
      If .BOF And .EOF = True Then
        MsgBox "Empty Database", vbInformation, "Delete Members Record"
        Exit Sub
      Else
        Call borrower_status
        If modeval = False Then
          res = MsgBox("Are you sure you want to Delete  " & adomembership!lastname & ", " & adomembership!firstname, vbYesNo + vbQuestion, "Confirmation")
            If res = vbYes Then
              .Delete
              .Requery
              Call clear_opt_txtbox_members
              Call LoadImage
              Call recno
                If .BOF And .EOF = True Then
                  Set imgpic = Nothing
                End If
            Else
              Exit Sub
            End If
        ElseIf modeval = True Then
          MsgBox "Cannot Delete. Borrower has UnReturned items", vbCritical, "Information"
        End If
      End If
    End With
End Sub

Private Sub borrower_status()
  On Error Resume Next
  Dim rs As New ADODB.Recordset
  rs.Open "Select * From rentreturn Where membershipid = '" & adomembership!membershipid & "' And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 1 Then
      modeval = False
      Exit Sub
    Else
      modeval = True
    End If
      Set rs = Nothing
End Sub

Private Sub txtMSearchMem_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtMSearchLN_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then KeyAscii = 0
End Sub
