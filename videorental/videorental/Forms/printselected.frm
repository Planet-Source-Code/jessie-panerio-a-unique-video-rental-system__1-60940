VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmprintselmem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print "
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "printselected.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "printselected.frx":628A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C00000&
         Height          =   2200
         Left            =   -120
         ScaleHeight     =   2145
         ScaleWidth      =   4635
         TabIndex        =   1
         Top             =   360
         Width           =   4695
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FFFF00&
            Caption         =   "Canc&el"
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
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1080
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3000
            TabIndex        =   6
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   38717
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   38504
         End
         Begin VB.CommandButton cmdPrintPreview 
            BackColor       =   &H00FFFF00&
            Caption         =   "&Print Preview"
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
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "up to"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   5
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmprintselmem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
  frmMain.Enabled = True
  frmMain.Show
End Sub

Private Sub cmdPrintPreview_Click()
  Dim adopanerio As New ADODB.Recordset
  If wfine = True And selek = False Then
    adopanerio.Open "SELECT * FROM rentreturn where datebor Between #" & DTPicker1.value & "# AND #" & DTPicker2.value & "# AND noofdayspenalty > '0' ORDER BY membershipid ASC", cnn, adOpenStatic, adLockReadOnly
    Set DataReport8.DataSource = adopanerio
    Call name_addr
    DataReport8.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.value & " to " & DTPicker2.value
    DataReport8.Show
    Set adopanerio = Nothing
  ElseIf selek = True And wfine = False Then
    adopanerio.Open "SELECT * FROM membership WHERE date Between #" & DTPicker1.value & "# AND #" & DTPicker2.value & "# ORDER BY lastname", cnn, adOpenStatic, adLockReadOnly
    Set DataReport3.DataSource = adopanerio
    Call name_addr
    DataReport3.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.value & " to " & DTPicker2.value
    DataReport3.Show
    Set adopanerio = Nothing
  ElseIf selek = False And wfine = False Then
    adopanerio.Open "SELECT * FROM rentreturn where datebor Between #" & DTPicker1.value & "# AND #" & DTPicker2.value & "# AND rentreturnstatus = 'UnReturned' AND duedate > #" & Date & "# ORDER BY itemidnumber", cnn, adOpenStatic, adLockReadOnly
    Set DataReport9.DataSource = adopanerio
    DataReport9.Sections("Section2").Controls.Item("lblasof").Caption = "as of " & DTPicker1.value & " to " & DTPicker2.value
    Call InOut
    DataReport9.Show
    Set adopanerio = Nothing
  End If
End Sub

Private Sub name_addr()
  Call setup_connected
  If wfine = True And selek = False Then
    With DataReport8.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
    End With
  ElseIf selek = True And wfine = False Then
    With DataReport3.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
    End With
  End If
  Set adopanerio = Nothing
End Sub

Private Sub Form_Load()
  DTPicker1.value = Date
  DTPicker2.value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Enabled = True
End Sub
