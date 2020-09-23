VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmemwfines 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
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
      TabCaption(0)   =   "Members with Fines"
      TabPicture(0)   =   "frmmemfines.frx":0000
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
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FFFF00&
            Caption         =   "&Cancel"
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
            TabIndex        =   2
            Top             =   1080
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3000
            TabIndex        =   3
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            Format          =   24510465
            CurrentDate     =   38503
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
            Format          =   24510465
            CurrentDate     =   38473
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
            TabIndex        =   7
            Top             =   480
            Width           =   615
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
            TabIndex        =   6
            Top             =   480
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmmemwfines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
  frmMain.Enabled = True
  frmMain.Show
End Sub

Private Sub cmdPrintPreview_Click()
  
  Dim adopanerio As New ADODB.Recordset
  
  adopanerio.Open "SELECT * FROM rentreturn where datebor Between #" & DTPicker1.value & "# AND #" & DTPicker2.value & "# AND noofdayspenalty > '0' ORDER BY membershipid ASC", cnn, adOpenStatic, adLockReadOnly
  Set DataReport8.DataSource = adopanerio
  DataReport8.Sections("Section2").Controls.Item("Label15").Caption = "as of " & DTPicker1.value & " to " & DTPicker2.value
  DataReport8.Show
  Set adopanerio = Nothing
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Enabled = True
End Sub

