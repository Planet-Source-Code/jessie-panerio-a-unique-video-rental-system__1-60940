VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmCredits 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   210
      Left            =   5040
      TabIndex        =   0
      Top             =   5280
      Width           =   180
      ExtentX         =   317
      ExtentY         =   370
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Kill App.Path & "\tmpfile.swf"
  Unload Me
  Load frmMain
  frmMain.Show
  frmMain.Enabled = True
End Sub

Private Sub Form_Load()
 On Error Resume Next
  Kill App.Path & "\tmpfile.swf"
  LoadDataIntoFile 101, App.Path & "\tmpfile.swf"
  WebBrowser1.Navigate App.Path & "\tmpfile.swf"
End Sub

Private Sub Form_Resize()
  WebBrowser1.Left = Me.ScaleLeft
  WebBrowser1.Top = Me.ScaleTop
  WebBrowser1.Width = Me.ScaleWidth
  WebBrowser1.Height = Me.ScaleHeight
End Sub
