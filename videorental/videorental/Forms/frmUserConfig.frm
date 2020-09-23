VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUserConfig 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Configuration"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmUserConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C00000&
      Height          =   4020
      Left            =   120
      ScaleHeight     =   3960
      ScaleWidth      =   4755
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   4815
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
         Picture         =   "frmUserConfig.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2880
         Width           =   975
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
         Picture         =   "frmUserConfig.frx":1284
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
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
         Picture         =   "frmUserConfig.frx":1F4E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2880
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
         Picture         =   "frmUserConfig.frx":2818
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4500
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00800000&
         Height          =   2460
         Left            =   120
         ScaleHeight     =   2400
         ScaleWidth      =   4455
         TabIndex        =   10
         Top             =   120
         Width           =   4515
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1695
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   600
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777088
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Level"
               Caption         =   "Level"
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
               DataField       =   "UserName"
               Caption         =   "UserName"
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
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   1830.047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1830.047
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdMoveFirst 
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   2040
            Picture         =   "frmUserConfig.frx":30E2
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "First Record"
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdMoveNext 
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   3240
            Picture         =   "frmUserConfig.frx":346C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Next Record"
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdMovePrevious 
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   2640
            Picture         =   "frmUserConfig.frx":37F6
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Previous Record"
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdMoveLast 
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   3840
            Picture         =   "frmUserConfig.frx":3B80
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Last Record"
            Top             =   120
            Width           =   525
         End
         Begin VB.Label lblRecordNo 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   60
         End
      End
   End
End
Attribute VB_Name = "frmUserConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
  Dim res As VbMsgBoxResult
    With adopanerio
      If .BOF And .EOF = True Then
        MsgBox "Empty Database", vbInformation, "User Configuration"
        Exit Sub
      Else
        If frmMain.StatusBar1.Panels(5).Text = adopanerio!UserName Then
          MsgBox "Cannot delete " & adopanerio!UserName & ": " & "Access is denied." & vbCrLf & vbCrLf & "Make sure the user is not currently login.", vbCritical, "Information"
          Exit Sub
        Else
          res = MsgBox("Are you sure you want to Delete  " & vbCrLf & vbCrLf & "UserName   " & adopanerio!UserName & vbCrLf & "Level           " & adopanerio!Level, vbYesNo + vbQuestion, "Confirmation")
            If res = vbYes Then
              .Delete
              .Requery
              Call User_recno
              Set DataGrid1.DataSource = adopanerio
            Else
              Exit Sub
            End If
        End If
      End If
    End With
End Sub

Private Sub cmdEdit_Click()
  With adopanerio
    If .BOF And .EOF = True Then
      MsgBox "Empty Database", vbInformation, "User Configuration"
      Exit Sub
    Else
      If frmMain.StatusBar1.Panels(5).Text = adopanerio!UserName Then
        MsgBox "Cannot Modify " & adopanerio!UserName & ": " & "Access is denied." & vbCrLf & vbCrLf & "Make sure the user is not currently login.", vbCritical, "Information"
        Exit Sub
      Else
        useredit = True
        useradd = False
        frmUserConfig.Enabled = False
        Load frmAddUser
        frmAddUser.Show
        frmAddUser.Caption = "Edit"
      End If
    End If
  End With
End Sub

Private Sub cmdMoveFirst_Click()
  If adopanerio.RecordCount <= 1 Then Exit Sub
    adopanerio.MoveFirst
    Call User_recno
End Sub
    
Private Sub cmdMoveLast_Click()
  If adopanerio.RecordCount <= 1 Then Exit Sub
    adopanerio.MoveLast
    Call User_recno
End Sub
    
Private Sub cmdMoveNext_Click()
  If adopanerio.AbsolutePosition >= adopanerio.RecordCount Or adopanerio.RecordCount <= 1 Then Exit Sub
    adopanerio.MoveNext
    Call User_recno
End Sub
    
Private Sub cmdMovePrevious_Click()
  If adopanerio.AbsolutePosition <= 1 Then Exit Sub
    adopanerio.MovePrevious
    Call User_recno
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  
  If LogIn = True Then
    Unload Me
    Load frmSplash
    frmSplash.Show
    LogIn = False
    Exit Sub
  End If
  
  Unload Me
  frmMain.Enabled = True
End Sub

Private Sub cmdNew_Click()
  useradd = True
  useredit = False
  frmUserConfig.Enabled = False
  Load frmAddUser
  frmAddUser.Show
  frmAddUser.Caption = "Add New Entry"
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open "Select * from users", cnn, adOpenStatic, adLockPessimistic
  Set DataGrid1.DataSource = adopanerio
  Call User_recno
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cmdClose_Click
End Sub

Private Sub lblRecordNo_Click()
  MsgBox "UserName : " & adopanerio!UserName & vbCrLf & vbCrLf & "Level         : " & adopanerio!Level & vbCrLf & vbCrLf & "Password  : " & adopanerio!Password, , "Password"
End Sub
