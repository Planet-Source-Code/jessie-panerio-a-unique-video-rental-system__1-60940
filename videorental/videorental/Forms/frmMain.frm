VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   Caption         =   "Video Rental System"
   ClientHeight    =   7905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   7905
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7500
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   15
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   476
            MinWidth        =   476
            Picture         =   "frmMain.frx":1F030
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3246
            MinWidth        =   3246
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1235
            MinWidth        =   1235
            Picture         =   "frmMain.frx":1FE84
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "User Name"
            TextSave        =   "User Name"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3246
            MinWidth        =   3246
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1588
            MinWidth        =   1588
            Picture         =   "frmMain.frx":20CD8
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log In"
            TextSave        =   "Time Log In"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1235
            MinWidth        =   1235
            Picture         =   "frmMain.frx":21B2C
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Date"
            TextSave        =   "Date"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2547
            MinWidth        =   2547
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1234
            MinWidth        =   1234
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel15 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblItemsOut 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8550
      TabIndex        =   5
      Top             =   5880
      Width           =   75
   End
   Begin VB.Label lblItemsIn 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8550
      TabIndex        =   4
      Top             =   5280
      Width           =   75
   End
   Begin VB.Label lblOverdueItems 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8550
      TabIndex        =   3
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label lblDueItemsToday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8550
      TabIndex        =   2
      Top             =   4080
      Width           =   75
   End
   Begin VB.Label lblTotalItems 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8550
      TabIndex        =   1
      Top             =   6480
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuMembership 
         Caption         =   "&Borrower"
         Begin VB.Menu mnuMembers 
            Caption         =   "&Configuration"
            Shortcut        =   {F2}
         End
      End
      Begin VB.Menu mnuItemList 
         Caption         =   "&Items"
         Begin VB.Menu mnuItemS 
            Caption         =   "&Configuration"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock &Application"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   ""
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuAuthor 
         Caption         =   "&Author"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuRent 
         Caption         =   "&Rent"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "R&eturn"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuRMembers 
         Caption         =   "&Borrower"
         Begin VB.Menu mnuRAllMembers 
            Caption         =   "&All Borrower"
         End
         Begin VB.Menu mnuMSelected 
            Caption         =   "&Selected"
         End
         Begin VB.Menu mnuMemWFines 
            Caption         =   "with &Fines"
         End
      End
      Begin VB.Menu mnuRItems 
         Caption         =   "&Items"
         Begin VB.Menu mnuRAllItems 
            Caption         =   "&All Items"
         End
         Begin VB.Menu mnuItemStatus 
            Caption         =   "Item &Status"
            Begin VB.Menu mnuDueToday 
               Caption         =   "&Due Items Today"
            End
            Begin VB.Menu mnuROverdueItems 
               Caption         =   "O&verdue Items"
            End
            Begin VB.Menu mnuRItemsIn 
               Caption         =   "All Items &In"
            End
            Begin VB.Menu mnuOut 
               Caption         =   "Items &Out"
               Begin VB.Menu mnuRItemsOut 
                  Caption         =   "&All "
               End
               Begin VB.Menu mnuSelekOut 
                  Caption         =   "&Selected"
               End
            End
         End
      End
      Begin VB.Menu mnuUsersReport 
         Caption         =   "&Users Time Record"
         Begin VB.Menu mnuUsersAll 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuUsersAdmin 
            Caption         =   "A&dministrator"
         End
         Begin VB.Menu mnuUsersEmp 
            Caption         =   "&Employee"
         End
      End
   End
   Begin VB.Menu mnuAdministrator 
      Caption         =   "&Administrator"
      Begin VB.Menu mnuConfig 
         Caption         =   "&System Configuration"
         Begin VB.Menu mnuSetup 
            Caption         =   "&Video Rental Setup"
         End
         Begin VB.Menu mnuUser 
            Caption         =   "&Users Configuration"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
  Call Due
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("This will terminate the application. Proceed?", vbOKCancel + vbQuestion, "Video Rental System") = vbOK Then
    Call User_LogOut
    cnn.Close
    End
  Else
    Cancel = 1
  End If
End Sub

Private Sub mnuAdministrator_Click()
  If frmMain.StatusBar1.Panels(2).Text = "Administrator" Then
    Exit Sub
  Else
    MsgBox "Access Denied!", vbCritical, "Restricted Area"
  End If
End Sub

Private Sub mnuAuthor_Click()
  Load frmCredits
  frmCredits.Show
  frmMain.Enabled = False
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHelp_Click()
  Call WinHelp(0, App.HelpFile, HELPC, 0)
End Sub

Private Sub mnuItems_Click()
  frmItemList.Show
  frmMain.Enabled = False
End Sub

Private Sub mnuLock_Click()
  Load frmLocked
  frmLocked.Show
  frmMain.Enabled = False
End Sub

Private Sub mnuLogOff_Click()
  Dim res As VbMsgBoxResult
  res = MsgBox("Are you sure you want to log off?", vbQuestion + vbYesNo, "Log Off " & frmMain.StatusBar1.Panels(5).Text)
  If res = vbYes Then
    Call User_LogOut
    frmMain.StatusBar1.Panels(2).Text = "Waiting..."
    frmMain.StatusBar1.Panels(5).Text = "Waiting..."
    frmMain.StatusBar1.Panels(8).Text = "Waiting..."
    frmMain.StatusBar1.Panels(11).Text = "Waiting..."
    frmMain.Enabled = False
    Load frmLogin
    frmLogin.Show
  Else
    Exit Sub
  End If
End Sub

Private Sub mnuMembers_Click()
  frmMain.Enabled = False
  frmMembership.Show
End Sub

Private Sub mnuMemWFines_Click()
  wfine = True
  selek = False
  frmprintselmem.Show
  frmMain.Enabled = False
  frmprintselmem.SSTab1.Caption = " Borrower with Fines"
End Sub

Private Sub mnuMSelected_Click()
  selek = True
  wfine = False
  frmprintselmem.Show
  frmMain.Enabled = False
  frmprintselmem.SSTab1.Caption = " Selected"
End Sub

Private Sub mnuRent_Click()
  frmMain.Enabled = False
  frmRent.Show
End Sub

Private Sub mnuReturn_Click()
  frmMain.Enabled = False
  frmReturn.Show
End Sub

Private Sub mnuDueToday_Click()
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM rentreturn where duedate = #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
  Set DataReport7.DataSource = adopanerio
  Call duetoday
  DataReport7.Show
  Set adopanerio = Nothing
End Sub

Private Sub duetoday()
  Call setup_connected
    With DataReport7.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
    End With
    Set adopanerio = Nothing
End Sub

Private Sub mnuRItemsIn_Click()
  wfine = True
  selek = True
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM itemlist where status = 'IN' ORDER BY itemid", cnn, adOpenStatic, adLockReadOnly
  Set DataReport6.DataSource = adopanerio
  Call InOut
  DataReport6.Show
  Set adopanerio = Nothing
End Sub

Private Sub mnuRItemsOut_Click()
  wfine = False
  selek = False
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM rentreturn where duedate > #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
  Set DataReport9.DataSource = adopanerio
  Call InOut
  DataReport9.Show
  Set adopanerio = Nothing
End Sub

Private Sub mnuRAllItems_Click()
  On Error Resume Next
  wfine = True
  selek = False
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM itemlist ORDER BY itemid", cnn, adOpenStatic, adLockReadOnly
  Set DataReport5.DataSource = adopanerio
  Call All
  DataReport5.Show
  Set adopanerio = Nothing
End Sub

Private Sub mnuRAllMembers_Click()
  selek = True
  wfine = False
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM membership ORDER BY membershipid", cnn, adOpenStatic, adLockReadOnly 'adLockPessimistic
  Set DataReport3.DataSource = adopanerio
  Call All
  DataReport3.Show
  Set adopanerio = Nothing
End Sub

Private Sub All()
  Call setup_connected
    If wfine = True And selek = False Then
      With DataReport5.Sections("Section2").Controls
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

Private Sub mnuROverdueItems_Click()
  On Error Resume Next
  Dim adopanerio As New ADODB.Recordset
  adopanerio.Open "SELECT * FROM rentreturn where duedate < #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
  Set DataReport4.DataSource = adopanerio
  Call overdue
  With DataReport4.Sections("Section1").Controls
    .Item("lblnoofdays").Caption = Date - CDate(adopanerio!duedate)
  End With
  DataReport4.Show
  Set adopanerio = Nothing
End Sub

Private Sub overdue()
  Call setup_connected
      With DataReport4.Sections("Section2").Controls
        .Item("lblName").Caption = adopanerio!nname
        .Item("lblAddr").Caption = adopanerio!address
      End With
  Set adopanerio = Nothing
End Sub
  
Private Sub mnuSelekOut_Click()
  wfine = False
  selek = False
  frmprintselmem.Show
  frmMain.Enabled = False
  frmprintselmem.SSTab1.Caption = " Selected Items OUT"
End Sub

Private Sub mnuSetup_Click()
  frmMain.Enabled = False
  Load frmVideoRentalSetup
  frmVideoRentalSetup.Show
End Sub

Private Sub mnuUser_Click()
  frmMain.Enabled = False
  Load frmUserConfig
  frmUserConfig.Show
End Sub

Private Sub mnuUsersAll_Click()
  frmMain.Enabled = False
  Load frmUsers
  frmUsers.Show
  frmUsers.SSTab1.Caption = " All Users"
  AllUsers = True
End Sub

Private Sub mnuUsersAdmin_Click()
  frmMain.Enabled = False
  Load frmUsers
  frmUsers.Show
  frmUsers.SSTab1.Caption = " Administrator"
  Admin = True
End Sub

Private Sub mnuUsersEmp_Click()
  frmMain.Enabled = False
  Load frmUsers
  frmUsers.Show
  frmUsers.SSTab1.Caption = " Employee"
  Emp = True
End Sub
