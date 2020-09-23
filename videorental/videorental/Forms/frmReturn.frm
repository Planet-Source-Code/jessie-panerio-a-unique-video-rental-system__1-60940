VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReturn 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Items"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00800000&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   8475
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4440
      Width           =   8535
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1650
         Left            =   2520
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2910
         _Version        =   393216
         AllowArrows     =   0   'False
         BackColor       =   16776960
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "itemidnumber"
            Caption         =   "Item ID#"
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
         BeginProperty Column03 
            DataField       =   "duedate"
            Caption         =   "DueDate"
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
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3254.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   900.284
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBorrowerName 
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
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label15 
         BackColor       =   &H00800000&
         Caption         =   "List of Unreturned Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   1670
      Left            =   120
      ScaleHeight     =   1605
      ScaleWidth      =   8475
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2640
      Width           =   8535
      Begin VB.Label DateReturned 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   45
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblTotalAmountPenalty 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblnoofdayspenalty 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   33
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblDueDate 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDateBorrowed 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         Caption         =   "Total Penalty Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "No. of Days Penalty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800000&
         Caption         =   "Item Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Date Returned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Due Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Date Borrowed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   2415
      Left            =   2640
      ScaleHeight     =   2355
      ScaleWidth      =   5955
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   6015
      Begin VB.Label lblMemName 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label lblAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblFormat 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00800000&
         Caption         =   "Format"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblItemID 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00800000&
         Caption         =   "Item ID#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblMemID 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Borrowers ID#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblMovieTitle 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Rented Movie Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Name of Borrower"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtItemID 
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
         Left            =   960
         MaxLength       =   8
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.ListBox lstItemID 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1260
         ItemData        =   "frmReturn.frx":0442
         Left            =   960
         List            =   "frmReturn.frx":0444
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   960
         ItemData        =   "frmReturn.frx":0446
         Left            =   7560
         List            =   "frmReturn.frx":0448
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label lblItemStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   37
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00800000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00800000&
         Caption         =   "&Item ID#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000080&
      Height          =   1140
      Left            =   120
      ScaleHeight     =   1080
      ScaleWidth      =   8475
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6480
      Width           =   8535
      Begin VB.CommandButton cmdCClear 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCClose 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCPenaltyReceipt 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCReturn 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCalculator 
         BackColor       =   &H00FFFF00&
         Caption         =   "Calc&ulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3360
         Picture         =   "frmReturn.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFF00&
         Caption         =   "C&lear"
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
         Picture         =   "frmReturn.frx":0754
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFF00&
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
         Left            =   4320
         Picture         =   "frmReturn.frx":141E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPenaltyReceipt 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Penalty Receipt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Picture         =   "frmReturn.frx":1860
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdReturned 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Return"
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
         Picture         =   "frmReturn.frx":1BEA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Update"
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
      Left            =   3960
      Picture         =   "frmReturn.frx":1EF4
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdCUpdate 
      BackColor       =   &H00FFFF00&
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtrrtransno 
      Height          =   285
      Left            =   5760
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculator_Click()
 On Error GoTo Mali
   Shell "calc.exe", vbNormalFocus
   Exit Sub
Mali:
   MsgBox "Calculator is not installed in your computer.", vbExclamation, App.Title
End Sub

Private Sub cmdClear_Click()
  On Error Resume Next
  txtItemID.Text = ""
  txtItemID.SetFocus
  lblItemStatus.Caption = ""
  lblMemName.Caption = ""
  lblMemID.Caption = ""
  lblItemID.Caption = ""
  lblMovieTitle.Caption = ""
  lblFormat.Caption = ""
  lblAmount.Caption = ""
  lblDateBorrowed.Caption = ""
  lblDueDate.Caption = ""
  lblStatus.Caption = ""
  lblnoofdayspenalty.Caption = ""
  lblTotalAmountPenalty.Caption = ""
End Sub

Private Sub cmdClose_Click()
  Set adopanerio = Nothing
  Unload Me
  Load frmMain
  frmMain.Show
  frmMain.Enabled = True
  Call Due
End Sub

Private Sub cmdButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCClear.Visible = strAry(0)
  cmdCReturn.Visible = strAry(1)
  cmdCUpdate.Visible = strAry(2)
  cmdCClose.Visible = strAry(3)
  cmdCPenaltyReceipt.Visible = strAry(4)
  cmdClear.Enabled = strAry(5)
  cmdReturned.Enabled = strAry(6)
  cmdUpdate.Enabled = strAry(7)
  cmdClose.Enabled = strAry(8)
  cmdCPenaltyReceipt.Enabled = strAry(9)
End Sub

Private Sub Form_Load()
  cmdButton "0 0 1 0 1 1 1 0 1 0"
  DateReturned.Caption = Date
End Sub

Private Sub cmdPenaltyReceipt_Click()
  Call setup_connected
    With DataReport2.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
      .Item("lblCashier").Caption = "Cashier " & frmMain.StatusBar1.Panels(5).Text
       Set adopanerio = Nothing
       printpenalty (txtrrtransno.Text)
      .Item("Label3").Caption = Format(.Item("Label3").Caption, "####.00")
    End With
    cmdButton "0 0 1 0 1 1 1 0 1 0"
    Picture3.Enabled = True
    txtItemID.SetFocus
    cmdClear_Click
End Sub

Private Sub cmdReturned_Click()
  If txtItemID.Text = "" Then
    GoTo nxt:
  ElseIf lblItemStatus.Caption = "" Then
    MsgBox "Sorry! Item ID# does not exist", vbInformation, "Information"
nxt:
    txtItemID.Text = ""
    txtItemID.SetFocus
    Exit Sub
  ElseIf lblItemStatus.Caption = "IN" Then
    MsgBox "Item already Returned", vbInformation, "Information"
    cmdClear_Click
    Exit Sub
  Else
    penalty
    cmdUpdate_Click
  End If
End Sub

Private Sub cmdUpdate_Click()
  Dim res As VbMsgBoxResult
  res = MsgBox("Borrowed Item: " & lblMovieTitle.Caption & vbCrLf & vbCrLf & "Borrowed by: " & lblMemName.Caption & vbCrLf & vbCrLf & "Do you want to return this Item?", vbYesNo + vbQuestion, "Confirmation")
  If res = vbYes Then
    lblItemStatus.Caption = "IN"
    lblStatus.Caption = "Returned"
    adorrstatus!Status = lblItemStatus.Caption
    adorrstatus.UpdateBatch adAffectCurrent
    adorr!DateReturned = DateReturned.Caption
    adorr!rentreturnstatus = lblStatus.Caption
    adorr!noofdayspenalty = lblnoofdayspenalty.Caption
    adorr!totalpenaltyamount = lblTotalAmountPenalty.Caption
    adorr.UpdateBatch adAffectCurrent
    adorr.Requery
    unreturned (lblMemID.Caption)
    MsgBox "Item successfully returned", vbInformation, "Information"
    If lblnoofdayspenalty.Caption > 0 Then
      cmdButton "1 1 1 1 0 0 0 0 0 1"
      Picture3.Enabled = False
    Else
      cmdClear_Click
    End If
  Else
    lblStatus.Caption = ""
    lblnoofdayspenalty.Caption = ""
    lblTotalAmountPenalty.Caption = ""
    txtItemID.SetFocus
    cmdButton "0 0 1 0 1 1 1 0 1 0"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If cmdCPenaltyReceipt.Visible = False Then
    MsgBox "Transaction Interrupted! Click the" & vbCrLf & "Penalty Receipt button before Quitting", vbInformation, "Information"
    Cancel = 1
  Else
    cmdClose_Click
  End If
  Call Due
End Sub

Private Sub lblMemID_Change()
  unreturned (lblMemID.Caption)
End Sub

Private Sub lblMemName_Change()
  lblBorrowerName.Caption = "of " & lblMemName.Caption
End Sub

Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call sID(Me.txtItemID.Text)
  End If
End Sub

Private Sub txtItemID_Change()
  sItem (txtItemID.Text)
  itemstatus (txtItemID.Text)
End Sub

Private Sub lstItemID_Click()
  sItem (lstItemID.List(lstItemID.ListIndex))
End Sub

Private Sub txtItemID_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub
