VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRent 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rent Items"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmRent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00800000&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   8715
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3000
      Width           =   8775
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00800000&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   8355
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1800
         Width           =   8415
         Begin VB.CommandButton cmdCClose 
            BackColor       =   &H00FFFF00&
            Height          =   855
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
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
            Left            =   6360
            Picture         =   "frmRent.frx":0E42
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00FFFF00&
            Caption         =   "&Close"
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
            Left            =   7320
            Picture         =   "frmRent.frx":114C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdCOR 
            BackColor       =   &H00FFFF00&
            Height          =   855
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtItemCount 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton cmdPrintOR 
            BackColor       =   &H00FFFF00&
            Caption         =   "&Print Receipt"
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
            Left            =   5400
            Picture         =   "frmRent.frx":158E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   2  'Center
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
            Height          =   405
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtA 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtTitl 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtM 
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label12 
            BackColor       =   &H00800000&
            Caption         =   "Total Item/s Rented"
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
            Left            =   5160
            TabIndex        =   63
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label10 
            BackColor       =   &H00800000&
            Caption         =   "List of Item/s Rented "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackColor       =   &H00800000&
            Caption         =   "Item 5"
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
            TabIndex        =   51
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label7 
            BackColor       =   &H00800000&
            Caption         =   "Item 4"
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
            TabIndex        =   50
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H00800000&
            Caption         =   "Item 3"
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
            TabIndex        =   49
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label4 
            BackColor       =   &H00800000&
            Caption         =   "Item 2"
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
            TabIndex        =   48
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            Caption         =   "Item 1"
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
            TabIndex        =   47
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label11 
            BackColor       =   &H00800000&
            Caption         =   "Total Amount"
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
            Left            =   5760
            TabIndex        =   40
            Top             =   1320
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1335
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   16776960
         HeadLines       =   1
         RowHeight       =   20
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "membershipid"
            Caption         =   "Borrowers ID#"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
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
            DataField       =   "datebor"
            Caption         =   "Date Borrowed"
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
            DataField       =   "duedate"
            Caption         =   "Due Date"
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
            DataField       =   "amount"
            Caption         =   "Amount"
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
               Locked          =   -1  'True
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1679.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt13 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Text13"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt12 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt11 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt10 
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txt9 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Text9"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txt8 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Text8"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txt7 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Text7"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txt6 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txt5 
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt4 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt3 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txt2 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txt1 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdCharge 
         BackColor       =   &H00FFFF00&
         Caption         =   "C&harge"
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
         Left            =   4680
         Picture         =   "frmRent.frx":1918
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdCCharge 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00800000&
      Height          =   1335
      Left            =   4920
      ScaleHeight     =   1275
      ScaleWidth      =   3915
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3975
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
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
         ItemData        =   "frmRent.frx":21E2
         Left            =   7560
         List            =   "frmRent.frx":21E4
         TabIndex        =   21
         Top             =   120
         Width           =   3015
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
         Height          =   1020
         ItemData        =   "frmRent.frx":21E6
         Left            =   2400
         List            =   "frmRent.frx":21E8
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optItemID 
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
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
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
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   60
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   8715
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdCCancel 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCRent 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCNew 
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00FFFF00&
         Caption         =   "&New"
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
         Left            =   5880
         Picture         =   "frmRent.frx":21EA
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFF00&
         Caption         =   "Canc&el"
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
         Left            =   7800
         Picture         =   "frmRent.frx":2AB4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRent 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Rent"
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
         Left            =   6840
         Picture         =   "frmRent.frx":337E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
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
         Left            =   1800
         TabIndex        =   17
         Top             =   120
         Width           =   3975
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
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Borrowers Name"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Movie Title"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   4635
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4695
      Begin VB.ListBox lstMemID 
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
         Height          =   1020
         ItemData        =   "frmRent.frx":37C0
         Left            =   3120
         List            =   "frmRent.frx":37C2
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtMemID 
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
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtautonum 
         Height          =   285
         Left            =   3240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtrrtransno 
         Height          =   285
         Left            =   3120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optMemID 
         BackColor       =   &H00800000&
         Caption         =   "&Borrowers ID#"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modev As Boolean
Dim item_control_no

Private Sub cmdCalculator_Click()
  On Error GoTo Mali
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Mali:
    MsgBox "Calculator is not installed in your computer.", vbExclamation, App.Title
End Sub

Private Sub cmdCancel_Click()
  txtautonum.Text = Val(txtautonum.Text) - 1
  txtautonum.Text = Format(txtautonum.Text, "00000000")
  txtrrtransno.Text = ""
  If Text1.Text = "" Then
    cmdButton "0 1 1 1 0 1 0 0 0 1 1 0"
    Call clear
  Else
    cmdButton "0 1 1 1 0 1 0 0 0 1 0 1"
    txtItemID.Text = ""
    txtstatus.Text = ""
    lblMovieTitle.Caption = ""
  End If
  Picture1.Enabled = False
  Picture3.Enabled = False
End Sub

Private Sub cmdCharge_Click()
  Picture1.Enabled = False
  Picture3.Enabled = False
  Dim count
  txtTitl.Text = adorr!Title
  txtA.Text = adorr!amount
  txtM.Text = adorr!rrtransno
  txtAmount.Text = Val(txtAmount.Text) + Val(txtA.Text)
  txtAmount.Text = Format(txtAmount.Text, "####.00")
    If Text1.Text = "" Then
      Text1.Text = txtTitl.Text
      Text6.Text = txtM.Text
      count = 1
    ElseIf Text1.Text <> "" And Text2.Text = "" Then
      Text2.Text = txtTitl.Text
      Text7.Text = txtM.Text
      count = 2
    ElseIf Text1.Text <> "" And Text2.Text <> "" And Text3.Text = "" Then
      Text3.Text = txtTitl.Text
      Text8.Text = txtM.Text
      count = 3
    ElseIf Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text = "" Then
      Text4.Text = txtTitl.Text
      Text9.Text = txtM.Text
      count = 4
    Else
      Text5.Text = txtTitl.Text
      Text10.Text = txtM.Text
      count = 5
    End If
    txtItemCount.Text = count
    cmdButton "0 1 1 1 1 1 0 0 0 0 0 1"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim response As Integer
  If cmdCOR.Visible = False Then
    MsgBox "Printing the Official Receipt is" & vbCrLf & "needed to complete the transaction", vbInformation, "Information"
    Cancel = 1
  ElseIf cmdCNew.Visible = True Then
    MsgBox "Click the Cancel Button" & vbCrLf & "to cancel transaction", vbInformation, "Information"
    Cancel = 1
  Else
    Unload Me
    Load frmMain
    frmMain.Show
    frmMain.Enabled = True
  End If
  Call Due
End Sub


Private Sub cmdNew_Click()
  If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
    MsgBox "Sorry maximum of 5 Items can be rented", vbCritical, "Information"
    Exit Sub
  Else
    If Text1.Text <> "" Then
      Picture1.Enabled = False
      Picture3.Enabled = True
      txtItemID.Text = ""
      txtstatus.Text = ""
      lblMovieTitle.Caption = ""
      optItemID.value = True
      txtItemID.SetFocus
      GoTo continue:
    Else
      Picture1.Enabled = True
      Picture3.Enabled = True
      optMemID.value = True
      txtMemID.SetFocus
      Call clear
continue:
      rs_act_autonum
      txtautonum.Text = adoautonum!transno
      txtautonum.Text = Val(txtautonum.Text) + 1
      txtautonum.Text = Format(txtautonum.Text, "00000000")
      txtrrtransno.Text = txtautonum.Text
    End If
      cmdButton "1 0 0 1 1 0 1 1 0 0 1 0"
  End If
End Sub

Private Sub cmdPrintOR_Click()
  frmRent.Enabled = False
  Load frmCashier
  frmCashier.Show
  cmdButton "0 1 1 1 0 1 0 0 0 1 1 0"
End Sub

Private Sub Form_Load()
  rentreturnconnect
  Set DataGrid1.DataSource = adorr
  Call clear
  cmdButton "0 1 1 1 0 1 0 0 0 1 1 0"
  Picture1.Enabled = False
  Picture3.Enabled = False
End Sub

Private Sub cmdButton(ByVal strBinary As String)
  Dim strAry() As String
  strAry = Split(strBinary)
  cmdCNew.Visible = strAry(0)
  cmdCRent.Visible = strAry(1)
  cmdCCancel.Visible = strAry(2)
  cmdCCharge.Visible = strAry(3)
  cmdCClose.Visible = strAry(4)
  cmdNew.Enabled = strAry(5)
  cmdRent.Enabled = strAry(6)
  cmdCancel.Enabled = strAry(7)
  cmdCharge.Enabled = strAry(8)
  cmdClose.Enabled = strAry(9)
  cmdCOR.Visible = strAry(10)
  cmdPrintOR.Enabled = strAry(11)
End Sub

Private Sub clear()
  With frmRent
    .txtMemID.Text = ""
    .lblMemName.Caption = ""
    .txtItemID.Text = ""
    .txtstatus.Text = ""
    .lblMovieTitle.Caption = ""
   End With
End Sub

Private Sub txtMemID_Change()
  MemID (txtMemID.Text)
End Sub

Private Sub txtItemID_Change()
  Item (txtItemID.Text)
End Sub

Private Sub lstMemID_Click()
  MemID (lstMemID.List(lstMemID.ListIndex))
End Sub

Private Sub lstItemID_Click()
  Item (lstItemID.List(lstItemID.ListIndex))
End Sub

Private Sub txtMemID_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call MID(Me.txtMemID.Text)
  End If
End Sub

Private Sub txtItemID_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call IID(Me.txtItemID.Text)
  End If
End Sub

Private Sub optItemID_Click()
  optMemID.value = False
  txtItemID.SetFocus
  txtItemID.Locked = False
  lstItemID.Enabled = True
  txtMemID.Locked = True
  lstMemID.Enabled = False
End Sub

Private Sub optMemID_Click()
  optItemID.value = False
  txtMemID.SetFocus
  txtMemID.Locked = False
  lstMemID.Enabled = True
  txtItemID.Locked = True
  lstItemID.Enabled = False
End Sub

Private Sub cmdRent_Click()
  If lblMemName.Caption = "" Or lblMovieTitle.Caption = "" Then
    MsgBox "Missing Data! Either Membership ID# or" & vbCrLf & "Item ID# textfield does not have a value", vbInformation, "Information"
    Exit Sub
  Else
    Call item_control
      If modev = False Then
        Dim res As VbMsgBoxResult
        res = MsgBox("Borrowers Name: " & lblMemName.Caption & vbCrLf & vbCrLf & "Movie Title: " & lblMovieTitle.Caption & vbCrLf & vbCrLf & "No. of Rented Items: " & item_control_no & vbCrLf & vbCrLf & "Do you want to rent this Item?", vbYesNo + vbQuestion, "Confirmation")
        If res = vbYes Then
          Picture1.Enabled = False
          Picture3.Enabled = False
          txtstatus.Text = "OUT"
          txt13.Text = "UnReturned"
          adorr.AddNew
          WriteDataFromControls
          adorr.Update
          adoautonum.UpdateBatch adAffectCurrent
          adorent.UpdateBatch adAffectCurrent
          Set DataGrid1.DataSource = adorr
          cmdButton "1 1 1 0 1 0 0 0 1 0 1 0"
          cmdCharge_Click
        Else
          Exit Sub
        End If
      ElseIf modev = True Then
        MsgBox "Sorry! Borrower  ''" & lblMemName.Caption & "''  has already Rented 5 Items", vbCritical, "Information"
      End If
  End If
End Sub

Private Sub item_control()
  On Error Resume Next
  Dim rs As New ADODB.Recordset
  rs.Open "Select * From rentreturn Where membershipid = '" & txt5.Text & "' And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount < 5 Then
      modev = False
      item_control_no = rs.RecordCount
      Exit Sub
    Else
      modev = True
    End If
      Set rs = Nothing
End Sub

Private Sub WriteDataFromControls()
  On Error Resume Next
  adorent!Status = txtstatus.Text
  adoautonum!transno = txtautonum.Text
  adorr!rrtransno = txtrrtransno.Text
  adorr!itemidnumber = txt1.Text
  adorr!Title = txt2.Text
  adorr!Format = txt3.Text
  adorr!amount = txt4.Text
  adorr!membershipid = txt5.Text
  adorr!lastname = txt6.Text
  adorr!firstname = txt7.Text
  adorr!datebor = txt8.Text
  adorr!duedate = txt9.Text
  adorr!DateReturned = txt10.Text
  adorr!noofdayspenalty = txt11.Text
  adorr!totalpenaltyamount = txt12.Text
  adorr!rentreturnstatus = txt13.Text
End Sub

Private Sub txtMemID_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtItemID_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
  End If
End Sub
