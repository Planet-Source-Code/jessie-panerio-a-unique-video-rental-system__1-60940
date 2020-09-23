VERSION 5.00
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   -360
      ScaleHeight     =   3315
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   -120
      Width           =   5175
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hello"
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

