VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListofSchedule 
   BackColor       =   &H0080C0FF&
   Caption         =   "List of Schedules"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Print Schedule"
      Height          =   375
      Left            =   6540
      TabIndex        =   2
      Top             =   5520
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid grdListOfSched 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Class Schedule of SY 02 Curriculum 1999 First Semester"
      Height          =   420
      Left            =   1575
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmListofSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 PopulateFlexGrid "SELECT * FROM tblSchedule", grdListOfSched, True
End Sub
