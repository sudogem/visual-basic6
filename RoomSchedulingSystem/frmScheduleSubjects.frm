VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScheduleSubjects 
   Caption         =   "Form2"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   2940
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   435
      Left            =   6360
      TabIndex        =   3
      Top             =   4500
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pending Subjects :"
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
      Begin VB.ListBox lstAllSubjects 
         Height          =   2790
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   2355
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1260
      TabIndex        =   6
      Top             =   300
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "ScheduleSubjects :"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Class Name :"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmScheduleSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 GetAllSubjects lstAllSubjects
End Sub

Private Sub List1_Click()

End Sub
