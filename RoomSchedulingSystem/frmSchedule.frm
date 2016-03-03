VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchedule 
   Caption         =   "Schedule Form"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   615
      Left            =   8520
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cboSubjectModule 
      Height          =   315
      Left            =   1260
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   2520
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   4920
      TabIndex        =   18
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4048
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   600
      Left            =   7200
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cboRoom 
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Days"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   5775
      Begin VB.CheckBox Check1 
         Caption         =   "Sat"
         Height          =   195
         Index           =   6
         Left            =   5100
         TabIndex        =   17
         Tag             =   "Sat"
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fri"
         Height          =   195
         Index           =   5
         Left            =   4380
         TabIndex        =   16
         Tag             =   "Fri"
         Top             =   360
         Width           =   555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Thurs"
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   15
         Tag             =   "Thurs"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Wed"
         Height          =   195
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Tag             =   "Wed"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tue"
         Height          =   195
         Index           =   2
         Left            =   1860
         TabIndex        =   13
         Tag             =   "Tue"
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mon"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   12
         Tag             =   "Mon"
         Top             =   360
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sun"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Tag             =   "Sun"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox cboEndTime 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Text            =   "Combo4"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.ComboBox cboStartTime 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Text            =   "Combo3"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ComboBox cboInstructor 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Subject Module :"
      Height          =   495
      Left            =   210
      TabIndex        =   20
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "Room :"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "EndTime :"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "StartTime :"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Instructor :"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   915
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
 Unload Me
End Sub

