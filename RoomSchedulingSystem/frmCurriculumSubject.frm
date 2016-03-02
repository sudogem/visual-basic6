VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCurriculumSubject 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form for Curriculum Subject"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form2"
   ScaleHeight     =   6210
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstAllSubjects 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   210
      TabIndex        =   1
      Top             =   825
      Width           =   3915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   4380
      TabIndex        =   0
      Top             =   840
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   9022
      _Version        =   393216
      Tabs            =   12
      Tab             =   6
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCurriculumSubject.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmCurriculumSubject.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmCurriculumSubject.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List2(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmCurriculumSubject.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "List2(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmCurriculumSubject.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "List2(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmCurriculumSubject.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "List2(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "frmCurriculumSubject.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "List2(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmCurriculumSubject.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "List2(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "frmCurriculumSubject.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "List2(8)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "frmCurriculumSubject.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "List2(9)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Tab 10"
      TabPicture(10)  =   "frmCurriculumSubject.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "List2(10)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Tab 11"
      TabPicture(11)  =   "frmCurriculumSubject.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "List2(11)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).ControlCount=   1
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   11
         Left            =   -74820
         TabIndex        =   14
         Top             =   3600
         Width           =   3555
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Index           =   10
         Left            =   -74760
         TabIndex        =   13
         Top             =   3600
         Width           =   3550
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Index           =   9
         Left            =   -74760
         TabIndex        =   12
         Top             =   3600
         Width           =   3465
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Index           =   8
         Left            =   -74820
         TabIndex        =   11
         Top             =   3600
         Width           =   3550
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   7
         Left            =   -74820
         TabIndex        =   10
         Top             =   3600
         Width           =   3645
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   6
         Left            =   180
         TabIndex        =   9
         Top             =   3600
         Width           =   3525
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   5
         Left            =   -74760
         TabIndex        =   8
         Top             =   3600
         Width           =   3465
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   4
         Left            =   -74820
         TabIndex        =   7
         Top             =   3600
         Width           =   3585
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   3
         Left            =   -74760
         TabIndex        =   6
         Top             =   3660
         Width           =   3525
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   2
         Left            =   -74820
         TabIndex        =   5
         Top             =   3600
         Width           =   3585
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   1
         Left            =   -74820
         TabIndex        =   4
         Top             =   3660
         Width           =   3585
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Index           =   0
         Left            =   -74820
         TabIndex        =   3
         Top             =   3660
         Width           =   3550
      End
   End
   Begin VB.Label Label1 
      Caption         =   "List of all Subjects for the yearlevel from that sem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1740
      TabIndex        =   15
      Top             =   180
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "List of subjects :"
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   540
      Width           =   1380
   End
End
Attribute VB_Name = "frmCurriculumSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

End Sub

Private Sub cmdSave_Click()
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

 ' AddCurriculumSubject getCurriculumID, cboSubject.Text, cboYearLevel.Text, cboSemester.Text
End Sub
Private Sub Form_Load()
 GetAllSubjects lstAllSubjects
 ChangeTabCap
 
' GetYearLevel cboYearLevel
' GetListSemester cboSemester
End Sub

Private Sub ChangeTabCap()
  SSTab1.TabCaption(1) = "sdfsdf"
End Sub

Private Sub HideTabs()

End Sub
