VERSION 5.00
Begin VB.Form frmClassToSchedule 
   BackColor       =   &H00C0FFFF&
   Caption         =   "form for ClassToSchedule ..."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddCurriculum 
      Caption         =   "Add New Curriculum"
      Height          =   225
      Left            =   5040
      TabIndex        =   21
      Top             =   1575
      Width           =   1905
   End
   Begin VB.ComboBox cboCurriculum 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1830
      Width           =   1935
   End
   Begin VB.ComboBox cboSchoolYear 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   180
      Width           =   1935
   End
   Begin VB.ComboBox cboYearLevel 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   660
      Width           =   1935
   End
   Begin VB.ComboBox cboSemester 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1140
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   540
      Left            =   2145
      TabIndex        =   12
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   540
      Left            =   405
      TabIndex        =   11
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   540
      Left            =   3045
      TabIndex        =   10
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   540
      Left            =   1305
      TabIndex        =   9
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   600
      Left            =   4095
      TabIndex        =   8
      Top             =   2280
      Width           =   1020
   End
   Begin VB.TextBox txtStartDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   2235
   End
   Begin VB.ComboBox cboCourseType 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1620
      Width           =   2295
   End
   Begin VB.TextBox txtEndDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1140
      Width           =   2235
   End
   Begin VB.TextBox txtClass 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   180
      Width           =   2235
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Curriculum Year Implemented :"
      Height          =   435
      Left            =   3885
      TabIndex        =   19
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level :"
      Height          =   195
      Left            =   3900
      TabIndex        =   15
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester :"
      Height          =   195
      Left            =   3900
      TabIndex        =   14
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "School Year :"
      Height          =   255
      Left            =   3900
      TabIndex        =   13
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Course :"
      Height          =   255
      Left            =   450
      TabIndex        =   7
      Top             =   1620
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date :"
      Height          =   255
      Left            =   450
      TabIndex        =   6
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date :"
      Height          =   255
      Left            =   450
      TabIndex        =   5
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name :"
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmClassToSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCourse_Click()
  
End Sub

Private Sub cboCourseType_Click()
  If cboCourseType.Text = "Semestral" Then
    Label5.Visible = True
    Label8.Visible = True
    Label7.Visible = True
    Label6.Visible = True
    cboSchoolYear.Visible = True
    cboYearLevel.Visible = True
    cboSemester.Visible = True
    cboCurriculum.Visible = True
  Else
    Label5.Visible = False
    Label8.Visible = False
    Label7.Visible = False
    Label6.Visible = False
    cboSchoolYear.Visible = False
    cboYearLevel.Visible = False
    cboSemester.Visible = False
    cboCurriculum.Visible = False
  End If
  
End Sub

Private Sub cmdAdd_Click()
  UnLockedAllTextBoxes
  DisableAddEdit
  
End Sub

Private Sub cmdAddCurriculum_Click()
 frmCurriculum.Show
End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
Dim xxx As SubjectClass.clsClassToSched

 If MissingEntries = True Then
   MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
 Else
   'AddClassToSchedule txtClass, getCourseID, txtStartDate.Text, txtEndDate.Text
   'AddCurriculum getCourseID, cboCurriculum.Text
   'AddSemClass cboSchoolYear.Text, cboYearLevel.Text, cboSemester.Text, getCurriculumID
   'DisableSaveCancel
   'LockedAllTextBoxes
 End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  LockedAllTextBoxes
  GetCourseType cboCourseType
  'GetListCourse cboCourse
  GetSchoolYear cboSchoolYear
  GetYearLevel cboYearLevel
  GetListSemester cboSemester
  GetCurriculum cboCurriculum
End Sub
' check the missing input..
Private Function MissingEntries() As Boolean
 If Trim(txtClass.Text) = "" Or Trim(cboCourseType.Text) = "" Or Trim(txtStartDate.Text) = "" Or Trim(txtEndDate.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function
' locked all textboes......
Private Sub LockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Enabled = False
  Next ctrl
End Sub
' unlocked all textboxes....
Private Sub UnLockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then ctrl.Enabled = True
  Next ctrl
End Sub
' disable save ,cancel button...
Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdAdd.Enabled = True
 cmdEdit.Enabled = True
End Sub
'enable add ,edit button...
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 cmdAdd.Enabled = False
 cmdEdit.Enabled = False
End Sub

