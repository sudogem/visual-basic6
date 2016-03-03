VERSION 5.00
Begin VB.Form frmSemClass 
   Caption         =   "form for SemClass Table.....wala pa nahuman"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   540
      Left            =   4155
      TabIndex        =   9
      Top             =   2340
      Width           =   1065
   End
   Begin VB.ComboBox cboYearLevel 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   780
      Width           =   2295
   End
   Begin VB.ComboBox cboSemester 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1740
      Width           =   2295
   End
   Begin VB.ComboBox cboClassName 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   300
      Width           =   2295
   End
   Begin VB.ComboBox cboSchoolYear 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddSY 
      Caption         =   "Add &SchoolYear"
      Height          =   540
      Left            =   3855
      TabIndex        =   4
      Top             =   1200
      Width           =   1350
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   540
      Left            =   2070
      TabIndex        =   3
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   540
      Left            =   180
      TabIndex        =   2
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   540
      Left            =   3015
      TabIndex        =   1
      Top             =   2340
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   540
      Left            =   1125
      TabIndex        =   0
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Class Name :"
      Height          =   255
      Left            =   435
      TabIndex        =   13
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "School Year :"
      Height          =   255
      Left            =   435
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Semester :"
      Height          =   255
      Left            =   435
      TabIndex        =   11
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Year Level :"
      Height          =   255
      Left            =   435
      TabIndex        =   10
      Top             =   840
      Width           =   915
   End
End
Attribute VB_Name = "frmSemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public getCourseID As Long

Private Sub cmdAdd_Click()
  UnLockedAllTextBoxes
  DisableAddEdit
End Sub

Private Sub cmdAddSY_Click()
 frmAddSchoolYear.Show
End Sub
Private Sub cmdEdit_Click()
 DisableAddEdit
 UnLockedAllTextBoxes
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
   AddSemClass cboSchoolYear.Text, cboYearLevel.Text, cboSemester.Text, getCourseID
   DisableSaveCancel
   LockedAllTextBoxes
End Sub
Private Sub Form_Load()
 GetClass cboClassName
 GetYearLevel cboYearLevel
 GetListSemester cboSemester
 GetSchoolYear cboSchoolYear
 getCourseID = 100
End Sub
Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdAdd.Enabled = True
 cmdEdit.Enabled = True
End Sub
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 cmdAdd.Enabled = False
 cmdEdit.Enabled = False
End Sub
Private Sub LockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Enabled = False
  Next ctrl
End Sub
Private Sub UnLockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then ctrl.Enabled = True
  Next ctrl
End Sub
