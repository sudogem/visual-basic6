VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCurriculum 
   Caption         =   "Set Curriculum"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   540
      Left            =   210
      TabIndex        =   9
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   540
      Left            =   2520
      TabIndex        =   8
      Top             =   2100
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1800
      Left            =   4200
      TabIndex        =   7
      Top             =   105
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   3175
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtYearImplemented 
      Height          =   285
      Left            =   1605
      TabIndex        =   4
      Top             =   135
      Width           =   2475
   End
   Begin VB.TextBox txtDescription 
      Height          =   1320
      Left            =   1605
      TabIndex        =   3
      Top             =   615
      Width           =   2475
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   540
      Left            =   1365
      TabIndex        =   2
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   540
      Left            =   3675
      TabIndex        =   1
      Top             =   2100
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   645
      Left            =   5565
      TabIndex        =   0
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Year Implemented :"
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Description :"
      Height          =   315
      Left            =   525
      TabIndex        =   5
      Top             =   735
      Width           =   975
   End
End
Attribute VB_Name = "frmCurriculum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim state As AddEditState
Public getCurriculumId As Long

Private Sub Command1_Click()
 Dim RS As New ADODB.Recordset
 Dim cn As New ADODB.Connection
 
End Sub

Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub cmdAdd_Click()
  UnLockedAllTextBoxes
  ClearTextBoxes
  DisableAddEdit
  state = AddState
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
 If MissingEntries = True Then
   MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
 Else
   'MsgBox getCourseID
   AddCurriculum 100, txtYearImplemented.Text, txtDescription.Text
   'AddCurriculum
   DisableSaveCancel
   LockedAllTextBoxes
 End If
End Sub
' check the missing input..
Private Function MissingEntries() As Boolean
 If Trim(txtYearImplemented.Text) = "" Or Trim(txtDescription.Text) = "" Then
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
 'cmdAdd.Enabled = True
 'cmdEdit.Enabled = True
End Sub
'enable add ,edit button...
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 'cmdAdd.Enabled = False
 'cmdEdit.Enabled = False
End Sub

