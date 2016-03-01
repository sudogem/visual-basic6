VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAddSchoolYear 
   BackColor       =   &H00FFFFC0&
   Caption         =   "School Year"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   585
      Left            =   2415
      TabIndex        =   6
      Top             =   630
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   585
      Left            =   1365
      TabIndex        =   4
      Top             =   630
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   600
      Left            =   3465
      TabIndex        =   3
      Top             =   630
      Width           =   1020
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   585
      Left            =   405
      TabIndex        =   2
      Top             =   630
      Width           =   915
   End
   Begin VB.TextBox txtSchoolYear 
      Height          =   345
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   0
      Top             =   195
      Width           =   2955
   End
   Begin MSFlexGridLib.MSFlexGrid grdSchoolYear 
      Height          =   1650
      Left            =   4620
      TabIndex        =   5
      Top             =   105
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2910
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      Appearance      =   0
      FormatString    =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "School Year :"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   255
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddSchoolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim state As AddEditState
Public getSchoolYearId As String

Private Sub cmdAdd_Click()
 UnLockedAllTextBoxes
 DisableAddEdit
 ClearTextBoxes
 state = AddState
End Sub

Private Sub cmdDelete_Click()
Dim xxx As New SubjectClass.SchoolYear
Dim temp As String
grdSchoolYear.Col = 0
temp = Trim(grdSchoolYear.Text)
getSchoolYearId = IIf(grdSchoolYear.Text <> "", temp, "!")
If getSchoolYearId <> "!" Then
  If MsgBox("Are you sure you want to delete this SY " & getSchoolYearId & "? ", vbYesNo + vbCritical, "Delete record") = vbYes Then
     If Not xxx.DeleteSchoolYear(getSchoolYearId) = False Then
        MsgBox "The record was successfully deleted.", vbOKOnly + vbInformation, "Delete Schoolyear"
        UnPopulateFlexGrid grdSchoolYear
        PopulateFlexGrid "SELECT * FROM tblSchoolyear", grdSchoolYear, True
     Else
        MsgBox "Dll Object Error: DeleteSchoolYearError!!!", vbCritical + vbOKOnly, "Delete Error"
     End If
  End If
End If
Set xxx = Nothing
End Sub

Private Sub cmdEdit_Click()
 DisableAddEdit
 UnLockedAllTextBoxes
 state = EditState
 cmdSave.Enabled = True
 grdSchoolYear.Col = 0
 frmAddSchoolYear.txtSchoolYear = grdSchoolYear.Text
 getSchoolYearId = frmAddSchoolYear.txtSchoolYear.Text
 MsgBox "SY=" & getSchoolYearId
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
Dim xxx As New SubjectClass.SchoolYear
 
 If state = AddState Then
   If MissingEntries = True Then
       MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
   Else
       If xxx.AddSchoolYear(txtSchoolYear.Text) = False Then
          MsgBox "Please input the correct Schoolyear format.", vbOKOnly + vbCritical, "Error: Invalid Schoolyear Format."
       Else
          MsgBox "The record was successfully added.", vbInformation + vbOKOnly, "Add Schoolyear"
          DisableSaveCancel
          LockedAllTextBoxes
          UnPopulateFlexGrid grdSchoolYear
          PopulateFlexGrid "SELECT * FROM tblSchoolyear", grdSchoolYear, True
          Set xxx = Nothing
       End If
   End If
 ElseIf state = EditState Then
   If MissingEntries = True Then
       MsgBox "Please input the correct Schoolyear format.", vbOKOnly + vbCritical, "Error: Invalid Schoolyear Format."
   Else
       If xxx.AddSchoolYear(txtSchoolYear.Text) = False Then
          MsgBox "Please input the correct Schoolyear format."
       Else
          MsgBox "The record was successfully added.", vbInformation + vbOKOnly, "Add Schoolyear"
          DisableSaveCancel
          LockedAllTextBoxes
          UnPopulateFlexGrid grdSchoolYear
          PopulateFlexGrid "SELECT * FROM tblSchoolyear", grdSchoolYear, True
          Set xxx = Nothing
       End If
   End If
 End If
End Sub

Private Sub Form_Load()
  LockedAllTextBoxes
  PopulateFlexGrid "SELECT * FROM tblSchoolYear", grdSchoolYear, True
End Sub
' clear all textboxes......
Private Sub ClearTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Text = Empty
  Next ctrl
End Sub

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
' check the missing input.....
Private Function MissingEntries() As Boolean
 If Trim(txtSchoolYear.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function
Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdDelete.Enabled = False
 cmdAdd.Enabled = True
End Sub
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 cmdDelete.Enabled = False
End Sub

Private Sub grdSchoolYear_Click()
 grdSchoolYear.Cols = 1
 grdSchoolYear.RowSel = grdSchoolYear.Row
End Sub
