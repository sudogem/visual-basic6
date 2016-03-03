VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSubjectInfo 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Subject Information "
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   540
      Left            =   3060
      TabIndex        =   11
      Tag             =   "add"
      Top             =   2565
      Width           =   855
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1380
      Width           =   3015
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   720
      Left            =   4980
      TabIndex        =   5
      Top             =   2460
      Width           =   1155
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   540
      Left            =   2115
      TabIndex        =   4
      Top             =   2565
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   540
      Left            =   300
      TabIndex        =   3
      Tag             =   "add"
      Top             =   2565
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   540
      Left            =   4020
      TabIndex        =   2
      Top             =   2565
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   540
      Left            =   1200
      TabIndex        =   1
      Top             =   2565
      Width           =   855
   End
   Begin VB.TextBox txtUnits 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   780
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid grdSubject 
      Height          =   2115
      Left            =   4680
      TabIndex        =   12
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3731
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Subject Name:"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   420
      TabIndex        =   9
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Units:"
      Height          =   315
      Left            =   420
      TabIndex        =   8
      Top             =   900
      Width           =   615
   End
End
Attribute VB_Name = "frmSubjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim state As AddEditState
Public getSubjectID As Long

Private Sub cmdAdd_Click()
 state = AddState
 UnLockedAllTextBoxes
 ClearTextBoxes
 DisableAddEditDelete
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim xxx As New SubjectClass.clsSubject
Dim temp As Long

grdSubject.Col = 0
temp = Val(Trim(grdSubject.Text))
getSubjectID = IIf(grdSubject.Text <> "", temp, -1)
If getSubjectID <> -1 Then
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbCritical, "Delete Record") = vbYes Then
       If xxx.DeleteSubject(getSubjectID) = True Then
          MsgBox "Record was succesfully delete..."
          UnPopulateFlexGrid grdSubject
          PopulateFlexGrid "SELECT * FROM tblSubject", grdSubject, True
       Else
          MsgBox "Dll Object Error: DeleteModuleError!!!", vbCritical + vbOKOnly, "Delete Error"
       End If
     End If
End If
Set xxx = Nothing
End Sub

Private Sub cmdEdit_Click()
state = EditState
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdCancel.Enabled = True
cmdSave.Enabled = True
UnLockedAllTextBoxes
'MsgBox grdSubject.RowSel
grdSubject.Col = 0
frmSubjectInfo.getSubjectID = grdSubject.Text
MsgBox "Subject id=" & frmSubjectInfo.getSubjectID
grdSubject.Col = 1
frmSubjectInfo.txtSubject.Text = grdSubject.Text
grdSubject.Col = 2
frmSubjectInfo.txtDescription.Text = grdSubject.Text
grdSubject.Col = 3
frmSubjectInfo.txtUnits.Text = grdSubject.Text

End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub
Private Sub cmdSave_Click()
Dim xxx As New SubjectClass.clsSubject

If state = AddState Then
  If MissingEntries = True Then
      MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
  Else
    If CheckIfDigitFound(txtDescription, txtDescription.Text) = True Then
       MsgBox "Please input the correct description for subject.", vbCritical + vbOKOnly, "Error: Invalid Subject Description!"
    ElseIf IsNumeric(txtUnits.Text) = False Then
       MsgBox "Please input correct number of units for subject.", vbCritical + vbOKOnly, "Error: Invalid Subject Units!"
    Else
       If xxx.AddSubject(txtSubject.Text, txtDescription.Text, Val(txtUnits.Text)) = False Then
          MsgBox "DLL Error in Subject Class."
       Else
          MsgBox "The record was successfully added.", vbInformation + vbOKOnly, "Add Subject"
          DisableSaveCancel
          LockedAllTextBoxes
          UnPopulateFlexGrid grdSubject
          PopulateFlexGrid "SELECT * FROM tblSubject", grdSubject, True
          Set xxx = Nothing
       End If
    End If
  End If
ElseIf state = EditState Then
   If MissingEntries = True Then
        MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
   Else
       If CheckIfDigitFound(txtDescription, txtDescription.Text) = True Then
           MsgBox "Please input the correct description for subject.", vbCritical + vbOKOnly, "Error: Invalid Subject Description!"
       ElseIf IsNumeric(txtUnits.Text) = False Then
           MsgBox "Please input correct number of units for subject.", vbCritical + vbOKOnly, "Error: Invalid Subject Units!"
       Else
          If xxx.EditSubject(getSubjectID, txtSubject.Text, txtDescription.Text, txtUnits.Text) = False Then
            MsgBox "DLL Error in Subject Class."
          Else
            MsgBox "The record was successfully edited.", vbInformation + vbOKOnly, "Edit Subject"
            DisableSaveCancel
            LockedAllTextBoxes
            UnPopulateFlexGrid grdSubject
            PopulateFlexGrid "SELECT * FROM tblSubject", grdSubject, True
            Set xxx = Nothing
          End If
       End If
   End If
End If
End Sub

Private Sub Form_Load()
 LockedAllTextBoxes
 DisableSaveCancel
 PopulateFlexGrid "SELECT * FROM tblSubject", grdSubject, True
End Sub

' clear all textboxes......
Private Sub ClearTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Text = Empty
  Next ctrl
End Sub
' unlocked all textboxes....
Private Sub UnLockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then ctrl.Enabled = True
  Next ctrl
End Sub
' locked all textbaks
Private Sub LockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Enabled = False
  Next ctrl
End Sub
' check the missing input.....
Private Function MissingEntries() As Boolean
 If Trim(txtSubject.Text) = "" Or Trim(txtDescription.Text) = "" Or Trim(txtUnits.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function

Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 cmdAdd.Enabled = True
 cmdEdit.Enabled = True
 cmdDelete.Enabled = True
End Sub
Private Sub DisableAddEditDelete()
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 cmdAdd.Enabled = False
 cmdEdit.Enabled = False
 cmdDelete.Enabled = False
End Sub
Private Sub grdSubject_Click()
 grdSubject.RowSel = grdSubject.Row
End Sub
