VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmModule 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Add New Module"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   540
      Left            =   2880
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid grdModule 
      Height          =   2175
      Left            =   5220
      TabIndex        =   8
      Top             =   180
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   660
      Left            =   3900
      TabIndex        =   5
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   540
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtModDescription 
      Height          =   915
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox txtModule 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   180
      Width           =   3615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   540
      Left            =   180
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   540
      Left            =   1980
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Module Name:"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "frmModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim state As AddEditState
Public getModuleID As Long

Private Sub cmdAdd_Click()
  UnLockedAllTextBoxes
  ClearTextBoxes
  DisableAddEdit
  state = AddState
End Sub
Private Sub cmdDelete_Click()
Dim xxx As New SubjectClass.clsModule
Dim temp As Long
grdModule.Col = 0
temp = Val(Trim(grdModule.Text))
getModuleID = IIf(grdModule.Text <> "", temp, -1)
If getModuleID <> -1 Then
  If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbCritical, "Delete record") = vbYes Then
     If Not xxx.DeleteModule(getModuleID) = False Then
        MsgBox "The record was successfully deleted.", vbInformation + vbOKOnly, "Delete Record"
        UnPopulateFlexGrid grdModule
        PopulateFlexGrid "SELECT * FROM tblModule", grdModule, True
     Else
        MsgBox "Dll Object Error: DeleteModuleError!!!", vbCritical + vbOKOnly, "Delete Error"
     End If
  End If
End If
Set xxx = Nothing
End Sub

Private Sub cmdEdit_Click()
 cmdSave.Enabled = True
 cmdEdit.Enabled = False
 cmdAdd.Enabled = False
 cmdDelete.Enabled = False
 UnLockedAllTextBoxes
 state = EditState
 grdModule.Col = 0
 frmModule.getModuleID = grdModule.Text
 grdModule.Col = 1
 frmModule.txtModule.Text = grdModule.Text
 grdModule.Col = 2
 frmModule.txtModDescription.Text = grdModule.Text
End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub

' save module .......
Private Sub cmdSave_Click()
Dim xxx As New SubjectClass.clsModule
If state = AddState Then
  If MissingEntries = True Then
    MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
  Else
     If CheckIfDigitFound(txtModDescription, txtModDescription.Text) = True Then
        MsgBox "Please input the correct description for this module.", vbCritical + vbOKOnly, "Error: Invalid Module Description!"
     Else
        If xxx.AddModule(txtModule.Text, txtModDescription.Text) = False Then
           MsgBox "DLL Object Error: ModuleClassError."
        Else
           MsgBox "The record was successfully added.", vbInformation + vbOKOnly, "Add Module"
           DisableSaveCancel
           LockedAllTextBoxes
           UnPopulateFlexGrid grdModule
           PopulateFlexGrid "SELECT * FROM tblModule", grdModule, True
           Set xxx = Nothing
        End If
      End If
  End If
ElseIf state = EditState Then
  UnLockedAllTextBoxes
  cmdSave.Enabled = True
  If MissingEntries = True Then
    MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
  Else
     If CheckIfDigitFound(txtModDescription, txtModDescription.Text) = True Then
        MsgBox "Please input the correct description for this module.", vbCritical + vbOKOnly, "Error: Invalid Module Description!"
     Else
        If xxx.EditModule(getModuleID, txtModule.Text, txtModDescription.Text) = False Then
           MsgBox "DLL Object Error:  ModuleClassError."
        Else
           MsgBox "The record was successfully edited.", vbInformation + vbOKOnly, "Edit Module"
           DisableSaveCancel
           LockedAllTextBoxes
           UnPopulateFlexGrid grdModule
           PopulateFlexGrid "SELECT * FROM tblModule", grdModule, True
           Set xxx = Nothing
        End If
     End If
  End If
End If
End Sub
Private Sub Form_Load()
  LockedAllTextBoxes
  cmdSave.Enabled = False
  grdModule.Appearance = flexFlat
  PopulateFlexGrid "SELECT * FROM tblModule", grdModule, True
End Sub
' clear all textboxes......
Private Sub ClearTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Text = Empty
  Next ctrl
End Sub
' locked all textbaks
Private Sub LockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Enabled = False
  Next ctrl
End Sub
' unlocked all textbaks
Private Sub UnLockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then ctrl.Enabled = True
  Next ctrl
End Sub
' check the missing input..
Private Function MissingEntries() As Boolean
 If Trim(txtModule.Text) = "" Or Trim(txtModDescription.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function
Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdAdd.Enabled = True
 cmdEdit.Enabled = True
 cmdDelete.Enabled = True
End Sub
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdAdd.Enabled = False
 cmdEdit.Enabled = False
 cmdDelete.Enabled = False
End Sub

Private Sub grdModule_Click()
 grdModule.RowSel = grdModule.Row
End Sub
