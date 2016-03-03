VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRoomInfo 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Room Information"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdRoom 
      Height          =   2235
      Left            =   4560
      TabIndex        =   8
      Top             =   105
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3942
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txtRoomName 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   3075
   End
   Begin VB.TextBox txtRoomDescription 
      Height          =   915
      Left            =   1320
      TabIndex        =   4
      Top             =   660
      Width           =   3075
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   690
      Left            =   3180
      TabIndex        =   3
      Top             =   1740
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   480
      Left            =   1140
      TabIndex        =   2
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   480
      Left            =   105
      TabIndex        =   1
      Top             =   1845
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   480
      Left            =   2100
      TabIndex        =   0
      Top             =   1860
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Name :"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   780
      Width           =   975
   End
End
Attribute VB_Name = "frmRoomInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim state As AddEditState
Public getRoomID As String

Private Sub cmdAdd_Click()
  UnLockedAllTextBoxes
  ClearTextBoxes
  DisableAddEdit
  state = AddState
End Sub
' clear all textboxes......
Private Sub ClearTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Text = Empty
  Next ctrl
End Sub
' locked all textboes
Private Sub LockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Enabled = False
  Next ctrl
End Sub
' unlocked all textboxes
Private Sub UnLockedAllTextBoxes()
Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then ctrl.Enabled = True
  Next ctrl
End Sub
' check the missing input..
Private Function MissingEntries() As Boolean
 If Trim(txtRoomName.Text) = "" Or Trim(txtRoomDescription.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function

Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
' cmdCancel.Enabled = False
 cmdAdd.Enabled = True
 cmdEdit.Enabled = True
End Sub

Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 'cmdCancel.Enabled = True
 cmdAdd.Enabled = False
 cmdEdit.Enabled = False
End Sub


Private Sub cmdDelete_Click()
'Dim xxx As New SubjectClass.clsRoom
'Dim id As Long
'id = InputBox("Enter SubjectID to delete?")
'If xxx.DeleteSubject(id) = True Then
'  MsgBox "Succesfully delete.."
'Else
'  MsgBox "Delete Error.."
'End If

'grdRoom.Col = 0
'getRoomID = grdRoom.RowSel
'MsgBox grdRoom.Text
'If xxxDeleteRoom(getRoomID) = True Then
'   MsgBox "Succesfully delete.."
'   UnPopulateFlexGrid grdRoom
'   PopulateFlexGrid "SELECT * FROM tblRoom", grdRoom, True
'Else
'   MsgBox "Delete Error.."
'End If

End Sub


Private Sub cmdEdit_Click()
 DisableAddEdit
 UnLockedAllTextBoxes
 state = EditState
 cmdSave.Enabled = True
 grdRoom.Col = 0
 frmRoomInfo.txtRoomName = grdRoom.Text
 grdRoom.Col = 1
 frmRoomInfo.txtRoomDescription = grdRoom.Text
 getRoomID = frmRoomInfo.txtRoomName
 MsgBox "Room id=" & getRoomID
End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
Dim xxx As New SubjectClass.clsRoom
 
If state = AddState Then
  If MissingEntries = True Then
    MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
  Else
    If xxx.AddRoom(txtRoomName.Text, txtRoomDescription.Text) = False Then
      MsgBox "Please enter the correct values."
    Else
      MsgBox "The record was successfully added."
      DisableSaveCancel
      LockedAllTextBoxes
      UnPopulateFlexGrid grdRoom
      'frmRoomInfo.grdRoom.Clear
      PopulateFlexGrid "SELECT * FROM tblRoom", grdRoom, True
      Set xxx = Nothing
    End If
  End If
ElseIf state = EditState Then
  UnLockedAllTextBoxes
  cmdSave.Enabled = True
  If MissingEntries = True Then
    MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
  Else
    If xxx.EditRoom(getRoomID, txtRoomName.Text, txtRoomDescription.Text) = False Then
      MsgBox "DLL Object Error: RoomClassError."
    Else
      MsgBox "The record was successfully edited."
      DisableSaveCancel
      LockedAllTextBoxes
      UnPopulateFlexGrid grdRoom
      'frmRoomInfo.grdRoom.Clear
      PopulateFlexGrid "SELECT * FROM tblRoom", grdRoom, True
      Set xxx = Nothing
    End If
 End If
End If
 
End Sub

Private Sub Form_Load()
 DisableSaveCancel
 LockedAllTextBoxes
 PopulateFlexGrid "SELECT * FROM tblRoom", grdRoom, True
End Sub

Private Sub grdRoom_Click()
  grdRoom.RowSel = grdRoom.Row
End Sub

