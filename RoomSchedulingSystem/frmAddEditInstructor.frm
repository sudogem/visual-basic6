VERSION 5.00
Begin VB.Form frmAddEditInstructor 
   Caption         =   "Instructor Information"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFN 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   180
      Width           =   3255
   End
   Begin VB.TextBox txtLN 
      Height          =   345
      Left            =   1515
      TabIndex        =   2
      Top             =   1140
      Width           =   3255
   End
   Begin VB.TextBox txtPA 
      Height          =   675
      Left            =   1515
      TabIndex        =   3
      Top             =   1620
      Width           =   3255
   End
   Begin VB.TextBox txtCTel 
      Height          =   315
      Left            =   1515
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtCelNo 
      Height          =   345
      Left            =   1515
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox txtMN 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   660
      Width           =   3255
   End
   Begin VB.TextBox txtCA 
      Height          =   615
      Left            =   1515
      TabIndex        =   5
      Top             =   2940
      Width           =   3255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   465
      Left            =   4140
      TabIndex        =   11
      Top             =   4920
      Width           =   1035
   End
   Begin VB.TextBox txtPTel 
      Height          =   315
      Left            =   1515
      TabIndex        =   4
      Top             =   2460
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   480
      Left            =   1200
      TabIndex        =   9
      Top             =   4860
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   480
      Left            =   270
      TabIndex        =   8
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Lastname:"
      Height          =   195
      Left            =   390
      TabIndex        =   18
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Provincial  Address:"
      Height          =   195
      Left            =   390
      TabIndex        =   17
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Firstname:"
      Height          =   195
      Left            =   390
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Middlename:"
      Height          =   195
      Left            =   390
      TabIndex        =   15
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "City Tel.No:"
      Height          =   195
      Left            =   390
      TabIndex        =   14
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Cellphone No:"
      Height          =   195
      Left            =   390
      TabIndex        =   13
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "City Address :"
      Height          =   195
      Left            =   390
      TabIndex        =   12
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Provincial Tel :"
      Height          =   195
      Left            =   390
      TabIndex        =   10
      Top             =   2460
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddEditInstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
 UnLockedAllTextBoxes
 DisableAddEdit
 ClearTextboxes
End Sub

Private Sub cmdCancel_Click()
 ' later ....
 Me.Hide
End Sub

Private Sub cmdOk_Click()
 Unload Me
End Sub

'Private Sub cmdSave_Click()
'Dim xxx As New clsAddIns.clsAddInstructor
'Dim i As Long

'i = xxx.AddInstructor(txtFN.Text, txtMN.Text, txtLN.Text, txtPA.Text, txtPTel.Text, txtCA.Text, txtCTel.Text, txtCelNo.Text)
'If i = -1 Then
'  MsgBox "Error found !"
'End If
'Set xxx = Nothing
'End Sub

'Private Sub cmdSave_Click()
' If MissingEntries = True Then
'   MsgBox "Please input the missing entries.", vbOKOnly + vbCritical, "Missing Entries"
' Else
'   AddInstructor txtFN.Text, txtMN.Text, txtLN.Text, txtPA.Text, txtPTel.Text, txtCA.Text, txtCTel.Text, txtCelNo.Text
'   DisableSaveCancel
'   LockedAllTextBoxes
' End If
'End Sub


'clear textbaks
Private Sub ClearTextboxes()
Dim ctrl As Control
  For Each ctrl In Controls
   If TypeOf ctrl Is TextBox Then ctrl.Text = ""
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
 If Trim(txtFN.Text) = "" Or Trim(txtMN.Text) = "" Or Trim(txtLN.Text) = "" Or Trim(txtPA.Text) = "" Or Trim(txtPTel.Text) = "" Or Trim(txtCA.Text) = "" Or Trim(txtCTel.Text) = "" Or Trim(txtCelNo.Text) = "" Then
  MissingEntries = True
 Else
  MissingEntries = False
 End If
End Function

Private Sub DisableSaveCancel()
 cmdSave.Enabled = False
 cmdCancel.Enabled = False
 'cmdAdd.Enabled = True
' cmdEdit.Enabled = True
End Sub
Private Sub DisableAddEdit()
 cmdSave.Enabled = True
 cmdCancel.Enabled = True
 'cmdAdd.Enabled = False
 'cmdEdit.Enabled = False
End Sub
Private Sub Form_Load()
' LockedAllTextBoxes
' DisableSaveCancel
End Sub
Private Sub AddEditCheck()
 
End Sub
