VERSION 5.00
Begin VB.Form frmFindTeacher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "form for find teacher records."
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Find teacher with their corresponding section(s) being handled"
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4995
   End
End
Attribute VB_Name = "frmFindTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ans As Variant
Private Sub Command1_Click()
'perform search operation here
frmAddTeacher.Adodc1.RecordSource = _
   "SELECT * FROM tblPerson WHERE PLN = '" & txtLastname.Text & "' or PFN = '" & txtFirstname.Text & "' or PMi = '" & txtMi.Text & "' PersonType = '" & txtPersonType & "'"
   
 If Adodc1.Recordset.EOF = True Then
   MsgBox "Sorry,no record found!", vbOKOnly + vbMsgBoxHelpButton, "Search result"
   Unload Me
 Else
   MsgBox "Record found!", vbOKOnly + vbMsgBoxHelpButton, "Search result"
 End If
'With Adodc1
'    While Not .Recordset.EOF = True
    '.Recordset.Find "PLN = '" & txtLastname.Text & "'"
    '.Recordset.Find "PFN = ' " & txtFirstname.Text & "'"
    '.Recordset.Find "PMi = ' " & txtMi.Text & "'"
    '.Recordset.MoveNext
'    Wend
'End With

End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Form_Load()
 txtLastname.Text = ""
 txtFirstname.Text = ""
 txtMi.Text = ""
End Sub
