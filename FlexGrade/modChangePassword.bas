Attribute VB_Name = "modChangePassword"
Public Function ChangePassword(UserType As String) As String
Case "LOG_TEACHER"
Adodc1.RecordSource = "SELECT username,password,personid,persontype FROM tblPerson where username = '" & frmLog.txtUsername.Text & "'"
  Adodc1.Refresh
  If Adodc1.Recordset.RecordCount > 0 Then
    If Adodc1.Recordset.Fields("Password").Value = frmLog.txtPassword.Text Then
      Select Case UCase(Adodc1.Recordset.Fields(3).Value)
        Case "TEACHER"
           User_Type = LOG_TEACHER
           Adodc1.Recordset.Fields("Password").Value = frmChangeteacherpw.txtNewpw.Text
           Adodc1.Recordset.Update
           MsgBox "Your password was sucessfully changed!", vbOKOnly + vbInformation, "Confirmation"
           Unload Me
     End Select
    Else
       MsgBox "Invalid password"
    End If
  Else
       MsgBox "Invalid password rec >0"
  End If
End Function
