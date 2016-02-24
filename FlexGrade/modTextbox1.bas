Attribute VB_Name = "modTextbox"
Public Function Textbox1(mode As Integer) As Integer
 Select Case mode
  Case 1 ' form for staff
   frmAddViewStud.txtLastname.Enabled = False
   frmAddViewStud.txtFirstname.Enabled = False
   frmAddViewStud.txtMi.Enabled = False
   frmAddViewStud.cboGender.Enabled = False
   frmAddViewStud.txtAge.Enabled = False
   frmAddViewStud.txtAddress.Enabled = False
   frmAddViewStud.txtTelno.Enabled = False
   frmAddViewStud.txtGuardian.Enabled = False
  Case 2
   frmAddViewStud.txtLastname.Enabled = True
   frmAddViewStud.txtFirstname.Enabled = True
   frmAddViewStud.txtMi.Enabled = True
   frmAddViewStud.cboGender.Enabled = True
   frmAddViewStud.txtAge.Enabled = True
   frmAddViewStud.txtAddress.Enabled = True
   frmAddViewStud.txtTelno.Enabled = True
   frmAddViewStud.txtGuardian.Enabled = True
   
  Case 3 ' form for adding teacher
   frmAddTeacher.txtLastname.Enabled = False
   frmAddTeacher.txtFirstname.Enabled = False
   frmAddTeacher.txtMi.Enabled = False
   frmAddTeacher.cboGender.Enabled = False
   frmAddTeacher.txtAge.Enabled = False
   frmAddTeacher.txtTelno.Enabled = False
   frmAddTeacher.txtAddress.Enabled = False
   frmAddTeacher.txtUsername.Enabled = False
   frmAddTeacher.txtPassword.Enabled = False
   
  Case 4
   frmAddTeacher.txtLastname.Enabled = True
   frmAddTeacher.txtFirstname.Enabled = True
   frmAddTeacher.txtMi.Enabled = True
   frmAddTeacher.cboGender.Enabled = True
   frmAddTeacher.txtAge.Enabled = True
   frmAddTeacher.txtTelno.Enabled = True
   frmAddTeacher.cboGender.Enabled = True
   frmAddTeacher.txtAddress.Enabled = True
   frmAddTeacher.txtUsername.Enabled = True
   frmAddTeacher.txtPassword.Enabled = True
   
   
  Case 5 ' form for adding sections
   frmAddSect.Text1.Enabled = False
   frmAddSect.Text2.Enabled = False
   
  Case 6
   frmAddSect.Text1.Enabled = True
   frmAddSect.Text2.Enabled = True
    
 End Select
End Function

