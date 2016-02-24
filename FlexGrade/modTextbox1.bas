Attribute VB_Name = "modTextbox"
Public Function Textbox1(mode As Integer) As Integer
 Select Case mode
  Case 1 ' form for staff
   frmAddViewStud.txtLastname.Locked = True
   frmAddViewStud.txtFirstname.Locked = True
   frmAddViewStud.txtMi.Locked = True
   frmAddViewStud.cboGender.Locked = True
   frmAddViewStud.txtAge.Locked = True
   frmAddViewStud.txtAddress.Locked = True
   frmAddViewStud.txtTelno.Locked = True
   frmAddViewStud.txtGuardian.Locked = True
  Case 2
   frmAddViewStud.txtLastname.Locked = False
   frmAddViewStud.txtFirstname.Locked = False
   frmAddViewStud.txtMi.Locked = False
   frmAddViewStud.cboGender.Enabled = False
   frmAddViewStud.txtAge.Locked = False
   frmAddViewStud.txtAddress.Locked = False
   frmAddViewStud.txtTelno.Locked = False
   frmAddViewStud.txtGuardian.Locked = False
   
  Case 3 ' form for adding teacher/ enable textboxes
   frmAddTeacher.txtLastname.Locked = False
   frmAddTeacher.txtFirstname.Locked = False
   frmAddTeacher.txtMi.Locked = False
   frmAddTeacher.cboGender.Locked = False
   frmAddTeacher.txtAge.Locked = False
   frmAddTeacher.txtTelno.Locked = False
   frmAddTeacher.txtAddress.Locked = False
   frmAddTeacher.txtUsername.Locked = False
   frmAddTeacher.txtPassword.Locked = False
   
  Case 4 'disable textboxes
   frmAddTeacher.txtLastname.Locked = True
   frmAddTeacher.txtFirstname.Locked = True
   frmAddTeacher.txtMi.Locked = True
   frmAddTeacher.cboGender.Locked = True
   frmAddTeacher.txtAge.Locked = True
   frmAddTeacher.txtTelno.Locked = True
   frmAddTeacher.cboGender.Locked = True
   frmAddTeacher.txtAddress.Locked = True
   frmAddTeacher.txtUsername.Locked = True
   frmAddTeacher.txtPassword.Locked = True
   
   
  Case 5 ' form for adding sections
   frmAddSect.Text1.Locked = True
   frmAddSect.Text2.Locked = True
   
  Case 6
   frmAddSect.Text1.Locked = False
   frmAddSect.Text2.Locked = False
    
 End Select
End Function

