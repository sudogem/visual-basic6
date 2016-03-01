Attribute VB_Name = "GetData"
Option Explicit

' populate flexgrid with Table instructors ...........................................
Public Sub GetAllInstructors(flexgrid As MSFlexGrid)
  Dim RS As New ADODB.Recordset
  Dim CNN As New ADODB.Connection
  
  InitDBVariables
  CNN.Open ConnectionString
  RS.Open "Select * From tblInstructor", CNN, adOpenStatic, adLockOptimistic
  
  While Not RS.EOF
   'flexgrid.AddNew Array("Firstname", "Middlename", "Lastname", "ProvAddress", "ProvTel", "CityAddress", "CityTel", "Celphone"), Array(FN, MN, LN, PAdd, PTel, CAdd, CTel, Cel)
   
   RS.MoveNext
  Wend
  
  RS.Close
  CNN.Close
 
End Sub
'populate msflexgrid with all subjects ............................................
Public Sub GetAllSubject(flexgrid As MSFlexGrid)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblInstructor", cn, adOpenStatic, adLockOptimistic
  While Not RS.EOF
    
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close

End Sub
' populate combo box with Table instructors ...........................................
Public Sub GetAllInstructorsold(cbo As ComboBox)
  Dim RS As New ADODB.Recordset
  Dim CNN As New ADODB.Connection
  
  InitDBVariables
  CNN.Open ConnectionString
  RS.Open "Select * From tblInstructor", CNN, adOpenStatic, adLockOptimistic
  
  While Not RS.EOF
   
   RS.MoveNext
  Wend
  
  RS.Close
  CNN.Close
 
End Sub
'get the list of all curriculum.....................................
Public Sub GetCurriculum(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblCurriculum", cn, adOpenStatic, adLockOptimistic
  cbo.AddItem "---"
  While Not RS.EOF
    cbo.AddItem RS.Fields("YearFirstImplemented").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close
End Sub

'get the list of all courses.........................................
Public Sub GetListCourse(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblCourse", cn, adOpenStatic, adLockOptimistic
  cbo.AddItem "---"
  While Not RS.EOF
    cbo.AddItem RS.Fields("Course").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close
End Sub

' Get course type.........................................
Public Sub GetCourseType(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection
  
  InitDBVariables
  cn.Open ConnectionString
  RS.Open "Select * from tblCourseType", cn, adOpenStatic, adLockOptimistic
    
   While Not RS.EOF
      cbo.AddItem RS.Fields("CourseType").Value
      RS.MoveNext
   Wend
    
  RS.Close
  cn.Close
End Sub
' populate all subjects using listbox.......................................
Public Sub GetAllSubjects(lst As ListBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblSubject", cn, adOpenStatic, adLockOptimistic
  While Not RS.EOF
    lst.AddItem RS.Fields("Subject").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close

End Sub
'get semester.........................................
Public Sub GetListSemester(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblSemester", cn, adOpenStatic, adLockOptimistic
  cbo.AddItem "---"
  While Not RS.EOF
    cbo.AddItem RS.Fields("Semester").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close
End Sub
' get the lists of all schoolyear.........................................
Public Sub GetSchoolYear(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "Select * from tblSchoolYear", cn, adOpenStatic, adLockOptimistic
  cbo.AddItem "---"
  While Not RS.EOF
    cbo.AddItem RS.Fields("SchoolYear").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close
End Sub

' get all class name.........................................
Public Sub GetClass(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "SELECT tblClassToSchedule.Class, tblSemClass.YearLevel, tblSemClass.SchoolYear, tblSemClass.Semester, tblSemClass.CurriculumID FROM tblClassToSchedule INNER JOIN tblSemClass ON tblClassToSchedule.ClassToSchedID = tblSemClass.SemClassID", cn, adOpenStatic, adLockOptimistic
  While Not RS.EOF
    cbo.AddItem RS.Fields("Class").Value
    RS.MoveNext
  Wend
  
  RS.Close
  cn.Close
End Sub
'get year level.........................................
Public Sub GetYearLevel(cbo As ComboBox)
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

InitDBVariables
cn.Open ConnectionString
RS.Open "select * from tblYearLevel", cn, adOpenStatic, adLockPessimistic
  cbo.AddItem "---"
  While Not RS.EOF
    cbo.AddItem RS.Fields("YearLevel").Value
    RS.MoveNext
  Wend

RS.Close
cn.Close
End Sub

