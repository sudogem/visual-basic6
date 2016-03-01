VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mySchedConflictChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private RS As New ADODB.Recordset
Private CNN As New ADODB.Connection
Private myConnectionString As String

Private ConflictString As String
Private myNoConflict As Boolean

Private Const NONE As Integer = -1
Private Const NOTSET As String = ""
Private Const NODATETIME As Date = 0

Private myScheduleID As Long
Private myClassID As Long
Private mySubjectOrModuleID As Long
Private myCourseType As String
Private myInstructorID As Long
Private myRoom As String
Private myDays As String
Private myStartTime As String
Private myEndTime As String
Private myStartDate As String
Private myEndDate As String

Public Property Get NoConflict() As Boolean
  NoConflict = myNoConflict
End Property

Public Property Let ConnectionString(ConnectStr As String)
  myConnectionString = ConnectStr
End Property

Public Property Get ConnectionString() As String
  ConnectionString = myConnectionString
End Property

Public Property Let ClassID(ID As Long)
  myClassID = ID
End Property

Public Property Get ClassID() As Long
  ClassID = myClassID
End Property

Public Property Let SubjectOrModuleID(ID As Long)
  mySubjectOrModuleID = ID
End Property

Public Property Get SubjectOrModuleID() As Long
  SubjectOrModuleID = mySubjectOrModuleID
End Property

Public Property Let CourseType(ctype As String)
  myCourseType = ctype
End Property

Public Property Get CourseType() As String
  CourseType = myCourseType
End Property

Public Property Let InstructorID(ID As Long)
  myInstructorID = ID
End Property

Public Property Get InstructorID() As Long
  InstructorID = myInstructorID
End Property

Public Property Let Room(rm As String)
  myRoom = rm
End Property

Public Property Get Room() As String
  Room = myRoom
End Property

Public Property Let Days(D As String)
  myDays = D
End Property

Public Property Get Days() As String
  Days = myDays
End Property

Public Property Let StartTime(s As String)
  myStartTime = s
End Property

Public Property Get StartTime() As String
  StartTime = myStartTime
End Property

Public Property Let EndTime(E As String)
  myEndTime = E
End Property

Public Property Get EndTime() As String
  EndTime = myEndTime
End Property

Public Property Get StartDate() As String
  StartDate = myStartDate
End Property

Public Property Get Enddate() As String
  Enddate = myEndDate
End Property


Private Sub ConnectMe(QueryString As String)
  If RS.State = adStateOpen Or CNN.State = adStateOpen Then
    DisconnectMe
  End If
  CNN.Open ConnectionString
  RS.Open QueryString, CNN, adOpenStatic, adLockOptimistic
End Sub

Private Sub DisconnectMe()
  RS.Close
  CNN.Close
End Sub

Private Sub Class_Initialize()
  myScheduleID = NONE
  myClassID = NONE
  mySubjectOrModuleID = NONE
  myCourseType = NOTSET
  myInstructorID = NONE
  myRoom = NOTSET
  myDays = NOTSET
  myStartTime = NODATETIME
  myEndTime = NODATETIME
  myStartDate = NODATETIME
  myEndDate = NODATETIME
  myNoConflict = False
End Sub

Private Sub Class_Terminate()
  DisconnectMe
End Sub

Private Function GetStartDateAndEndDate() As Boolean
  ConnectMe "Select * from tblClassToSchedule where ClassToSchedID=" & myClassID
  If RS.RecordCount > 0 Then
    myStartDate = RS!StartDate
    myEndDate = RS!Enddate
    GetStartDateAndEndDate = True
  Else
    MsgBox "ClassID does not exist!", , "Missing Class"
    GetStartDateAndEndDate = False
  End If
  DisconnectMe
End Function

'use all three checking functions/routines
Public Sub CheckAll()
  Dim FilterString As String
  Dim DaysFilter As String
  Dim TimeFilter As String
  
  ConnectMe "SELECT tblClassSchedules.*, tblClassToSchedule.StartDate, " & _
           "tblClassToSchedule.EndDate FROM tblClassToSchedule INNER JOIN " & _
           "tblClassSchedules ON tblClassToSchedule.ClassToSchedID = " & _
           "tblClassSchedules.ClassToScheduleID where " & _
           "((StartDate=#" & myStartDate & "#) OR (StartDate=#" & myEndDate & "#)" & _
           " OR ((StartDate>#" & myStartDate & "#) AND (StartDate<#" & myEndDate & "#))" & _
           " OR (EndDate=#" & myStartDate & "#) OR (EndDate=#" & myEndDate & "#)" & _
           " OR ((EndDate>#" & myStartDate & "#) AND (EndDate<#" & "myEndDate#)))"
    
    ConflictString = ""
    DaysFilter = GetDaysFilter(myDays)
    TimeFilter = GetTimeFilter(myStartTime, myEndTime)
    
    'checks if schedule is in conflict with current Class
    RS.Filter = adFilterNone
    RS.Filter = "ClassToSchedID=" & myClassID & " AND " & DaysFilter & " AND " & TimeFilter
    If RS.RecordCount > 0 Then
      ' there is a conflict with the class schedule
      ConflictString = ConflictString & "The Class is assigned a schedule that is in conflict with current schedule" & vbCrLf
      ' we could choose to add or show the conflicting schedules
    End If
    
    'checks if room is available or not at the time specified by schedule
    RS.Filter = adFilterNone
    RS.Filter = "Room=" & myRoom & " AND " & DaysFilter & " AND " & TimeFilter
    If RS.RecordCount > 0 Then
      ' there is a conflict with the room assigned
      ConflictString = ConflictString & "The Room chosen is currently occupied by another Class" & vbCrLf
      ' we could choose to add or show the conflicting schedules
    End If
    
    'checks if instructor is available at that time
    RS.Filter = adFilterNone
    RS.Filter = "InstructorID=" & myInstructorID & " AND " & DaysFilter & " AND " & TimeFilter
    If RS.RecordCount > 0 Then
      ' there is a conflict with the instructor's schedule
      ConflictString = ConflictString & "The instructor is not available at the chosen time and day" & vbCrLf
      ' we could choose to add or show the conflicting schedules
    End If
  DisconnectMe
End Sub

Private Function GetDaysFilter(Days As String) As String
  Dim myfilter As String
  Dim prev As Boolean
  
  myfilter = "("
  prev = False
  
  'check for sunday
  If InStr(1, Days, "Sun", vbTextCompare) > 0 Then
    myfilter = myfilter & "(Days like *Sun*)"
    prev = True
  End If
  
  'check for monday
  If InStr(1, Days, "Mon", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Mon*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Mon*)"
    End If
  End If
  
  'check for tuesday
  If InStr(1, Days, "Tue", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Tue*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Tue*)"
    End If
  End If
  
  'check for wednesday
  If InStr(1, Days, "Wed", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Wed*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Wed*)"
    End If
  End If
  
  'check for thursday
  If InStr(1, Days, "Thur", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Thur*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Thur*)"
    End If
  End If
  
  'check for friday
  If InStr(1, Days, "Fri", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Fri*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Fri*)"
    End If
  End If
  
  'check for Saturday
  If InStr(1, Days, "Sat", vbTextCompare) > 0 Then
    If prev = False Then
      myfilter = myfilter & "(Days like *Sat*)"
      prev = True
    Else
      myfilter = myfilter & "OR (Days like *Sat*)"
    End If
  End If
  
  myfilter = myfilter & ")"
  GetDaysFilter = myfilter
End Function

Private Function GetTimeFilter(sStartTime As String, sEndTime As String) As String
  Dim myfilter As String
  
  myfilter = "("
  myfilter = myfilter & "(StartTime=" & sStartTime & ") OR"
  myfilter = myfilter & "((StartTime>" & sStartTime & ") AND (StartTime<" & sEndTime & ")) OR"
  myfilter = myfilter & "(EndTime=" & sEndTime & ") OR"
  myfilter = myfilter & "((EndTime>" & sStartTime & ") AND (EndTime<" & sEndTime & "))"
  myfilter = myfilter & ")"
  
  GetTimeFilter = myfilter
End Function
