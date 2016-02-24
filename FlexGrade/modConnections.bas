Attribute VB_Name = "Module1"
Option Explicit
Public Const LOG_PRINCIPAL = 0
Public Const LOG_TEACHER = 1
Public Const LOG_STAFF = 2
Public Const LOG_GUEST = 3

Public UserType As Integer

Public myDataSource As String


Public Function ConnectMe() As String
'myDataSource = "..\data\gradesys.mdb"
myDataSource = App.Path & "\data\gradesys.mdb"
ConnectMe = "Provider=Microsoft.Jet.OLEDB.4.0;" _
             & "Persist Security Info=False; Data Source = '" _
             & myDataSource & "'"
             '& "..\data\gradesys.mdb'"
End Function
