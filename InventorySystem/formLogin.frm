VERSION 5.00
Begin VB.Form formLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "formLogin"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   FillColor       =   &H80000001&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      Picture         =   "formLogin.frx":0000
      ScaleHeight     =   3375
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.TextBox username 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Text            =   "admin"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "admin"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CommandButton LoginButton 
         Appearance      =   0  'Flat
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaskColor       =   &H80000000&
         Picture         =   "formLogin.frx":282D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cancelButton 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         Picture         =   "formLogin.frx":2F88
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelButton_Click()
End
End Sub

Private Sub submit()
    Dim cmd As New ADODB.Command
    Dim rs As ADODB.Recordset
    Dim sConnString  As String
    Dim iCtr As Integer
    Dim ErrorCause As String
    Dim result
    Main
    sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath
    conn.Open sConnString
    
    cmd.ActiveConnection = conn
    
    cmd.CommandText = "SELECT * FROM USERS WHERE username='" & username.Text & "' AND status_flag=TRUE"
    cmd.CommandType = adCmdText
    Set rs = cmd.Execute
    
        If Not rs.EOF Then
            If (rs!username = username.Text And rs!password = password.Text) Then
                mUserID = rs!userid
                mName = rs!firstname & " " & rs!lastname
                'MsgBox "Welcome " & mName & "(" & mUserID & ")"
                Unload Me
                result = checkLevels(conn)
                If result = True Then
                    MsgBox "There are items that reached reorder level!" & vbCrLf & "To view details, go to 'Inventory Menu > Inventory Update'"
                End If
                Call ShowContainers
            Else
                If (rs!username <> username.Text And rs!password <> password.Text) Then
                    ErrorCause = "username and password"
                    username.SetFocus
                ElseIf (rs!password <> password.Text) Then
                    ErrorCause = "password"
                    password.SetFocus
                Else
                    ErrorCause = "username"
                    username.SetFocus
                End If
                
                Call ShowError("Login Failed! wrong " + ErrorCause + ".")
            End If
        Else
            Call ShowError("Login Failed! User does not exist.")
        End If
        
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub LoginButton_Click()
Call submit
End Sub

