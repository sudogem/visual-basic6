VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log-in"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6015
      Begin VB.CommandButton Command3 
         Caption         =   "&Help"
         Height          =   435
         Left            =   4140
         TabIndex        =   7
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   2700
         TabIndex        =   6
         Top             =   2100
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   435
         Left            =   1500
         TabIndex        =   5
         Top             =   2100
         Width           =   1155
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1500
         TabIndex        =   4
         Text            =   "txtPassword"
         Top             =   1560
         Width           =   3435
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1500
         TabIndex        =   3
         Text            =   "txtUsername"
         Top             =   1080
         Width           =   3435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Camputhaw Elementary School Grading System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   540
         Width           =   4875
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   300
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Welcome to "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1500
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Password  :"
         Height          =   375
         Left            =   540
         TabIndex        =   2
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Username  :"
         Height          =   315
         Left            =   540
         TabIndex        =   1
         Top             =   1140
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2580
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(UCase(txtUsername.Text)) = "GUEST" Then
  MsgBox "You logged as a GUEST."
  frmGuest.Show vbModal
Else
  Adodc1.RecordSource = "SELECT username,password,personid,persontype FROM tblPerson where username = '" & txtUsername.Text & "'"
  Adodc1.Refresh
  
  tempUsername = txtUsername.Text
  'frmMainfrmIns.Caption = "The username is " & tempUsername  ' save username
  If Adodc1.Recordset.RecordCount > 0 Then
    If Adodc1.Recordset.Fields("Password").Value = txtPassword.Text Then
      ' you can enter now...
      tempPassword = txtPassword.Text                       ' save password
        frmMainfrmIns.Caption = "The username and password is " & tempUsername & " /" & tempPassword ' save username
      Select Case UCase(Adodc1.Recordset.Fields(3).Value)
        Case "PRINCIPAL"
          UserType = LOG_PRINCIPAL
          Me.Hide
          frmMainfrmPrinc.Show
        Case "TEACHER"
          UserType = LOG_TEACHER
          Me.Hide
          frmMainfrmIns.Show
        Case "STAFF"
          UserType = LOG_STAFF
          frmMainfrmStaff.Show
          Me.Hide
      End Select
     
    Else
     MsgBox "Invalid Password!,please try again", vbOKOnly + vbCritical, "Log-in Message"
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"
    End If
  Else
     MsgBox "Invalid Username!,please try again", vbOKOnly + vbCritical, "Log-in Message"
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"
  End If
End If
End Sub
Private Sub Command2_Click()
 End
End Sub
Private Sub Form_Load()
 Adodc1.ConnectionString = ConnectMe
 Adodc1.CommandType = adCmdUnknown
 Adodc1.RecordSource = "SELECT * FROM tblPerson"
 frmLog.txtUsername.Text = ""
 frmLog.txtPassword.Text = ""
 'MsgBox "open splash ...", vbOKOnly
End Sub
