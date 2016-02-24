VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log-in"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1860
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   1620
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      Top             =   1860
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Welcome to Camputhaw Elementary Grading System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   540
      TabIndex        =   6
      Top             =   240
      Width           =   5235
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(UCase(Text1.Text)) = "GUEST" Then
 'perform for guest
MsgBox "You logged as a GUEST."

Else
  Adodc1.RecordSource = "SELECT username,password,personid,persontype FROM tblPerson where username = '" & Text1.Text & "'"
  Adodc1.Refresh
  If Adodc1.Recordset.RecordCount > 0 Then
    If Adodc1.Recordset.Fields("Password").Value = Text2.Text Then
      ' you can enter now...
      Select Case UCase(Adodc1.Recordset.Fields(3).Value)
        Case "PRINCIPAL"
          UserType = LOG_PRINCIPAL
          frmMainfrmPrinc.Show
        Case "TEACHER"
          UserType = LOG_TEACHER
        Case "STAFF"
          UserType = LOG_STAFF
      End Select
      MsgBox "ok........"
    Else
      'access denied....
      
    End If
  Else
     'access denied...
  End If
End If

  'Unload Me
  'frmStudDetails.Show
  'frmTGradeSum.Show
End Sub
Private Sub Command2_Click()
 End
End Sub
Private Sub Form_Load()
  Adodc1.ConnectionString = ConnectMe
 Adodc1.CommandType = adCmdUnknown
 Adodc1.RecordSource = "SELECT * FROM tblPerson"

 'MsgBox "open splash ...", vbOKOnly
End Sub
