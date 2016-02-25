VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddTeacher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teacher Information "
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   4800
      Width           =   10575
      Begin VB.CommandButton cmdfindteacher 
         Caption         =   "&Find Teacher"
         Height          =   375
         Left            =   5340
         TabIndex        =   14
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton cmdhelp 
         Caption         =   "&Help"
         Height          =   375
         Left            =   9060
         TabIndex        =   16
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   7740
         TabIndex        =   15
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3900
         TabIndex        =   13
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2580
         TabIndex        =   12
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdSAve 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "View Teachers:"
      ForeColor       =   &H00FF0000&
      Height          =   2400
      Left            =   3360
      TabIndex        =   28
      Top             =   2400
      Width           =   7335
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1935
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   780
      Top             =   5700
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   $"frmAddInstructor.frx":0000
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
   Begin VB.Frame Frame2 
      Caption         =   "Teacher Password"
      ForeColor       =   &H00FF0000&
      Height          =   2400
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   3135
      Begin VB.TextBox txtPassword 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Text            =   "txtPassword"
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txtUsername 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Text            =   "txtUsername"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Password :"
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Username :"
         Height          =   315
         Left            =   360
         TabIndex        =   25
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teacher Information"
      ForeColor       =   &H00FF0000&
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txttype 
         Height          =   285
         Left            =   1260
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtTelno 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Text            =   "txtTelno"
         Top             =   1380
         Width           =   2115
      End
      Begin VB.TextBox txtAge 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   4620
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   900
         Width           =   615
      End
      Begin VB.ComboBox cboGender 
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmAddInstructor.frx":0136
         Left            =   1260
         List            =   "frmAddInstructor.frx":0140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox txtAddress 
         DataSource      =   "Adodc1"
         Height          =   675
         Left            =   4620
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmAddInstructor.frx":0152
         Top             =   1380
         Width           =   5775
      End
      Begin VB.TextBox txtMi 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Text            =   "txtMi"
         Top             =   420
         Width           =   2115
      End
      Begin VB.TextBox txtFirstname 
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   4620
         TabIndex        =   2
         Text            =   "txtFirstname"
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox txtLastname 
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Text            =   "txtLastname"
         Top             =   480
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "Tel No :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Age :"
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Gender :"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Address :"
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Middlename:"
         Height          =   255
         Left            =   6960
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Firstname  :"
         Height          =   315
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Lastname  :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmAddTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
 Textbox1 (3)            'enable textboxes
 cmdAdd.Enabled = False
 cmdCancel.Enabled = True
 cmdSAve.Enabled = True
 Adodc1.Recordset.AddNew
End Sub

Private Sub cmdCancel_Click()
Adodc1.Recordset.CancelUpdate 'cancel add new record
Set frmAddTeacher = Nothing
Me.Hide
End Sub

Private Sub cmdfindteacher_Click()
 frmFindTeacher.Show vbModal
End Sub

Private Sub cmdModify_Click()
Textbox1 (3)   'enable textboxes
cmdModify.Enabled = False
cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdSAve.Enabled = True
'Adodc1.Recordset.EditMode

End Sub

Private Sub cmdSave_Click()
 Textbox1 (4)  'disable textboxes
 cmdSAve.Enabled = False
 cmdCancel.Enabled = False
 cmdModify.Enabled = True
 cmdAdd.Enabled = True
 txttype.Text = "Teacher"
 Adodc1.Recordset.MoveFirst
 DataGrid1.Refresh
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
frmAddTeacher.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 13
      SendKeys "{Tab}"
    Case 40
      SendKeys "{Tab}"
    Case 38
      SendKeys "+{Tab}"
  End Select
End Sub

Private Sub Form_Load()
Textbox1 (4)   'disable textboxes
cmdCancel.Enabled = False
cmdSAve.Enabled = False

'configure adodc1 connection and etc
Adodc1.CommandType = adCmdUnknown
Adodc1.ConnectionString = ConnectMe
Adodc1.RecordSource = "SELECT * from tblPerson" ' where tblPerson.PersonType = 'Teacher'"
Adodc1.Refresh

Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Set txtLastname.DataSource = Adodc1
txtLastname.DataField = "PLN"
Set txtFirstname.DataSource = Adodc1
txtFirstname.DataField = "PFN"
Set txtMi.DataSource = Adodc1
txtMi.DataField = "PMI"
Set cboGender.DataSource = Adodc1
cboGender.DataField = "Gender"
Set txtAge.DataSource = Adodc1
txtAge.DataField = "Age"
Set txtTelno.DataSource = Adodc1
txtTelno.DataField = "Telno"
Set txtAddress.DataSource = Adodc1
txtAddress.DataField = "Address"
Set txtUsername.DataSource = Adodc1
txtUsername.DataField = "UserName"
Set txtPassword.DataSource = Adodc1
txtPassword.DataField = "Password"
Set txttype.DataSource = Adodc1
txttype.DataField = "PersonType"

'FilterStatus = True
End Sub
Sub MyFilter()
'If FilterStatus = True Then
'  Adodc1.Recordset.Filter = adFilterNone
'  If Len(Trim(txtLastname.Text)) > 0 Then
'    Adodc1.Recordset.Filter = "PLN like '" & txtLastname.Text & "%'"
'  End If
'  If Len(Trim(txtFirstname.Text)) > 0 Then
'    Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & "PFN like '" & txtFirstname.Text & "%'"
'  End If
'  If Len(Trim(txtMi.Text)) > 0 Then
'    Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & "PMi like '" & txtMi.Text & "%'"
'  End If
  Adodc2.RecordSource = "select * from tblPerson where PMi like '" & Trim(txtMi.Text) & "%'"
  Adodc2.Refresh
  Text1.Text = Adodc2.Recordset.RecordCount
'End If
End Sub
