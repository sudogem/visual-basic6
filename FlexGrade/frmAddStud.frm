VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddViewStud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form for adding students from staff"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Existing students:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   5880
      TabIndex        =   23
      Top             =   0
      Width           =   4815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3915
         Left            =   180
         TabIndex        =   25
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6906
         _Version        =   393216
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
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   120
      TabIndex        =   17
      Top             =   4620
      Width           =   10575
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6960
         TabIndex        =   26
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find student"
         Height          =   375
         Left            =   5160
         TabIndex        =   24
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdSAve 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1380
         TabIndex        =   21
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2580
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3900
         TabIndex        =   19
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   495
         Left            =   8520
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame S 
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3480
         Top             =   2220
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.TextBox txtGuardian 
         Height          =   285
         Left            =   2100
         TabIndex        =   16
         Text            =   "txtGuardian"
         Top             =   4140
         Width           =   2235
      End
      Begin VB.TextBox txtTelno 
         Height          =   285
         Left            =   2100
         TabIndex        =   14
         Text            =   "txtTelno"
         Top             =   3660
         Width           =   2235
      End
      Begin VB.TextBox txtAddress 
         Height          =   495
         Left            =   2100
         TabIndex        =   12
         Text            =   "txtAddress"
         Top             =   3000
         Width           =   2235
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   2100
         TabIndex        =   10
         Text            =   "txtAge"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmAddStud.frx":0000
         Left            =   2100
         List            =   "frmAddStud.frx":0002
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2040
         Width           =   1035
      End
      Begin VB.TextBox txtMi 
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Text            =   "txtMi"
         Top             =   1560
         Width           =   2235
      End
      Begin VB.TextBox txtFirstname 
         Height          =   315
         Left            =   2100
         TabIndex        =   5
         Text            =   "txtFirstname"
         Top             =   1020
         Width           =   2235
      End
      Begin VB.TextBox txtLastname 
         Height          =   285
         Left            =   2100
         TabIndex        =   4
         Text            =   "txtLastname"
         Top             =   540
         Width           =   2235
      End
      Begin VB.Label Label16 
         Caption         =   "Guardians Name :"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Tel. No :"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   3180
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Age :"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Gender :"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "Middlename :"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label10 
         Caption         =   "Firstname :"
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Lastname :"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   540
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmAddViewStud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp As String

Private Sub cmdAdd_Click()
 Textbox1 (2)       ' enable textboxes
 cmdAdd.Enabled = False
 cmdSAve.Enabled = True
 cmdCancel.Enabled = True
 cmdModify.Enabled = False
 Adodc1.Recordset.AddNew
End Sub

Private Sub cmdCancel_Click()
' cancel addnew operation
 Adodc1.Recordset.CancelUpdate
 Unload Me
End Sub
Private Sub cmdOK_Click()
MsgBox "ok.."
End Sub
Private Sub cmdModify_Click()  ' modify student record
' hapit nako maboang ani....
 Textbox1 (2)              'enable all textbox
 cmdModify.Enabled = False
 cmdSAve.Enabled = True
 cmdCancel.Enabled = True
 cmdAdd.Enabled = False
 Adodc1.RecordSource = _
 "UPDATE tblPerson SET Lastname ='" & txtLastname.Text & "',Firstname ='" & txtFirstname.Text & "',Mi = '" & txtMi.Text & "',address = '" & txtAddress.Text & "',Gender = '" & cboGender.Text & "',Telno = '" & txtTelno.Text & "',GuardiansName = '" & txtGuardian.Text & "'"
End Sub
Private Sub cmdSave_Click()   ' save student record
 ' hapit nako maboang mama tabang!!!!!!!!!!!!!!!!!!!!!!!
Dim stat As Boolean
Dim err As Boolean

cmdAdd.Enabled = True
cmdCancel.Enabled = False
stat = True
If Trim(txtLastname.Text) = 0 Or txtFirstname.Text = "" Or txtMi.Text = "" Or txtAge.Text = "" Or cboGender.Text = "" Or txtAddress.Text = "" Or txtTelno.Text = "" Or txtGuardian.Text = "" Then
    MsgBox "Please input the missing information....", vbOKOnly + vbCritical, "Warning!"
    
If txtLastname.Text = "" Then
  Label9.ForeColor = &HFF&
  stat = False
Else
  Label9.ForeColor = &H80000012
  stat = True
End If
If txtFirstname.Text = "" Then
  Label10.ForeColor = &HFF&
  stat = False
Else
  Label10.ForeColor = &H80000012
  stat = True
End If
If txtMi.Text = "" Then
  Label11.ForeColor = &HFF&
  stat = False
Else
  Label11.ForeColor = &H80000012
  stat = True
End If
If cboGender.Text = "" Then
  Label12.ForeColor = &HFF&
  stat = False
End If
If txtAge.Text = "" Then
  Label13.ForeColor = &HFF&
  stat = False
Else
  Label13.ForeColor = &H80000012
  stat = True
End If
If txtAddress.Text = "" Then
  Label14.ForeColor = &HFF&
  stat = False
Else
  Label14.ForeColor = &H80000012
  stat = True
End If
If txtTelno.Text = "" Then
  Label15.ForeColor = &HFF&
Else
  Label15.ForeColor = &H80000012
  stat = True
End If
If txtGuardian.Text = "" Then
  Label16.ForeColor = &HFF&
  stat = False
Else
  Label16.ForeColor = &H80000012
  stat = True
End If
Exit Sub
End If
If stat = False Then
      Textbox1 (2)                  'enable textbox
      cmdSAve.Enabled = True
      cmdModify.Enabled = False
      Adodc1.Recordset.CancelUpdate ' cancel
      MsgBox "cancel ok..."
Else
      Adodc1.Recordset.Update
      MsgBox "update ok..."
      Textbox1 (1)                  'disable textboxes
      cmdSAve.Enabled = False
      cmdModify.Enabled = True
      'Adodc1.Refresh
End If
End Sub
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Textbox1 (1) 'disable textboxes
 cmdSAve.Enabled = False
 cmdCancel.Enabled = False
 GenderType
 Adodc1.ConnectionString = ConnectMe
 Adodc1.CommandType = adCmdUnknown
 Adodc1.RecordSource = "SELECT * from tblStudinfo"
 Adodc1.ConnectionTimeout = 40
 
 Set txtLastname.DataSource = Adodc1
 txtLastname.DataField = "Lastname"
 Set txtFirstname.DataSource = Adodc1
 txtFirstname.DataField = "Firstname"
 Set txtMi.DataSource = Adodc1
 txtMi.DataField = "Mi"
 Set txtAddress.DataSource = Adodc1
 txtAddress.DataField = "Address"
 Set txtAge.DataSource = Adodc1
 txtAge.DataField = "Age"
 Set cboGender.DataSource = Adodc1
 cboGender.DataField = "Gender"
 Set txtTelno.DataSource = Adodc1
 txtTelno.DataField = "Telno"
 Set txtGuardian.DataSource = Adodc1
 txtGuardian.DataField = "Guardiansname"
 
End Sub
Sub GenderType()
 With cboGender
 .AddItem "Female"
 .AddItem "Male"
End With
End Sub
