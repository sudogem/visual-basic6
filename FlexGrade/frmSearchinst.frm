VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSearchinst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "search teacher w/ their corresponding section being handled for principalForm1"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLastname 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Text            =   "txtLastname"
      Top             =   720
      Width           =   2355
   End
   Begin VB.TextBox txtFirstname 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "txtFirstname"
      Top             =   1140
      Width           =   2355
   End
   Begin VB.TextBox txtMi 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "txtMi"
      Top             =   1560
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Teacher Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6195
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   5955
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmSearchinst.frx":0000
            Height          =   2775
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   4895
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "PLN"
               Caption         =   "PLN"
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
               DataField       =   "PFN"
               Caption         =   "PFN"
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
            BeginProperty Column02 
               DataField       =   "PMi"
               Caption         =   "PMi"
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
            BeginProperty Column03 
               DataField       =   "SectionName"
               Caption         =   "SectionName"
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
            BeginProperty Column04 
               DataField       =   "GradeLevel"
               Caption         =   "GradeLevel"
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
            BeginProperty Column05 
               DataField       =   "SchoolYear"
               Caption         =   "SchoolYear"
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
            BeginProperty Column06 
               DataField       =   "SectionID"
               Caption         =   "SectionID"
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
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   915.024
               EndProperty
            EndProperty
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         DataField       =   "SchoolYear"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton cmdRest 
         Caption         =   "&Reset All records"
         Height          =   375
         Left            =   2460
         TabIndex        =   8
         Top             =   5760
         Width           =   1515
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4020
         TabIndex        =   7
         Top             =   5760
         Width           =   1515
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3360
         Top             =   1920
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         RecordSource    =   $"frmSearchinst.frx":0015
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
      Begin VB.Label Label4 
         Caption         =   "School Year  :"
         Height          =   315
         Left            =   420
         TabIndex        =   9
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Middlename:"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Firstname :"
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Lastname :"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSearchinst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
 Unload Me
 End Sub

Private Sub Form_Load()
  txtLastname.Text = ""
  txtFirstname.Text = ""
  txtMi.Text = ""
End Sub

Private Sub txtFirstname_Change()
On Error Resume Next
  findTeacher
End Sub

Private Sub txtLastname_Change()
On Error Resume Next
  findTeacher
End Sub

Private Sub txtMi_Change()
On Error Resume Next
  findTeacher
End Sub

Sub findTeacher()
If Trim(txtLastname.Text) = "" And Trim(txtFirstname.Text) = "" And Trim(txtFirstname.Text) = "" Then
  Adodc1.Recordset.Filter = adFilterNone
  Adodc1.Refresh
Else
 Adodc1.Recordset.Filter = "PLN like '" & txtLastname.Text & "*'"
 Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & " and PFN like '" & txtFirstname.Text & "*'"
 Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & " and PMi like '" & txtMi.Text & "*'"
End If
End Sub
