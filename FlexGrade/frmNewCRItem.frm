VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewCRItem 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Teacher ClassRecord"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   915
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   8355
      Begin VB.CommandButton cmdHLp 
         Caption         =   "&Help"
         Height          =   435
         Left            =   6960
         TabIndex        =   11
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   5700
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   435
         Left            =   4320
         TabIndex        =   9
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pick a Section , Grading period and SchoolYear ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1260
      Width           =   8355
      Begin VB.Frame Frame4 
         Caption         =   "School Year"
         ForeColor       =   &H00FF0000&
         Height          =   1635
         Left            =   4860
         TabIndex        =   6
         Top             =   2340
         Width           =   3315
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   540
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM tblSchoolYear;"
            Caption         =   "Adodc3"
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmNewCRItem.frx":0000
            Height          =   1215
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2143
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
            ColumnCount     =   1
            BeginProperty Column00 
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Grading Period"
         ForeColor       =   &H00FF0000&
         Height          =   1995
         Left            =   4860
         TabIndex        =   4
         Top             =   300
         Width           =   3315
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   420
            Top             =   1200
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   582
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM tblGradingPeriod;"
            Caption         =   "Adodc2"
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
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frmNewCRItem.frx":0015
            Height          =   1395
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2461
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "Gradingperiod"
               Caption         =   "Gradingperiod"
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
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Class Record"
         ForeColor       =   &H00FF0000&
         Height          =   3675
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4695
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   660
            Top             =   2400
            Visible         =   0   'False
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"frmNewCRItem.frx":002A
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmNewCRItem.frx":0186
            Height          =   3195
            Left            =   120
            TabIndex        =   3
            Top             =   300
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   5636
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
               DataField       =   "ClassRecordID"
               Caption         =   "ClassRecordID"
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
               DataField       =   "Subject"
               Caption         =   "Subject"
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
            BeginProperty Column03 
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
            BeginProperty Column04 
               DataField       =   "Gradelevel"
               Caption         =   "Gradelevel"
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
               DataField       =   "PersonID"
               Caption         =   "PersonID"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   915.024
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
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   7635
   End
End
Attribute VB_Name = "frmNewCRItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdOk_Click()
 ClassRecordID = Adodc1.Recordset.Fields("ClassRecordId").Value
 GradingPeriod = Adodc2.Recordset.Fields("GradingPeriod").Value
 SchoolYear = Adodc1.Recordset.Fields("Schoolyear").Value
 SectionID = Adodc1.Recordset.Fields("SectionID").Value
 frmAcademic.Show vbModal
End Sub
Private Sub Form_Load()
  Adodc1.RecordSource = _
  "SELECT tblClassRecord.ClassRecordID, tblSubject.Subject, tblSection.SectionName, tblSection.Gradelevel, tblSection.SchoolYear, tblClassRecord.PersonID " & _
  "FROM tblSubject INNER JOIN (tblSection INNER JOIN tblClassRecord ON tblSection.SectionID = tblClassRecord.SectionID) ON tblSubject.SubjectID = tblClassRecord.SubjectID " & _
  "WHERE PersonId = " & Teacher
  
  Label1.Caption = "Where would you like to put this " & Category & "."
  'MsgBox Teacher
End Sub
