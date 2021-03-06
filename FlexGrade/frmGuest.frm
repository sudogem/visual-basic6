VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGuest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "form for guest"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Student Grade Information"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   8055
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   4020
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.TextBox txtSectionName 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Print"
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtGradeLevel 
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmGuest.frx":0000
         Height          =   3075
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5424
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
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Lastname"
            Caption         =   "Lastname"
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
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Section :"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grade Level:"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Look-Up Student Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   8055
      Begin MSDataListLib.DataCombo dcboSchyr 
         Bindings        =   "frmGuest.frx":0015
         DataField       =   "SchoolYear"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   1620
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SchoolYear"
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   315
         Left            =   5340
         TabIndex        =   8
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset..."
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   1860
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4260
         Top             =   1320
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"frmGuest.frx":002A
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
      Begin VB.TextBox txtMi 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "txtMi"
         Top             =   1200
         Width           =   2715
      End
      Begin VB.TextBox txtFirstname 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "txtFirstname"
         Top             =   780
         Width           =   2715
      End
      Begin VB.TextBox txtLastname 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Text            =   "txtLastname"
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "School Year :"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Middlename :"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Firstname :"
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Lastname :"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearchStud_Click()
'search student grade

End Sub

Private Sub cmdReset_Click()
 Adodc1.RecordSource = "SELECT * FROM tblStudinfo"
 Set DataGrid1.DataSource = Adodc1
 Adodc1.Refresh
 DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub txtFirstname_Change()
  findstud
 'Adodc1.Recordset.Filter = "Firstname like '" & txtFirstname.Text & "*'"
End Sub

Private Sub txtLastname_Change()
On Error Resume Next
  findstud
 'Adodc1.Recordset.Filter = "Lastname like '" & txtLastname.Text & "*'"
End Sub

Private Sub txtMi_Change()
  findstud
 'Adodc1.Recordset.Filter = "Mi like '" & txtMi.Text & "'"
End Sub

Sub findstud()
  Adodc1.Recordset.Filter = "Firstname like '" & txtFirstname.Text & "*'"
  Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & " and Lastname like '" & txtLastname.Text & "*'"
  Adodc1.Recordset.Filter = Adodc1.Recordset.Filter & " and Mi like '" & txtMi.Text & "'"
End Sub
