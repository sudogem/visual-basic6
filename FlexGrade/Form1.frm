VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Reset all"
      Height          =   555
      Left            =   5640
      TabIndex        =   22
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   615
      Left            =   5640
      TabIndex        =   21
      Top             =   4380
      Width           =   735
   End
   Begin VB.CommandButton cmdarrowleft 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   3900
      Width           =   735
   End
   Begin VB.CommandButton cmdarrowright 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   3480
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Assign student in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   6540
      TabIndex        =   13
      Top             =   60
      Width           =   5295
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   330
         Left            =   3180
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "tblSchoolYear"
         Caption         =   "Adodc6"
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
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   3120
         Top             =   1860
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "tblStudSection"
         Caption         =   "Adodc5"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   3180
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "tblGradeLevel"
         Caption         =   "Adodc4"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   3180
         Top             =   300
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
         CommandType     =   2
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
         RecordSource    =   "tblSection"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   840
         Top             =   5280
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
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
      Begin MSDataListLib.DataCombo dcboSchYr 
         Bindings        =   "Form1.frx":0000
         DataField       =   "SchoolYear"
         DataSource      =   "Adodc6"
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SchoolYear"
         Text            =   "DataCombo3"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton cmdAddsec 
         Caption         =   "New &Section"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1860
         Width           =   1875
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Form1.frx":0015
         DataField       =   "GradeLevel"
         DataSource      =   "Adodc4"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "7"
         BoundColumn     =   "GradeLevel"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form1.frx":002A
         DataField       =   "SectionName"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   "DataCombo1"
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form1.frx":005B
         Height          =   2715
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4789
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "StudID"
            Caption         =   "StudID"
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
         BeginProperty Column02 
            DataField       =   "Firstname"
            Caption         =   "Firstname"
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
            DataField       =   "MI"
            Caption         =   "MI"
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
            DataField       =   "Address"
            Caption         =   "Address"
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
            DataField       =   "Age"
            Caption         =   "Age"
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
            DataField       =   "Gender"
            Caption         =   "Gender"
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
         BeginProperty Column07 
            DataField       =   "TelNo"
            Caption         =   "TelNo"
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
         BeginProperty Column08 
            DataField       =   "GuardiansName"
            Caption         =   "GuardiansName"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   915.024
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Section        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade Level :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1260
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search student record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin MSAdodcLib.Adodc Adodc7 
         Height          =   330
         Left            =   180
         Top             =   5280
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         RecordSource    =   "tblStudinfo"
         Caption         =   "Adodc7"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3000
         Top             =   5220
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
         RecordSource    =   "SELECT * FROM tblStudSection"
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1500
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1500
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   660
         Width           =   1995
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1020
         Width           =   1995
      End
      Begin VB.CommandButton cmdAddns 
         Caption         =   "&New students "
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1860
         Width           =   1635
      End
      Begin VB.CommandButton cmdviewrec 
         Caption         =   "&View record"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1860
         Width           =   1635
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":0070
         Height          =   2835
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5001
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483634
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   3
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "StudID"
            Caption         =   "StudID"
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
         BeginProperty Column02 
            DataField       =   "Firstname"
            Caption         =   "Firstname"
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
            DataField       =   "MI"
            Caption         =   "MI"
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
            DataField       =   "Address"
            Caption         =   "Address"
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
            DataField       =   "Age"
            Caption         =   "Age"
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
            DataField       =   "Gender"
            Caption         =   "Gender"
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
         BeginProperty Column07 
            DataField       =   "TelNo"
            Caption         =   "TelNo"
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
         BeginProperty Column08 
            DataField       =   "GuardiansName"
            Caption         =   "GuardiansName"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   915.024
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lastname :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Middlename :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSchYr_Click()
  Adodc4.RecordSource = _
 "SELECT tblGradeLevel.*, tblSection.*, tblSchoolYear.* FROM tblSchoolYear INNER JOIN (tblGradeLevel INNER JOIN tblSection ON [tblGradeLevel].[GradelevelID]=[tblSection].[GradelevelID]) ON [tblSchoolYear].[SchoolYear]= '" & cboSchYr.Text & "'"
End Sub

Private Sub cmdAddns_Click()
  frmAddViewStud.Show vbModal
End Sub

Private Sub cmdAddsec_Click()
 frmAddSect.Show vbModal
End Sub

Private Sub cmdarrowright_Click()

Adodc1.Recordset.AddNew
  With Adodc1.Recordset
    !SectionID = Adodc3.Recordset.Fields("SectionID").Value
    !StudID = Adodc7.Recordset.Fields("StudID").Value
    '!GradeLevelID = Adodc4.Recordset.Fields("GradeLevelID").Value
    .Update
  End With
  Adodc2.RecordSource = "SELECT * FROM tblStudinfo"
  Adodc2.Refresh
End Sub

Private Sub cmdviewrec_Click()
 Textbox1 (1)
 frmAddViewStud.Show vbModal
 frmAddViewStud.cmdSAve.Enabled = False
 frmAddViewStud.cmdModify.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command4_Click()
 End
End Sub
Private Sub DataCombo3_Click(Area As Integer)
'Adodc4.RecordSource = "SELECT tblGradeLevel.*, tblSection.*, tblSchoolYear.* FROM tblSchoolYear INNER JOIN (tblGradeLevel INNER JOIN tblSection ON [tblGradeLevel].[GradelevelID]=[tblSection].[GradelevelID]) ON [tblSchoolYear].[SchoolYear]=[tblSection].[SchoolYear]"
End Sub

Private Sub dcboSchyr_Click(Area As Integer)
' Adodc4.RecordSource = _
' "SELECT tblGradeLevel.*, tblSection.*,tblSchoolYear.* FROM tblSchoolYear,tblSection,tblGradeLevel WHERE tblSection.SchoolYear =  '" & dcboSchYr.Text & "' and tblGradeLevel.GradeLevelID = tblSection.GradeLevel"
Adodc4.RecordSource = "SELECT tblGradeLevel.*, tblSection.*, tblSchoolYear.* FROM tblSchoolYear INNER JOIN (tblGradeLevel INNER JOIN tblSection ON [tblGradeLevel].[GradelevelID]=[tblSection].[GradelevelID]) ON [tblSchoolYear].[SchoolYear]=[tblSection].[SchoolYear] where tblSection.SchoolYear = '" & dcboSchyr.Text & "'"
'Adodc4.Refresh
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
 DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
 'Adodc4.RecordSource = _
 '"SELECT tblGradeLevel.*, tblSection.*, tblSchoolYear.* FROM tblSchoolYear INNER JOIN (tblGradeLevel INNER JOIN tblSection ON [tblGradeLevel].[GradelevelID]=[tblSection].[GradelevelID]) ON [tblSchoolYear].[SchoolYear]= "
 ' "SELECT tblGradeLevel.*, tblSection.*, tblSchoolYear.* FROM tblSchoolYear INNER JOIN (tblGradeLevel INNER JOIN tblSection ON [tblGradeLevel].[GradelevelID]=[tblSection].[GradelevelID]) ON [tblSchoolYear].[SchoolYear]=[tblSection].[SchoolYear]"
End Sub

Private Sub Text1_Change()
 Adodc1.Recordset.Filter = "Lastname = '" & Text1.Text & "%'"

End Sub

Private Sub Text2_Change()
 Adodc1.Recordset.Filter = "Firstname = '" & Text2.Text & "%'"
 Adodc1.Refresh
End Sub

Private Sub Text3_Change()
   Adodc1.Recordset.Filter = "Mi = ' " & Text3.Text & "%'"
   Adodc1.Refresh
End Sub
