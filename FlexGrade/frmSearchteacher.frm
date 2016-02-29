VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmSearchteacher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "search teacher w/ their corresponding section being handled for principal"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5580
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   1260
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSearchteacher.frx":0000
      Left            =   5580
      List            =   "frmSearchteacher.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frmSearchteacher.frx":0050
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   5580
      TabIndex        =   9
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DBCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4140
      Top             =   4620
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\brainwired\VB\exer1\GradeSys.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\brainwired\VB\exer1\GradeSys.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select gender from tblstudinfo"
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
      Height          =   2535
      Left            =   180
      TabIndex        =   6
      Top             =   2700
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1860
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   780
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2355
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher Information"
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
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Search results :"
      Height          =   435
      Left            =   180
      TabIndex        =   10
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Grade Level :"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "School Year :"
      Height          =   315
      Left            =   4440
      TabIndex        =   7
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Mi :"
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Firstname :"
      Height          =   375
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Lastname :"
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearchteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()
   Adodc1.Recordset.Filter = "Lastname = ' " & Text1.Text & "'"
   
End Sub

Private Sub Text2_Change()

End Sub
