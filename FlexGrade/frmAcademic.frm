VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAcademic 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmAcademic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   120
      TabIndex        =   8
      Top             =   2700
      Width           =   4395
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   3300
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
      RecordSource    =   "tblAcademic"
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   300
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAcademic.frx":0442
         Height          =   375
         Left            =   300
         TabIndex        =   6
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Text3 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   1740
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text2 
         Height          =   915
         Left            =   1620
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmAcademic.frx":0457
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Date   :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Description  :"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Max Grade  :"
         Height          =   435
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAcademic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOk_Click()
 cmdOk.Enabled = False
 Adodc1.Recordset.AddNew
 With Adodc1.Recordset
   !Category = Category
   !MaxGrade = Text1.Text
   !Description = Text2.Text
   !GradingPeriod = GradingPeriod
   !Date = Text3.Text
   .Update
 End With
 AcademicID = Adodc1.Recordset.Fields("AcademicID").Value
 Adodc1.Refresh
 MsgBox AcademicID
 Unload Me
 frmInitItem.Show vbModal
End Sub

Private Sub Form_Load()
 frmAcademic.Caption = Category & " for " & GradingPeriod
 frmAcademic.Frame1.Caption = "Input New " & Category
End Sub
