VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewSchyr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New SchoolYear"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmNewSchyr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add School Year"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmNewSchyr.frx":0442
         Height          =   495
         Left            =   300
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   1080
         Visible         =   0   'False
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
         RecordSource    =   "tblSchoolYear"
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
      Begin MSMask.MaskEdBox Text1 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   420
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "School Year  :"
         Height          =   375
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNewSchyr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo Err
If StrComp(Text1.Text, "____-____", vbBinaryCompare) = 0 Then
 MsgBox "Please input the schoolyear.", vbOKOnly + vbCritical, "Add SchoolYear"
 Text1.SetFocus
 Exit Sub
Else
 Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields(0).Value = Text1.Text
 Adodc1.Recordset.Update
 MsgBox "Schoolyear " & Text1.Text & " was successfully added to database.", vbOKOnly + vbInformation, "Add Successful"
 Unload Me
 Exit Sub
End If
Err:
 MsgBox Err.Description
End Sub
Private Sub Form_Activate()
 Text1.SetFocus
End Sub

