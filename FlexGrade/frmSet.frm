VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Set SchoolYear"
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
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin VB.CommandButton cmdNewSchyr 
         Caption         =   "&Add School Year"
         Height          =   375
         Left            =   420
         TabIndex        =   5
         Top             =   840
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3420
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2340
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmSet.frx":0000
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SchoolYear"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         Caption         =   "School Year:"
         Height          =   315
         Left            =   420
         TabIndex        =   2
         Top             =   420
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   1560
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
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdNewSchyr_Click()
 frmNewSchyr.Show vbModal
End Sub
Private Sub cmdOk_Click()
If DataCombo1.Text = "" Then
  MsgBox "Please choose SchoolYear from list or add one by clicking on the Add schoolyear button", vbOKOnly, "Missing Information"
Else
  SchoolYear = DataCombo1.Text
  Select Case UserType
    Case LOG_STAFF
      frmStaff.Show vbModal
    Case LOG_PRINCIPAL
      frmClassRec.Show vbModal
  End Select
End If
End Sub
Private Sub Command1_Click()
 Unload Me
End Sub

