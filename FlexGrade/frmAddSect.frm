VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAddSect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Add Sections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   3480
         Top             =   1320
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "SELECT * FROM tblSchoolyear"
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
      Begin MSDataListLib.DataCombo dcboSchyr 
         Bindings        =   "frmAddSect.frx":0000
         DataField       =   "SchoolYear"
         DataSource      =   "Adodc4"
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "SchoolYear"
         Text            =   "DataCombo1"
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   480
         Top             =   2400
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
         Left            =   2520
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM tblGradeLevel"
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
      Begin VB.TextBox Text2 
         DataField       =   "GradeLevel"
         DataSource      =   "Adodc2"
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   4095
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   2940
            TabIndex        =   13
            Top             =   300
            Width           =   975
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "&Modify"
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            Top             =   300
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Close"
         Height          =   435
         Left            =   3420
         TabIndex        =   8
         Top             =   2760
         Width           =   915
      End
      Begin VB.CommandButton cmdMovLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   2460
         TabIndex        =   7
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdMovnext 
         Caption         =   ">"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdMovPrev 
         Caption         =   "<"
         Height          =   375
         Left            =   1140
         TabIndex        =   5
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton cmdMovFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   420
         TabIndex        =   4
         Top             =   2760
         Width           =   675
      End
      Begin VB.TextBox Text1 
         DataField       =   "SectionName"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   2580
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXGrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM tblSection"
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
      Begin VB.Label Label3 
         Caption         =   "School Year :"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Grade Level :"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Section Name :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmAddSect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Command12_Click()
End Sub

Private Sub cmdAdd_Click()
 Textbox1 (6)
 'Adodc1.Recordset.AddNew
 'Adodc2.Recordset.AddNew
 Adodc1.Recordset.AddNew
 Adodc2.Recordset.AddNew
 Adodc4.Recordset.AddNew
 cmdAdd.Enabled = False
 cmdSAve.Enabled = True
 cmdCancel.Enabled = True
 
End Sub
Private Sub cmdCancel_Click()
'Adodc1.Recordset.CancelUpdate
'Adodc3.Recordset.CancelUpdate
Adodc1.Recordset.CancelUpdate
Adodc2.Recordset.CancelUpdate
Unload Me
End Sub

Private Sub cmdModify_Click()
 Textbox1 (5)
 cmdModify.Enabled = False
 cmdCancel.Enabled = True
 cmdSAve.Enabled = True
 cmdAdd.Enabled = False
 'Adodc1.RecordSource = _
 '"UPDATE INTO tblSection SET Sectionname = '" & Text1.Text & "'"
 Adodc3.RecordSource = _
 "UPDATE INTO tblGradeLevel.GradeLevel,tblSection.SectionName SET tblGradeLevel.GradeLevel = '" & Text2.Text & "',tblSection.SectionName = '" & Text1.Text & "'"
 'Adodc3.RecordSource = _
 '"UPDATE tblGradeLevel INNER JOIN tblSection ON tblGradeLevel.GradelevelID = tblSection.GradelevelID SET tblSection.SectionName = '" & text1.Text & "', tblGradeLevel.GradeLevel = '" & text2.Text "'"
 Adodc3.Recordset.Update Array("STUDID", "STUDLN"), Array("013-2000-0442", "Alvin")
  
End Sub

Private Sub cmdMovFirst_Click()
' Adodc1.Recordset.MoveFirst
' Adodc2.Recordset.MoveFirst
 Adodc3.Recordset.MoveFirst
End Sub

Private Sub cmdMovLast_Click()
' Adodc1.Recordset.MoveLast
' Adodc2.Recordset.MoveLast
 Adodc3.Recordset.MoveLast

End Sub

Private Sub cmdMovnext_Click()
 'Adodc1.Recordset.MoveNext
  Adodc3.Recordset.MoveNext
Text3.Text = Adodc1.Recordset.Fields("SectionName").Value
 If Adodc3.Recordset.EOF = True Then
  MsgBox "Sorry, no more records.", vbOKOnly + vbCritical, "Warning"
  'Adodc1.Recordset.MoveLast
  Adodc3.Recordset.MoveLast
  
 End If
 
End Sub

Private Sub cmdMovPrev_Click()
 Adodc3.Recordset.MovePrevious
 If Adodc3.Recordset.BOF = True Then
  MsgBox "Sorry, no more records.", vbOKOnly + vbCritical, "Warning"
  'Adodc1.Recordset.MoveFirst
  'Adodc2.Recordset.MoveFirst
  Adodc3.Recordset.MoveFirst
 End If
  
End Sub

Private Sub cmdSave_Click()
 ' tabang mama help me..
 Textbox1 (5)
 'Adodc1.Recordset.Update ' for section
 'Adodc2.Recordset.Update 'for gradelevel
 'Adodc3.Recordset.Update 'for gradelevel
 ' Adodc3.RecordSource = _
 '"UPDATE INTO tblGradeLevel.GradeLevel,tblSection.SectionName SET tblGradeLevel.GradeLevel = '" & Text2.Text & "',tblSection.SectionName = '" & Text1.Text & "'"

 'Adodc1.Refresh
 'Adodc2.Refresh
 'UPDATE tblGradeLevel INNER JOIN (tblSchoolYear INNER JOIN (tblSection INNER JOIN tblStudSection ON tblSection.SectionID = tblStudSection.SectionID) ON tblSchoolYear.SchoolYear = tblSection.SchoolYear) ON tblGradeLevel.GradelevelID = tblSection.GradelevelID SET;
 'Adodc3.RecordSource = "UPDATE tblGradeLevel INNER JOIN (tblSchoolYear INNER JOIN tblSection ON tblSchoolYear.SchoolYear = tblSection.SchoolYear) ON tblGradeLevel.GradelevelID = tblSection.GradelevelID SET SectionName = '" & Text1.Text & "', GradeLevel = '" & Text2.Text & "'"

 Adodc1.Recordset!SectionName = Text1.Text
 Adodc2.Recordset!GradeLevel = Text2.Text
 Adodc4.Recordset!SchoolYear = dcboSchyr.Text
 cmdSAve.Enabled = False
 cmdAdd.Enabled = True
 cmdModify.Enabled = True
 cmdCancel.Enabled = False
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Textbox1 (4)
 cmdCancel.Enabled = False
 cmdSAve.Enabled = False
 Adodc1.ConnectionString = ConnectMe
 Adodc2.ConnectionString = ConnectMe
 Adodc1.RecordSource = "SELECT * FROM tblSection"
 Adodc1.Refresh
 Adodc2.RecordSource = "SELECT * FROM tblGradeLevel"
 Adodc2.Refresh
 Set Text1.DataSource = Adodc1
 Text1.DataField = "SectionName"
 Set Text2.DataSource = Adodc2
 Text2.DataField = "GradeLevel"
 
End Sub
