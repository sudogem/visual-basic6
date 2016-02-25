VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddTeacher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form for principal : adding teacher"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   60
      TabIndex        =   14
      Top             =   2940
      Width           =   9495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.TextBox txtLastname 
      DataField       =   "Lastname"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtFirstname 
      DataField       =   "Firstname"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   5700
      TabIndex        =   5
      Tag             =   "1"
      Text            =   "Text2"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtMi 
      DataField       =   "Mi"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   8880
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "1"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6180
      MultiLine       =   -1  'True
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "frmAddTeacher.frx":0000
      Top             =   900
      Width           =   3075
   End
   Begin VB.ComboBox cboGender 
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frmAddTeacher.frx":0008
      Left            =   1800
      List            =   "frmAddTeacher.frx":0012
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   900
      Width           =   1275
   End
   Begin VB.TextBox txtAge 
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   4380
      TabIndex        =   1
      Tag             =   "1"
      Text            =   "Text4"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtTelno 
      DataField       =   "TelNo"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Tag             =   "1"
      Text            =   "Text6"
      Top             =   1440
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   5880
      Top             =   1740
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   688
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\brainwired\VB\exer1\GradeSys.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\brainwired\VB\exer1\GradeSys.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblStudinfo"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lastname      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   13
      Top             =   420
      Width           =   1395
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
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mi :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8460
      TabIndex        =   11
      Top             =   420
      Width           =   435
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender         :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   8
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.No         :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1500
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdModify_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub txtLastname_Change()

End Sub
