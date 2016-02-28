VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChangeteacherpw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frmChangeteacherpw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtConfirmpw 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtPersonType 
         Height          =   285
         Left            =   300
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   -60
         Top             =   2100
         Visible         =   0   'False
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\fleXgrade\gradeSys.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblPerson"
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
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   2040
         Width           =   1395
      End
      Begin VB.TextBox txtNewpw 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtOldpw 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Confirm Password :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "New Password :"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Old Password   :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmChangeteacherpw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 If txtNewpw.Text = txtConfirmpw.Text And txtOldpw.Text = Adodc1.Recordset.Fields("password").Value Then
    Adodc1.Recordset.Update
      'Adodc1.Recordset!UserName = Text4.Text
    Adodc1.Recordset!Password = txtNewpw.Text
    Adodc1.Recordset.Update
  Else: MsgBox "Sorry, dili pwede"
  End If


End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = ConnectMe
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM tblPerson"
End Sub
