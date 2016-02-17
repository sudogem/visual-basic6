VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form formAddUser 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid grdUser 
      Height          =   3225
      Left            =   5460
      TabIndex        =   25
      Top             =   1290
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5689
      _Version        =   393216
      BackColor       =   16777215
      ColumnHeaders   =   0   'False
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
   Begin VB.ComboBox regAccountType 
      Height          =   315
      ItemData        =   "formAddUser.frx":0000
      Left            =   2310
      List            =   "formAddUser.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   3630
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   345
      Left            =   3390
      TabIndex        =   8
      Top             =   4170
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add User"
      Height          =   345
      Left            =   2340
      TabIndex        =   7
      Top             =   4170
      Width           =   945
   End
   Begin VB.TextBox regContactNumber 
      Height          =   285
      Left            =   2340
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox regAddress 
      Height          =   285
      Left            =   2340
      TabIndex        =   4
      Top             =   2850
      Width           =   2415
   End
   Begin VB.TextBox regLastName 
      Height          =   285
      Left            =   2340
      TabIndex        =   3
      Top             =   2460
      Width           =   2415
   End
   Begin VB.TextBox regFirstName 
      Height          =   285
      Left            =   2340
      TabIndex        =   2
      Top             =   2070
      Width           =   2415
   End
   Begin VB.TextBox regPassword 
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox regUserName 
      Height          =   285
      Left            =   2340
      TabIndex        =   0
      Top             =   1290
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      TabIndex        =   27
      Top             =   810
      Width           =   435
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   8250
      TabIndex        =   26
      Top             =   810
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   1770
      TabIndex        =   24
      Top             =   810
      Width           =   585
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2340
      TabIndex        =   23
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   22
      Top             =   3660
      Width           =   195
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   21
      Top             =   3240
      Width           =   195
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   20
      Top             =   2850
      Width           =   195
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   19
      Top             =   2460
      Width           =   195
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   18
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   17
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "  :  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   16
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT TYPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   15
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NUMBER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   14
      Top             =   3270
      Width           =   1665
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   13
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   12
      Top             =   2100
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   11
      Top             =   2490
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   10
      Top             =   1710
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   9
      Top             =   1350
      Width           =   1185
   End
End
Attribute VB_Name = "formAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsers As New ADODB.Recordset

Private Sub cmdAdd_Click()
    Dim result As Boolean
    openConnection
    result = AddUser(conn, regUserName.Text, regPassword.Text, regFirstName.Text, regLastName.Text, regAddress.Text, regContactNumber.Text, regAccountType.Text)
    closeConnection
    If result = True Then
        Call cmdReset_Click
        Call refreshUserInfo
    Else
        formError.errorLabel.Caption = "Cannot add user"
        formError.Show vbModal
    End If
End Sub

Private Sub cmdReset_Click()
    regUserName.Text = ""
    regPassword.Text = ""
    regFirstName.Text = ""
    regLastName.Text = ""
    regAddress.Text = ""
    regContactNumber.Text = ""
    regAccountType.Text = ""
End Sub

Private Sub Form_Load()
    Call getUsers
End Sub

Private Sub getUsers()
    Dim strquery As String
    strquery = "SELECT username, utype FROM users WHERE status_flag=TRUE ORDER BY username ASC"
    Call openConnection
    rsUsers.CursorLocation = adUseClient
    rsUsers.Open strquery, conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rsUsers.ActiveConnection = Nothing
    Call closeConnection
    'show in grid
    Set grdUser.DataSource = rsUsers
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set grdUser.DataSource = Nothing
    rsUsers.Close
    Set rsUsers = Nothing
End Sub

Private Sub refreshUserInfo()
    Call openConnection
    Set rsUsers.ActiveConnection = conn
    rsUsers.Requery
    Set rsUsers.ActiveConnection = Nothing
    Call closeConnection
End Sub

