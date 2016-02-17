VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form formEditUser 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.ComboBox regAccountType 
      Height          =   315
      ItemData        =   "formEditUser.frx":0000
      Left            =   9780
      List            =   "formEditUser.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   3900
      Width           =   1335
   End
   Begin VB.TextBox regUserName 
      Height          =   285
      Left            =   9780
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox regPassword 
      Height          =   285
      Left            =   9780
      TabIndex        =   1
      Top             =   1950
      Width           =   2415
   End
   Begin VB.TextBox regFirstName 
      Height          =   285
      Left            =   9780
      TabIndex        =   2
      Top             =   2340
      Width           =   2415
   End
   Begin VB.TextBox regLastName 
      Height          =   285
      Left            =   9780
      TabIndex        =   3
      Top             =   2730
      Width           =   2415
   End
   Begin VB.TextBox regAddress 
      Height          =   285
      Left            =   9780
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox regContactNumber 
      Height          =   285
      Left            =   9780
      TabIndex        =   5
      Top             =   3510
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update User"
      Height          =   345
      Left            =   9780
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdUser 
      Height          =   3225
      Left            =   180
      TabIndex        =   8
      Top             =   1560
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5689
      _Version        =   393216
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select (double click) a user below  to be edited :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   180
      TabIndex        =   27
      Top             =   1320
      Width           =   5730
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
      Left            =   7950
      TabIndex        =   26
      Top             =   1620
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
      Left            =   7950
      TabIndex        =   25
      Top             =   1980
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
      Left            =   7950
      TabIndex        =   24
      Top             =   2760
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
      Left            =   7950
      TabIndex        =   23
      Top             =   2370
      Width           =   1185
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
      Left            =   7950
      TabIndex        =   22
      Top             =   3150
      Width           =   1185
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
      Left            =   7950
      TabIndex        =   21
      Top             =   3540
      Width           =   1665
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
      Left            =   7950
      TabIndex        =   20
      Top             =   3960
      Width           =   1635
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
      Left            =   9540
      TabIndex        =   19
      Top             =   1590
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
      Left            =   9540
      TabIndex        =   18
      Top             =   1950
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
      Left            =   9540
      TabIndex        =   17
      Top             =   2340
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
      Left            =   9540
      TabIndex        =   16
      Top             =   2730
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
      Left            =   9540
      TabIndex        =   15
      Top             =   3120
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
      Left            =   9540
      TabIndex        =   14
      Top             =   3510
      Width           =   195
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
      Left            =   9540
      TabIndex        =   13
      Top             =   3930
      Width           =   195
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
      Left            =   9780
      TabIndex        =   12
      Top             =   870
      Width           =   1335
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
      Left            =   9210
      TabIndex        =   11
      Top             =   870
      Width           =   585
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
      Left            =   3090
      TabIndex        =   10
      Top             =   840
      Width           =   615
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
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   435
   End
End
Attribute VB_Name = "formEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsers As New ADODB.Recordset
Private userid As Long

Private Sub cmdUpdate_Click()
    If userid = -1 Then
        Exit Sub
    End If

    Dim result As Boolean
    openConnection
    result = EditUser(conn, userid, regUserName.Text, regPassword.Text, regFirstName.Text, regLastName.Text, regAddress.Text, regContactNumber.Text, regAccountType.Text)
    closeConnection
    If result = True Then
        Call resetEntries
        Call refreshUserInfo
    Else
        formError.errorLabel.Caption = "Cannot edit user"
        formError.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    userid = -1
    Call getUsers
End Sub

Private Sub getUsers()
    Dim strquery As String
    strquery = "SELECT userid, username, utype FROM users WHERE status_flag=TRUE ORDER BY username ASC"
    Call openConnection
    rsUsers.CursorLocation = adUseClient
    rsUsers.Open strquery, conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rsUsers.ActiveConnection = Nothing
    Call closeConnection
    'show in grid
    Set grdUser.DataSource = rsUsers
    grdUser.Columns("userid").Visible = False
End Sub

Private Sub refreshUserInfo()
    Call openConnection
    Set rsUsers.ActiveConnection = conn
    rsUsers.Requery
    Set rsUsers.ActiveConnection = Nothing
    grdUser.Columns("userid").Visible = False
    Call closeConnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set grdUser.DataSource = Nothing
    rsUsers.Close
    Set rsUsers = Nothing
End Sub

Private Sub getUserDetailInfo(ByVal id As Long)
    Dim rstemp As New ADODB.Recordset
    Dim strquery As String
    
    Call openConnection
    Set rstemp = getUserInfo(conn, id)
    If Not rstemp.EOF Then
        userid = id
        regUserName.Text = rstemp!username & ""
        regPassword.Text = rstemp!password & ""
        regFirstName.Text = rstemp!firstname & ""
        regLastName.Text = rstemp!lastname & ""
        regAddress.Text = rstemp!address & ""
        regContactNumber.Text = rstemp![contact_number] & ""
        regAccountType.Text = rstemp!utype & ""
    Else
        Call resetEntries
    End If
    rstemp.Close
    Call closeConnection
    Set rstemp = Nothing
End Sub

Private Sub resetEntries()
    userid = -1
    regUserName.Text = ""
    regPassword.Text = ""
    regFirstName.Text = ""
    regLastName.Text = ""
    regAddress.Text = ""
    regContactNumber.Text = ""
    regAccountType.Text = ""
End Sub

Private Sub grdUser_Click()
    If Not (rsUsers.EOF Or rsUsers.EOF) Then
        Call getUserDetailInfo(CLng(rsUsers!userid & ""))
    End If
End Sub

