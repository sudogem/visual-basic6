VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formInventoryUpdate 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   12285
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSet 
      Caption         =   "SET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2310
      MouseIcon       =   "formInventoryUpdate.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1290
      Width           =   1275
   End
   Begin VB.TextBox regLevel 
      Height          =   285
      Left            =   2310
      TabIndex        =   5
      Top             =   900
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid grdItems 
      Height          =   2865
      Left            =   1170
      TabIndex        =   0
      Top             =   2520
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   5054
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   13321
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
            LCID            =   13321
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please re-order or purchase these items."
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1170
      TabIndex        =   9
      Top             =   5430
      Width           =   3165
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "List of items that is below or equal to your reorder level :"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   8
      Top             =   2250
      Width           =   4605
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "New Reorder Level :"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   570
      TabIndex        =   7
      Top             =   930
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder Level"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000066FF&
      Height          =   255
      Left            =   8940
      TabIndex        =   4
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MyItemNum101"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8010
      TabIndex        =   3
      Top             =   180
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000066FF&
      BorderStyle     =   3  'Dot
      X1              =   8490
      X2              =   10710
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Level "
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   390
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2070
      TabIndex        =   1
      Top             =   390
      Width           =   1335
   End
End
Attribute VB_Name = "formInventoryUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdSet_Click()
On Error GoTo errhandler
    Dim level As Double
    Dim result As Boolean
    
    level = CDbl(regLevel.Text)
    If level < 0 Then
        formError.errorLabel.Caption = "level needs to be positive"
        formError.Show vbModal
        Exit Sub
    End If
    
    openConnection
    result = setReorderlevel(conn, level)
    closeConnection
    
    If result = True Then
        Call showReorderLevel
        Call showItemsBelowReorderLevel
        regLevel.Text = ""
    Else
        formError.errorLabel.Caption = "update not successful"
        formError.Show vbModal
    End If
    
    Exit Sub
errhandler:
    formError.errorLabel.Caption = "invalid number!"
    formError.Show vbModal
End Sub

Private Sub Form_Load()
    openConnection
    Call showReorderLevel
    Call showItemsBelowReorderLevel
    closeConnection
End Sub

Private Sub showReorderLevel()
    openConnection
    lblLevel.Caption = getReorderlevel(conn)
    closeConnection
End Sub

Private Sub showItemsBelowReorderLevel()
    Dim strquery As String
    strquery = "SELECT * FROM item WHERE qty<=" & lblLevel.Caption
    Set grdItems.DataSource = Nothing
    
    openConnection
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    rs.CursorLocation = adUseClient
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    closeConnection
    
    Set grdItems.DataSource = rs
End Sub
