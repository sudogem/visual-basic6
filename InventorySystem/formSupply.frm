VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form formSupply 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12450
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo cboItems 
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "SAVE INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MouseIcon       =   "formSupply.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   2595
   End
   Begin VB.CommandButton supplyDeleteItem 
      Caption         =   "DELETE ITEM"
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
      Left            =   1980
      MouseIcon       =   "formSupply.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4650
      Width           =   1425
   End
   Begin VB.CommandButton supplyAddItem 
      Caption         =   "ADD ITEM"
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
      Left            =   570
      MouseIcon       =   "formSupply.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4650
      Width           =   1245
   End
   Begin VB.TextBox regCost 
      Height          =   285
      Left            =   1770
      TabIndex        =   5
      Top             =   4050
      Width           =   1215
   End
   Begin VB.TextBox regQty 
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Top             =   3690
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdItems 
      Height          =   1785
      Left            =   4770
      TabIndex        =   21
      Top             =   3180
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   3149
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
   Begin VB.TextBox regSupplier 
      Height          =   285
      Left            =   1830
      TabIndex        =   1
      Top             =   1410
      Width           =   3045
   End
   Begin VB.TextBox regContact 
      Height          =   285
      Left            =   1830
      TabIndex        =   2
      Top             =   1800
      Width           =   2115
   End
   Begin MSComCtl2.DTPicker regDate 
      Height          =   315
      Left            =   1830
      TabIndex        =   0
      Top             =   1020
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20250625
      CurrentDate     =   37545
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Pane"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   8940
      TabIndex        =   26
      Top             =   2910
      Width           =   1395
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
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   4050
      Width           =   195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COST"
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
      Left            =   630
      TabIndex        =   24
      Top             =   4080
      Width           =   1005
   End
   Begin VB.Label Label6 
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
      Left            =   1560
      TabIndex        =   23
      Top             =   3690
      Width           =   195
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY "
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
      Left            =   630
      TabIndex        =   22
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Items :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   2
      Left            =   630
      TabIndex        =   20
      Top             =   3210
      Width           =   1365
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose an item from the drop down list below:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   2910
      Width           =   3765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Left            =   630
      TabIndex        =   18
      Top             =   2520
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
      Index           =   1
      Left            =   1170
      TabIndex        =   17
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
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
      Left            =   630
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER"
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
      Left            =   630
      TabIndex        =   15
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
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
      Left            =   630
      TabIndex        =   14
      Top             =   1830
      Width           =   1005
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
      Left            =   1590
      TabIndex        =   13
      Top             =   1050
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
      Left            =   1590
      TabIndex        =   12
      Top             =   1410
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
      Left            =   1590
      TabIndex        =   11
      Top             =   1800
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
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Left            =   630
      TabIndex        =   9
      Top             =   540
      Width           =   885
   End
End
Attribute VB_Name = "formSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As ADODB.Recordset
Dim rsItems As ADODB.Recordset

Private Sub initRS()
    Set rstemp = New ADODB.Recordset
    rstemp.Fields.Append "ItemID", adVarChar, 50
    rstemp.Fields.Append "Qty", adDouble
    rstemp.Fields.Append "Cost", adDouble
    rstemp.Open
    Set grdItems.DataSource = rstemp
End Sub

Private Sub setItemCombo()
    openConnection
    Set rsItems = New ADODB.Recordset
    Set rsItems = getItemCodes(conn)
    closeConnection
    
    Set cboItems.RowSource = rsItems
    cboItems.ListField = "itemid"
End Sub


Private Sub cmdSave_Click()
    Call SubmitInfo
End Sub

Private Sub Form_Load()
    'MsgBox CLng(Now)
    regDate.Value = Now()
    Call initRS
    Call setItemCombo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set grdItems.DataSource = Nothing
    rstemp.Close
    Set rstemp = Nothing
End Sub

Private Sub clearMaster()
    regDate.Value = Now
    regSupplier.Text = ""
    regContact.Text = ""
End Sub

Private Sub clearDetail()
On Error Resume Next
    cboItems.Text = ""
    regQty.Text = ""
    regCost.Text = ""
End Sub

Private Sub clearGrid()
    Set grdItems.DataSource = Nothing
    rstemp.Close
    Set rstemp = Nothing
    initRS
End Sub

Private Sub SubmitInfo()
    Dim id As Double
    Dim result As Boolean
    Dim theDate As String

    'check entries
    'put code here
    If Trim(regSupplier.Text) = "" And Trim(regContact.Text) = "" Then
        formError.errorLabel.Caption = "Missing or invalid Data"
        formError.Show vbModal
        Exit Sub
    End If
    
    'check that it has at least one item
    'put code here
    If rstemp.EOF = True And rstemp.BOF = True Then
        formError.errorLabel.Caption = "Must have at least one item!"
        formError.Show vbModal
        Exit Sub
    End If
    
    openConnection
    conn.BeginTrans
    
    id = addInTransactionMaster(conn, regDate.Value, regSupplier.Text, regContact.Text, mUserID)
    If id = -1 Then
        conn.RollbackTrans
        closeConnection
        formError.errorLabel.Caption = "Cannot Add Information!"
        formError.Show vbModal
        Exit Sub
    End If
        
    rstemp.MoveFirst
    While (Not rstemp.EOF)
        result = addInTransactionDetail(conn, id, rstemp!itemid, rstemp!qty, rstemp!cost)
        If result = False Then
            conn.RollbackTrans
            closeConnection
            formError.errorLabel.Caption = "Cannot Add Information!"
            formError.Show vbModal
            Exit Sub
        End If
        rstemp.MoveNext
    Wend
        
    conn.CommitTrans
    closeConnection
    
    
    'clear entry fields
    Call clearMaster
    Call clearDetail
    'clear grid
    Call clearGrid
End Sub




Private Sub supplyAddItem_Click()
On Error GoTo errhandler
    'check entries
    Dim qty As Double
    Dim cost As Double
    
    If Trim(cboItems.Text) = "" Then
        formError.errorLabel.Caption = "please choose an item."
        formError.Show vbModal
        Exit Sub
    End If
        
    qty = CDbl(regQty.Text)
    cost = CDbl(regCost.Text)
    
    rstemp.AddNew Array("itemid", "qty", "cost"), Array(cboItems.Text, qty, cost)
    Call clearDetail
    Exit Sub
errhandler:
    formError.errorLabel.Caption = "Invalid or missing Entry!"
    formError.Show vbModal
End Sub

Private Sub supplyDeleteItem_Click()
On Error Resume Next
    rstemp.Delete
End Sub
