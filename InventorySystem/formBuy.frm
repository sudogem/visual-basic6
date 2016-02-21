VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formBuy 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.TextBox regTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Left            =   7950
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5550
      Width           =   2145
   End
   Begin VB.TextBox regContact 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2010
      Width           =   2145
   End
   Begin VB.TextBox regAddress 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1620
      Width           =   5715
   End
   Begin VB.TextBox regBuyer 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1230
      Width           =   2145
   End
   Begin VB.TextBox regQty 
      Height          =   285
      Left            =   1620
      TabIndex        =   5
      Top             =   3930
      Width           =   1215
   End
   Begin VB.TextBox regPrice 
      Height          =   285
      Left            =   1620
      TabIndex        =   6
      Top             =   4290
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddPackage 
      Caption         =   "ADD ITEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   420
      MouseIcon       =   "formBuy.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4890
      Width           =   1245
   End
   Begin VB.CommandButton cmdDeletePackage 
      Caption         =   "DELETE ITEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1860
      MouseIcon       =   "formBuy.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4890
      Width           =   1425
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
      Left            =   4620
      MouseIcon       =   "formBuy.frx":0614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   2595
   End
   Begin MSDataGridLib.DataGrid grdPackages 
      Height          =   1785
      Left            =   4620
      TabIndex        =   10
      Top             =   3390
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   3149
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
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
   Begin MSDataListLib.DataCombo cboPackage 
      Height          =   315
      Left            =   2010
      TabIndex        =   4
      Top             =   3420
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker regDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20905985
      CurrentDate     =   37545
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   8910
      TabIndex        =   31
      Top             =   5280
      Width           =   1245
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
      Left            =   1440
      TabIndex        =   29
      Top             =   2010
      Width           =   195
   End
   Begin VB.Label Label13 
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
      Left            =   480
      TabIndex        =   28
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
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
      Left            =   480
      TabIndex        =   27
      Top             =   420
      Width           =   885
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
      Left            =   1170
      TabIndex        =   26
      Top             =   420
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   25
      Top             =   1620
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
      Left            =   1440
      TabIndex        =   24
      Top             =   1230
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
      Left            =   1440
      TabIndex        =   23
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   22
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BUYER"
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
      Left            =   480
      TabIndex        =   21
      Top             =   1260
      Width           =   885
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
      Left            =   480
      TabIndex        =   20
      Top             =   900
      Width           =   855
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
      Left            =   1470
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Package"
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
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose a package from the drop down list below:"
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
      Left            =   450
      TabIndex        =   17
      Top             =   3150
      Width           =   3765
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Package List :"
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
      Left            =   480
      TabIndex        =   16
      Top             =   3450
      Width           =   1485
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
      Left            =   480
      TabIndex        =   15
      Top             =   3960
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
      Left            =   1410
      TabIndex        =   14
      Top             =   3930
      Width           =   195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SELL PRICE"
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
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   1005
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
      Left            =   1410
      TabIndex        =   12
      Top             =   4290
      Width           =   195
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
      Left            =   8790
      TabIndex        =   11
      Top             =   3150
      Width           =   1395
   End
End
Attribute VB_Name = "formBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As ADODB.Recordset
Dim rsPackages As ADODB.Recordset

Private Sub initRS()
    Set rstemp = New ADODB.Recordset
    rstemp.Fields.Append "PackageID", adVarChar, 50
    rstemp.Fields.Append "Qty", adDouble
    rstemp.Fields.Append "Price", adDouble
    rstemp.Fields.Append "Total", adDouble
    rstemp.Open
    Set grdPackages.DataSource = rstemp
End Sub

Private Function getPrice(pID As String) As Double
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    
    strquery = "SELECT price FROM package WHERE packageid='" & pID & "'"
    openConnection
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, cmdtext
    If rs.BOF = True And rs.EOF = True Then
        getPrice = -1
    Else
        getPrice = rs!price
    End If
    rs.Close
    closeConnection
    Set rs = Nothing
    Exit Function
errhandler:
    Set rs = Nothing
    getPrice = -1
End Function

Private Sub setPackageCombo()
    openConnection
    Set rsPackages = New ADODB.Recordset
    Set rsPackages = getPackageCodes(conn)
    closeConnection
    Set cboPackage.RowSource = rsPackages
    cboPackage.ListField = "packageid"
End Sub

Private Sub cboPackage_Change()
    Dim price As Double
    price = getPrice(cboPackage.Text)
    If price = -1 Then
        regPrice.Text = 0
    Else
        regPrice.Text = price
    End If
End Sub

Private Sub cmdAddPackage_Click()
On Error GoTo errhandler
    'check entries
    Dim qty As Double
    Dim price As Double
    Dim result As Boolean
'    Dim itemqty As Double
'    Dim packageitemqty As Double
'    Dim itemid As String
    
    If Trim(cboPackage.Text) = "" Then
        formError.errorLabel.Caption = "please choose an item."
        formError.Show vbModal
        Exit Sub
    End If
    
    qty = CDbl(regQty.Text)
    openConnection
    result = checkItemLevelIfPwedeOutAni(conn, cboPackage.Text, qty)
    closeConnection
    If result = False Then
        'dili pwede!!!
        formError.errorLabel.Caption = "items does not reached required level"
        formError.Show vbModal
        Exit Sub
    End If
    
    price = CDbl(regPrice.Text)
    
    rstemp.AddNew Array("packageid", "qty", "price", "total"), Array(cboPackage.Text, qty, price, price * qty)
    Call clearDetail
    Call showTotal
    Exit Sub
errhandler:
    formError.errorLabel.Caption = "Invalid or missing Entry!"
    formError.Show vbModal
End Sub

Private Sub cmdDeletePackage_Click()
On Error Resume Next
    rstemp.Delete
    Call showTotal
End Sub

Private Sub cmdSave_Click()
    Call SubmitInfo
End Sub

Private Sub Form_Load()
    'MsgBox CLng(Now)
    regDate.Value = Now()
    Call initRS
    Call setPackageCombo
    Call showTotal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set grdPackages.DataSource = Nothing
    rstemp.Close
    Set rstemp = Nothing
End Sub

Private Sub clearMaster()
    regDate.Value = Now()
    regBuyer.Text = ""
    regAddress.Text = ""
    regContact.Text = ""
End Sub

Private Sub clearDetail()
On Error Resume Next
    cboPackage.Text = ""
    regQty.Text = ""
    regPrice.Text = ""
End Sub

Private Sub clearGrid()
    Set grdPackages.DataSource = Nothing
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
    If Trim(regBuyer.Text) = "" And Trim(regAddress.Text) = "" And Trim(regContact.Text) = "" Then
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
    
    id = addOutTransactionMaster(conn, regDate.Value, regBuyer.Text, regAddress.Text, regContact.Text, Val(regTotal.Text), mUserID)
    If id = -1 Then
        conn.RollbackTrans
        closeConnection
        formError.errorLabel.Caption = "Cannot Add Information!"
        formError.Show vbModal
        Exit Sub
    End If
        
    rstemp.MoveFirst
    While (Not rstemp.EOF)
        result = addOutTransactionDetail(conn, id, rstemp!packageid, rstemp!qty, rstemp!price)
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
    
    result = checkLevels(conn)
    If result = True Then
        formError.errorLabel.Caption = "there are items that reached reorderlevel!"
        formError.Show vbModal
    End If
    
    closeConnection
    
    If (MsgBox("Do you want to print receipt?", vbYesNo) = vbYes) Then
        Call showBuyerReport(id)
    End If
        
    'clear entry fields
    Call clearMaster
    Call clearDetail
    'clear grid
    Call clearGrid
End Sub

Private Sub showBuyerReport(ByVal id As Double)
Dim strShape As String
Dim strSql1 As String
Dim strSql2 As String
'Dim page As New PageSet.PrinterControl
Dim sConnString As String
    
    sConnString = "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & strPath & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
   
    devReports.devReports.Open sConnString
    'GENERATE THE SHAPE
        strSql1 = "SELECT out_mas.* FROM out_mas WHERE transid=" & id
        strSql2 = "SELECT out_det.* FROM out_det"
        
        strShape = "SHAPE " & _
            "{" & strSql1 & "} AS cmdSalesReport_Mas APPEND ({" & strSql2 & "}  AS cmdSalesReport_Det " & _
            "RELATE  'transid' TO 'transid') AS cmdSalesReport_Det"
    
        devReports.Commands("cmdSalesReport_Mas").CommandText = strShape
    
    'If Printer.Orientation <> vbPRORPortrait Then
        'SET THE PAGE ORIENTATION
            'page.ChngOrientationPortrait
            
    drpSalesReport.Sections(1).Controls("lblReportHeader").Caption = "Sales Receipt"
    drpSalesReport.Show vbModal
    
Set devReports = Nothing
End Sub

Private Sub showTotal()
    Dim total As Double
    total = 0
    If rstemp.BOF = True And rstemp.EOF = True Then
        'total = 0  do nothing
    Else
        rstemp.MoveFirst
        While Not rstemp.EOF
            total = total + rstemp!total
            rstemp.MoveNext
        Wend
    End If
    regTotal.Text = total
End Sub

Private Sub supplyDeleteItem_Click()
On Error Resume Next
    rstemp.Delete
End Sub

