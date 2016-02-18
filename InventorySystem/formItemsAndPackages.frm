VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formAddItemsAndPackages 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   10005
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "PACKAGES"
      Height          =   4575
      Left            =   630
      TabIndex        =   27
      Top             =   3900
      Width           =   8865
      Begin VB.TextBox regPackageName 
         Height          =   285
         Left            =   2100
         TabIndex        =   8
         Top             =   1530
         Width           =   2625
      End
      Begin VB.TextBox regPackageQty 
         Height          =   285
         Left            =   2100
         TabIndex        =   7
         Top             =   1140
         Width           =   945
      End
      Begin VB.TextBox regPackagePrice 
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   1920
         Width           =   1185
      End
      Begin VB.CommandButton cmdAddPackage 
         Caption         =   "ADD PACKAGE"
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
         Left            =   2100
         MouseIcon       =   "formItemsAndPackages.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2310
         Width           =   1575
      End
      Begin VB.TextBox regPackageID 
         Height          =   285
         Left            =   2100
         TabIndex        =   6
         Top             =   750
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid grdPackages 
         Height          =   1305
         Left            =   300
         TabIndex        =   28
         Top             =   3150
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2302
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.Line Line1 
         BorderColor     =   &H000066FF&
         BorderStyle     =   3  'Dot
         X1              =   5970
         X2              =   8190
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lblItemID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MyItemNum101"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5490
         TabIndex        =   42
         Top             =   150
         Width           =   3285
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1170
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         ForeColor       =   &H000066FF&
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
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
         Left            =   540
         TabIndex        =   39
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label5 
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
         Left            =   1890
         TabIndex        =   38
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKAGE PRICE"
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
         Left            =   540
         TabIndex        =   37
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Left            =   1860
         TabIndex        =   36
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H000066FF&
         Height          =   285
         Left            =   270
         TabIndex        =   35
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Top             =   2790
         Width           =   465
      End
      Begin VB.Label Label22 
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
         Left            =   1890
         TabIndex        =   33
         Top             =   1530
         Width           =   255
      End
      Begin VB.Label Label23 
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
         Left            =   1890
         TabIndex        =   32
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKAGE NAME"
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
         Left            =   540
         TabIndex        =   31
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKAGE ID"
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
         Left            =   540
         TabIndex        =   30
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000066FF&
         Height          =   255
         Left            =   6720
         TabIndex        =   29
         Top             =   570
         Width           =   705
      End
   End
   Begin VB.TextBox regitemqty 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   2460
      Width           =   1245
   End
   Begin VB.TextBox regitemprice 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   2850
      Width           =   1245
   End
   Begin VB.CommandButton cmdAddItem 
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
      Left            =   2400
      MouseIcon       =   "formItemsAndPackages.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3270
      Width           =   1245
   End
   Begin VB.TextBox regitemdesc 
      Height          =   825
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1530
      Width           =   3255
   End
   Begin VB.TextBox regitemtype 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1140
      Width           =   2565
   End
   Begin VB.TextBox regitemid 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   750
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid grdItems 
      Height          =   2895
      Left            =   6090
      TabIndex        =   11
      Top             =   870
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Label Label30 
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
      Left            =   1170
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label29 
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
      Left            =   600
      TabIndex        =   25
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM QUANTITY"
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
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label27 
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
      Left            =   2190
      TabIndex        =   23
      Top             =   2490
      Width           =   195
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM PRICE"
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
      Top             =   2910
      Width           =   1545
   End
   Begin VB.Label Label24 
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
      Left            =   2190
      TabIndex        =   21
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
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
      Left            =   6090
      TabIndex        =   20
      Top             =   90
      Width           =   585
   End
   Begin VB.Label Label20 
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
      Left            =   6750
      TabIndex        =   19
      Top             =   90
      Width           =   465
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select (double click) an item below to add a package:"
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
      Height          =   375
      Left            =   6090
      TabIndex        =   18
      Top             =   510
      Width           =   3345
   End
   Begin VB.Label Label17 
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
      Left            =   2190
      TabIndex        =   17
      Top             =   1500
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
      Left            =   2190
      TabIndex        =   16
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM DESCRIPTION"
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
      Top             =   1530
      Width           =   1545
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
      Left            =   2190
      TabIndex        =   14
      Top             =   780
      Width           =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM TYPE"
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
      TabIndex        =   13
      Top             =   1170
      Width           =   1545
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM ID"
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
      TabIndex        =   12
      Top             =   780
      Width           =   1545
   End
End
Attribute VB_Name = "formAddItemsAndPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents rsItems As ADODB.Recordset
Attribute rsItems.VB_VarHelpID = -1
Private rsItemsTemp As New ADODB.Recordset
Private rsPackages As New ADODB.Recordset
Private itemid As String
Private allowGetData As Boolean

Private Sub cmdAddItem_Click()
On Error Resume Next
    Dim result As Boolean
    Dim qty As Long
    Dim price As Double
    qty = CLng(regItemQTY.Text)
    If Err.Number <> 0 Then
        Err.Clear
        formError.errorLabel.Caption = "Invalid Number Format"
        formError.Show vbModal
        Exit Sub
    End If
    
    price = CDbl(regItemPrice.Text)
    If Err.Number <> 0 Then
        Err.Clear
        formError.errorLabel.Caption = "Invalid Number Format"
        formError.Show vbModal
        Exit Sub
    End If
    
    Call openConnection
    result = addItem(conn, regItemID.Text, regItemType.Text, regItemDesc.Text, qty, price)
    Call closeConnection
    
    If result = False Then
        formError.errorLabel.Caption = "Sorry, Cannot add Item because 'Item ID' already exist."
        formError.Show vbModal
    Else
        Call resetEntries
        Call refreshItemsInfo
    End If
End Sub

Private Sub cmdAddPackage_Click()
On Error Resume Next
    Dim result As Boolean
    Dim qty As Long
    Dim price As Double
        
    If Trim$(regPackageID.Text) = "" Or Trim(regPackageName.Text) = "" Then
        formError.errorLabel.Caption = "You need to fill up PackageID and PackageName"
        Exit Sub
    End If
    
    qty = CLng(regPackageQTY.Text)
    If Err.Number <> 0 Then
        Err.Clear
        formError.errorLabel.Caption = "Invalid Qty Entry"
        formError.Show vbModal
        Exit Sub
    End If
    price = CDbl(regPackagePrice.Text)
    If Err.Number <> 0 Then
        Err.Clear
        formError.errorLabel.Caption = "Invalid Price Entry"
        formError.Show vbModal
        Exit Sub
    End If
    
    Call openConnection
    result = addPackage(conn, Trim(regPackageID.Text), lblItemID.Caption, Trim(regPackageName.Text), qty, price)
    Call closeConnection
    
    If result = False Then
        formError.errorLabel.Caption = "Cannot Add Package"
        formError.Show vbModal
        Exit Sub
    Else
        resetEntries
        Call showPackages
    End If
    
End Sub

Private Sub Form_Load()
    Set rsItems = rsItemsTemp
    Call getItems
End Sub

Private Sub getItems()
    Dim strquery As String
    strquery = "SELECT itemid as ItemID, itype as ItemType, desc as Description, qty, price FROM item"
    allowGetData = False
    Call openConnection
    rsItems.CursorLocation = adUseClient
    rsItems.Open strquery, conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rsItems.ActiveConnection = Nothing
    Call closeConnection
    Set grdItems.DataSource = rsItems
    allowGetData = True
    Call showPackages
End Sub

Private Sub getPackages(id As String)
    Dim strquery As String
    Set grdPackages.DataSource = Nothing
    
    If rsPackages.State = adStateOpen Then
        rsPackages.Close
    End If
    strquery = "SELECT packageid as PackageID, name as Name, qty, price FROM package WHERE itemid='" & id & "'"
    Call openConnection
    rsPackages.CursorLocation = adUseClient
    rsPackages.Open strquery, conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rsPackages.ActiveConnection = Nothing
    Call closeConnection
    Set grdPackages.DataSource = rsPackages
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set grdItems.DataSource = Nothing
    rsItems.Close
    Set rsItems = Nothing
    rsPackages.Close
    Set rsPackages = Nothing
End Sub

Private Sub refreshItemsInfo()
    allowGetData = False
    Call openConnection
    Set rsItems.ActiveConnection = conn
    rsItems.Requery
    Set rsItems.ActiveConnection = Nothing
    Call closeConnection
    allowGetData = True
    Call showPackages
End Sub

Private Sub resetEntries()
    itemid = ""
    regItemID.Text = ""
    regItemType.Text = ""
    regItemDesc.Text = ""
    regItemQTY.Text = ""
    regItemPrice.Text = ""
    
    regPackageID.Text = ""
    regPackageName.Text = ""
    regPackagePrice.Text = ""
    regPackageQTY.Text = ""
End Sub

Private Sub rsItems_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Call showPackages
End Sub

Private Sub showPackages()
    If allowGetData = True Then
        If rsItems.EOF = True Or rsItems.BOF = True Then
            'blank dayon ang packages
            lblItemID.Caption = ""
            cmdAddPackage.Enabled = False
            Call getPackages("")
        Else
            lblItemID.Caption = rsItems!itemid & ""
            cmdAddPackage.Enabled = True
            Call getPackages(rsItems!itemid & "")
        End If
    End If
End Sub
