VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formDeleteItemsAndPackages 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   16200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   16200
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "PACKAGES"
      Height          =   5115
      Left            =   9840
      TabIndex        =   6
      Top             =   360
      Width           =   9105
      Begin VB.TextBox regPackageName 
         Height          =   285
         Left            =   5730
         TabIndex        =   11
         Top             =   2910
         Width           =   2625
      End
      Begin VB.TextBox regPackageQTY 
         Height          =   285
         Left            =   5730
         TabIndex        =   10
         Top             =   2520
         Width           =   945
      End
      Begin VB.TextBox regPackagePrice 
         Height          =   285
         Left            =   5700
         TabIndex        =   9
         Top             =   3300
         Width           =   1185
      End
      Begin VB.CommandButton cmdDeletePackage 
         Caption         =   "DELETE PACKAGE"
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
         Left            =   5700
         MouseIcon       =   "formDeleteItemsAndPackages.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   3690
         Width           =   1995
      End
      Begin VB.TextBox regPackageID 
         Height          =   285
         Left            =   5730
         TabIndex        =   7
         Top             =   2130
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid grdPackages 
         Height          =   3255
         Left            =   240
         TabIndex        =   12
         Top             =   1710
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5741
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
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5040
         TabIndex        =   27
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000066FF&
         Height          =   285
         Left            =   4140
         TabIndex        =   26
         Top             =   1680
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
         Left            =   4170
         TabIndex        =   25
         Top             =   2550
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
         Left            =   5520
         TabIndex        =   24
         Top             =   2520
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
         Left            =   4170
         TabIndex        =   23
         Top             =   3360
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
         Left            =   5490
         TabIndex        =   22
         Top             =   3330
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Package"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000066FF&
         Height          =   285
         Left            =   210
         TabIndex        =   21
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   900
         Width           =   495
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
         Left            =   5520
         TabIndex        =   19
         Top             =   2910
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
         Left            =   5520
         TabIndex        =   18
         Top             =   2130
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
         Left            =   4170
         TabIndex        =   17
         Top             =   2940
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
         Left            =   4170
         TabIndex        =   16
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Package Information for Item Id: "
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
         Left            =   210
         TabIndex        =   15
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select (double click) a package below to be deleted:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   3435
      End
      Begin VB.Label lblItemID 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM ID 10022"
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
         Height          =   255
         Left            =   2910
         TabIndex        =   13
         Top             =   150
         Width           =   3975
      End
   End
   Begin VB.TextBox regItemQTY 
      Height          =   285
      Left            =   6360
      TabIndex        =   5
      Top             =   3090
      Width           =   1245
   End
   Begin VB.TextBox regItemPrice 
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Top             =   3480
      Width           =   1245
   End
   Begin VB.CommandButton cmdDeleteItem 
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
      Left            =   6360
      MouseIcon       =   "formDeleteItemsAndPackages.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3870
      Width           =   1605
   End
   Begin VB.TextBox regItemDesc 
      Height          =   825
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox regItemType 
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   1770
      Width           =   2565
   End
   Begin VB.TextBox regItemID 
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   1380
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid grdItems 
      Height          =   2835
      Left            =   510
      TabIndex        =   28
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5001
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
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Please take note that if you delete this item then all the packages listed below will also be deleted."
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
      Left            =   4590
      TabIndex        =   44
      Top             =   930
      Width           =   5025
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
      Left            =   5100
      TabIndex        =   43
      Top             =   450
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
      Left            =   4530
      TabIndex        =   42
      Top             =   450
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
      Left            =   4590
      TabIndex        =   41
      Top             =   3150
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
      Left            =   6150
      TabIndex        =   40
      Top             =   3120
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
      Left            =   4590
      TabIndex        =   39
      Top             =   3540
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
      Left            =   6150
      TabIndex        =   38
      Top             =   3510
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
      Left            =   510
      TabIndex        =   37
      Top             =   450
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
      Left            =   1140
      TabIndex        =   36
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select (double click) an item below to be deleted:"
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
      Left            =   510
      TabIndex        =   35
      Top             =   930
      Width           =   3435
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
      Left            =   6150
      TabIndex        =   34
      Top             =   2130
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
      Left            =   6150
      TabIndex        =   33
      Top             =   1770
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
      Left            =   4590
      TabIndex        =   32
      Top             =   2160
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
      Left            =   6150
      TabIndex        =   31
      Top             =   1410
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
      Left            =   4590
      TabIndex        =   30
      Top             =   1800
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
      Left            =   4590
      TabIndex        =   29
      Top             =   1410
      Width           =   1545
   End
End
Attribute VB_Name = "formDeleteItemsAndPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents rsItems As ADODB.Recordset
Attribute rsItems.VB_VarHelpID = -1
Private rsItemsTemp As New ADODB.Recordset
Private rsPackages As New ADODB.Recordset
Private itemid As String
Private packageid As String
Private allowGetData As Boolean

Private Sub cmdDeleteItem_Click()
On Error Resume Next
    Dim result As Boolean
    Dim choice As VbMsgBoxResult
    
    choice = MsgBox("Are you sure you want to delete this item?", vbYesNo, "Confirm Delete")
    If choice <> vbYes Then
        Exit Sub
    End If
    
    Call openConnection
    result = deleteItem(conn, itemid)
    Call closeConnection
    
    If result = False Then
        'formError.errorLabel.Caption = "Cannot delete Item"
        'formError.Show vbModal
        
        MsgBox "Cannot delete item because it is used in other transactions!"
    Else
        Call resetEntries
        Call refreshItemsInfo
    End If
End Sub

Private Sub cmdDeletePackage_Click()
On Error Resume Next
    Dim result As Boolean
    Dim choice As VbMsgBoxResult
    
    choice = MsgBox("Are you sure you want to delete this package?", vbYesNo, "Confirm Delete")
    If choice <> vbYes Then
        Exit Sub
    End If
    
    Call openConnection
    result = deletePackage(conn, packageid)
    Call closeConnection
    
    If result = False Then
        formError.errorLabel.Caption = "Cannot delete Package"
        formError.Show vbModal
    Else
        Call resetEntries
        Call refreshItemsInfo
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

Private Sub putItemDetails()
    If rsItems.BOF = True Or rsItems.EOF = True Then
       'do nothing lang
       Call resetEntries
    Else
       'show details
       itemid = rsItems!itemid & ""
       regItemID.Text = rsItems!itemid & ""
       regItemDesc.Text = rsItems!Description & ""
       regItemQTY.Text = rsItems!qty
       regItemType.Text = rsItems!itemType & ""
       regItemPrice.Text = rsItems!price & ""
    End If
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
    
    packageid = ""
    regPackageID.Text = ""
    regPackageName.Text = ""
    regPackagePrice.Text = ""
    regPackageQTY.Text = ""
End Sub

Private Sub grdItems_DblClick()
     Call putItemDetails
End Sub

Private Sub grdPackages_Click()
    Call putPackageDetails
End Sub

Private Sub putPackageDetails()
    If rsPackages.BOF = True Or rsPackages.EOF = True Then
       'do nothing lang
       Call resetEntries
    Else
       'show details
       packageid = rsPackages!packageid & ""
       regPackageID.Text = packageid
       regPackageName.Text = rsPackages!Name & ""
       regPackageQTY.Text = rsPackages!qty & ""
       regPackagePrice.Text = rsPackages!price & ""
    End If
End Sub

Private Sub rsItems_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Call showPackages
End Sub

Private Sub showPackages()
    If allowGetData = True Then
        If rsItems.EOF = True Or rsItems.BOF = True Then
            'blank dayon ang packages
            lblItemID.Caption = ""
            cmdDeletePackage.Enabled = False
            Call getPackages("")
        Else
            lblItemID.Caption = rsItems!itemid & ""
            cmdDeletePackage.Enabled = True
            Call getPackages(rsItems!itemid & "")
        End If
    End If
End Sub

