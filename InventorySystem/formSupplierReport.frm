VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formSupplierReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Supplier Report"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdShowReport 
      Caption         =   "&Show Report"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Report Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtReportHeader 
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Myriad Pro"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Text            =   " Daily Supply Report"
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: When the Transaction No. appears, copy it onto the Referror Claim Slip "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         Top             =   7920
         Visible         =   0   'False
         Width           =   8295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   8400
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin MSComCtl2.DTPicker dtpCDFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Pro"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50855937
         CurrentDate     =   37545
      End
      Begin MSComCtl2.DTPicker dtpCDTo 
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Pro"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50855937
         CurrentDate     =   37545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   8400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: When the Transaction No. appears, copy it onto the Referror Claim Slip "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   7920
         Visible         =   0   'False
         Width           =   8295
      End
   End
End
Attribute VB_Name = "formSupplierReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShowReport_Click()
Dim strShape As String
Dim strSql1 As String
Dim strSql2 As String
Dim strTemp As String
Dim intX As Integer
'Dim page As New PageSet.PrinterControl
Dim sConnString As String
    
    sConnString = "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & strPath & ";Data Provider=MICROSOFT.JET.OLEDB.4.0"
   
    devReports.devReports.Open sConnString
    'GENERATE THE SHAPE
        strSql1 = "SELECT in_mas.* FROM in_mas WHERE date>=#" & FormatDateTime(dtpCDFrom.Value, vbShortDate) & "# AND date<=#" & FormatDateTime(dtpCDTo.Value, vbShortDate) & "#"
        strSql2 = "SELECT in_det.* FROM in_det"
        
        strShape = "SHAPE " & _
            "{" & strSql1 & "} AS cmdSupplierReport_Mas APPEND ({" & strSql2 & "}  AS cmdSupplierReport_Det " & _
            "RELATE  'transid' TO 'transid') AS cmdSupplierReport_Det"
    
        devReports.Commands("cmdSupplierReport_Mas").CommandText = strShape
    
    'If Printer.Orientation <> vbPRORPortrait Then
        'SET THE PAGE ORIENTATION
            'page.ChngOrientationPortrait
    'MsgBox strSql1
    drpSupplierReport.Sections(1).Controls("lblReportHeader").Caption = txtReportHeader.Text
    drpSupplierReport.Show vbModal
    
Set devReports = Nothing
End Sub

Private Sub Form_Load()
   dtpCDFrom.Value = Now
   dtpCDTo.Value = Now
End Sub
