VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmToolbar 
   Caption         =   "Tool Bar Display"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "frmToolbar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   5685
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Toolbar tbTest 
      Align           =   1  'Align Top
      Height          =   1650
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   2910
      ButtonWidth     =   1429
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Object.ToolTipText     =   "Add Values"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reporting"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "iNFO"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Quit"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cboAirLine 
         Height          =   315
         ItemData        =   "frmToolbar.frx":164A
         Left            =   360
         List            =   "frmToolbar.frx":165D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtDestination 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtFlightNum 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Air-Line Company"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Flight Number"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   120
      Picture         =   "frmToolbar.frx":1686
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5415
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":92E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":9602
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":991C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":9C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":9F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":A26A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":A584
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmToolbar.frx":A89E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim blnSave As Boolean

Private Sub cmdSave_Click()
'BookNum Destination Comp
  Call validatE
  If blnSave = False Then
    MsgBox "Not enough data", vbExclamation
    txtFlightNum.SetFocus
    Exit Sub
  End If
  With rs
    .Open "SELECT BookNum, Destination, Comp FROM Bookings"
    .AddNew
    !BookNum = Trim(txtFlightNum.Text)
    !Destination = Trim(txtDestination.Text)
    !Comp = cboAirLine.Text
    .Update
    .Close
  End With
  MsgBox "Data saved", vbInformation
  clearText
End Sub
Private Sub clearText()
  Dim ctrl As Control
  For Each ctrl In Controls
    If TypeOf ctrl Is TextBox Then
      ctrl.Text = ""
    End If
  Next ctrl
  Image1.Visible = True
  Frame1.Visible = False
End Sub
Private Sub validatE()
  blnSave = True
  If Trim(txtFlightNum.Text) = "" Then blnSave = False
  If Trim(txtDestination.Text) = "" Then blnSave = False
  If cboAirLine.Text = "" Then blnSave = False
End Sub
Private Sub Form_Load()
  Set cn = New ADODB.Connection
  Set rs = New ADODB.Recordset
  Dim mydatasource As String
  'mydatasource = App.Path & "'AirBookings.mdb'"
  mydatasource = "D:\_outsourceproject\Tps\Toolbar\airbookings.mdb"
  With cn
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=mydatasource;Persist Security Info=False"
    .Open
      End With
  With rs
    .ActiveConnection = cn
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
  End With
End Sub

Private Sub tbTest_ButtonClick(ByVal Button As ComctlLib.Button)
  Image1.Visible = True
  Frame1.Visible = False
  'add
  If Button.Index = 1 Then
    Image1.Visible = False
    Frame1.Visible = True
  End If
  
  
  'delete & quick display
  If (Button.Index = 2) Or (Button.Index = 4) Then
    Dim lfNum As String
    
    lfNum = UCase(InputBox("Please enter the filght number to delete bookings"))
    If Trim(lfNum) = "" Then
      Exit Sub
    End If
    If rs.State = adStateOpen Then
      rs.Close
    End If
    rs.Open "SELECT * FROM Bookings WHERE UCase(BookNum) = '" & lfNum & "'"
    If rs.EOF Then
      rs.Close
      MsgBox "Record not listed", vbInformation
      Exit Sub
    End If
    If Button.Index = 2 Then
      If MsgBox("Do you really want to delete the record for flight " & rs!BookNum & " to " & rs!Destination, vbYesNo + vbQuestion) = vbYes Then
        If rs.State = adStateOpen Then
          rs.Close
        End If
        rs.Open "DELETE * FROM Bookings WHERE UCase(BookNum) = '" & lfNum & "'"
        MsgBox "Record deleted", vbInformation
      End If
    End If
    If Button.Index = 4 Then
      MsgBox "Flight Number " & rs!BookNum & vbCrLf & "Destination " & rs!Destination & vbCrLf & "Air Line " & rs!Comp, vbInformation
      If rs.State = adStateOpen Then
        rs.Close
      End If
    End If
    Exit Sub
  End If
  
  
   If Button.Index = 3 Then
     MsgBox "Under construction"
     Exit Sub
   End If
    
   If Button.Index = 5 Then
     MsgBox "Nice place to add grid display", vbInformation
     Exit Sub
   End If
   
   
  If Button.Index = 6 Then
    frmReadMe.Visible = True
  End If
  If Button.Index = 7 Then
    'clean up
    'close and destroy connections and recordset refrences
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    Unload Me
    Exit Sub
  End If
End Sub
