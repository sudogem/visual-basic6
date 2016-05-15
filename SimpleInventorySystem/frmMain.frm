VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory System"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6015
      Begin TabDlg.SSTab tabs 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   7011
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   -2147483638
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmMain.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Image1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label12"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Contacts"
         TabPicture(1)   =   "frmMain.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "cmdCView"
         Tab(1).Control(2)=   "cmdCAdd"
         Tab(1).Control(3)=   "cmdCDelete"
         Tab(1).Control(4)=   "cmdCEdit"
         Tab(1).Control(5)=   "cmdCSave"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Products"
         TabPicture(2)   =   "frmMain.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdPView"
         Tab(2).Control(1)=   "Frame4"
         Tab(2).Control(2)=   "cmdPAdd"
         Tab(2).Control(3)=   "cmdPDelete"
         Tab(2).Control(4)=   "cmdPEdit"
         Tab(2).Control(5)=   "cmdPSave"
         Tab(2).ControlCount=   6
         Begin VB.CommandButton cmdCSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   -70680
            TabIndex        =   38
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdPSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   -70680
            TabIndex        =   37
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdPEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -70680
            TabIndex        =   36
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdPDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   35
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdPAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   -73560
            TabIndex        =   34
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdCEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -70680
            TabIndex        =   33
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdCDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72000
            TabIndex        =   32
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdCAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   -73560
            TabIndex        =   31
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmdCView 
            Caption         =   "View"
            Height          =   375
            Left            =   -74880
            TabIndex        =   10
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Frame Frame4 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   8
            Top             =   360
            Width           =   5535
            Begin VB.TextBox txtDescription 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txtDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   1320
               Width           =   2295
            End
            Begin VB.TextBox txtTotalPrice 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox txtQuantity 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox txtUnitPrice 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   240
               Width           =   2295
            End
            Begin MSComctlLib.ListView lvwProducts 
               Height          =   1935
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   3413
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   12648447
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
            Begin VB.Label Label11 
               Caption         =   "Description:"
               Height          =   255
               Left            =   1680
               TabIndex        =   30
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label10 
               Caption         =   "Date Purchased:"
               Height          =   255
               Left            =   1680
               TabIndex        =   29
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "TotalPrice:"
               Height          =   255
               Left            =   1680
               TabIndex        =   28
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label8 
               Caption         =   "Quantity:"
               Height          =   255
               Left            =   1680
               TabIndex        =   27
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label7 
               Caption         =   "UnitPrice:"
               Height          =   255
               Left            =   1680
               TabIndex        =   26
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdPView 
            Caption         =   "View"
            Height          =   375
            Left            =   -74880
            TabIndex        =   7
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   5
            Top             =   360
            Width           =   5535
            Begin VB.TextBox txtEmail 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   1680
               Width           =   2535
            End
            Begin VB.TextBox txtCCel 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txtCPhone 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox txtCAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtCName 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   240
               Width           =   2535
            End
            Begin MSComctlLib.ListView lvwContacts 
               Height          =   1935
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   3413
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   12648447
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               NumItems        =   0
            End
            Begin VB.Label Label6 
               Caption         =   "Email:"
               Height          =   255
               Left            =   1800
               TabIndex        =   20
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "Cel. #."
               Height          =   255
               Left            =   1800
               TabIndex        =   19
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "Phone #."
               Height          =   255
               Left            =   1800
               TabIndex        =   18
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Address:"
               Height          =   255
               Left            =   1800
               TabIndex        =   17
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Name:"
               Height          =   255
               Left            =   1800
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Label Label12 
            Caption         =   $"frmMain.frx":0054
            Height          =   735
            Left            =   240
            TabIndex        =   39
            Top             =   2880
            Width           =   5295
         End
         Begin VB.Image Image1 
            Height          =   2055
            Left            =   240
            Picture         =   "frmMain.frx":0127
            Top             =   600
            Width           =   5295
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inventory System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1440
         TabIndex        =   1
         Top             =   210
         Width           =   3045
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim lstmain1 As ListItem
Dim lstmain2 As ListItem
Dim clmhead1 As ColumnHeader
Dim clmhead2 As ColumnHeader

Private Sub cmdCEdit_Click()
cmdCSave.Visible = True
        txtCName.Locked = False
        txtCAddress.Locked = False
        txtCPhone.Locked = False
        txtCCel.Locked = False
        txtEmail.Locked = False
        txtCName.SetFocus
End Sub

Private Sub cmdCSave_Click()
Set rs = db.OpenRecordset("Contacts")
cmdCSave.Visible = False
        txtCName.Locked = True
        txtCAddress.Locked = True
        txtCPhone.Locked = True
        txtCCel.Locked = True
        txtEmail.Locked = True
        
        While Not rs.EOF
        If lvwContacts.SelectedItem.Text = rs!ContactID Then
            rs.Edit
            rs!Name = txtCName.Text
            rs!Address = txtCAddress.Text
            rs!PhoneNo = txtCPhone.Text
            rs!Celno = txtCCel.Text
            rs!Email = txtEmail
            rs.Update
            rs.MoveLast
            rs.MoveNext
        Else
            rs.MoveNext
        End If
        Wend
End Sub

Private Sub cmdPAdd_Click()
frmAddProducts.Show vbModal
End Sub

Private Sub cmdCAdd_Click()
frmAddContacts.Show vbModal
End Sub

Private Sub cmdPEdit_Click()
cmdPSave.Visible = True
            txtUnitPrice.Locked = False
            txtQuantity.Locked = False
            txtTotalPrice.Locked = False
            txtDescription.Locked = False
            txtDate.Locked = True
            txtUnitPrice.SetFocus
End Sub

Private Sub cmdPSave_Click()
cmdPSave.Visible = False
            txtUnitPrice.Locked = True
            txtQuantity.Locked = True
            txtTotalPrice.Locked = True
            txtDescription.Locked = True
            txtDate.Locked = True
            
End Sub

Private Sub Form_Load()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\datas.mdb")
Set clmhead1 = lvwContacts.ColumnHeaders.Add(, , "Contact ID", lvwContacts.Width)
Set clmhead2 = lvwProducts.ColumnHeaders.Add(, , "Product ID", lvwProducts.Width)
ContactView
ProductView
End Sub

Private Sub cmdCView_Click()
ContactView
rs.Close
End Sub
Private Sub cmdPView_Click()
ProductView
rs.Close
End Sub


Private Sub lvwContacts_Click()
Set rs = db.OpenRecordset("Contacts")
While Not rs.EOF
If lvwContacts.SelectedItem.Text = rs!ContactID Then
    txtCName.Text = rs!Name
    txtCAddress.Text = rs!Address
    txtCPhone.Text = rs!PhoneNo
    txtCCel.Text = rs!Celno
    txtEmail = rs!Email
    rs.MoveLast
    rs.MoveNext
Else
    rs.MoveNext
End If
Wend

End Sub

Private Sub lvwProducts_Click()
Set rs = db.OpenRecordset("Products")
While Not rs.EOF
If lvwProducts.SelectedItem.Text = rs!ProductID Then
    txtUnitPrice.Text = rs!UnitPrice
    txtQuantity.Text = rs!Quantity
    txtTotalPrice.Text = rs!TotalPrice
    txtDescription.Text = rs!Description
    txtDate = rs!Date
    rs.MoveLast
    rs.MoveNext
Else
    rs.MoveNext
End If
Wend
End Sub

Public Sub cmdCDelete_Click()
Set rs = db.OpenRecordset("Contacts")
        txtCName.Text = ""
        txtCAddress.Text = ""
        txtCPhone.Text = ""
        txtCCel.Text = ""
        txtEmail.Text = ""
    cmdCDelete.Enabled = False
    cmdCEdit.Enabled = False
While Not rs.EOF
    If frmMain.lvwContacts.SelectedItem.Text = rs!ContactID Then
    rs.Delete
    ContactView
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            rs.MoveLast
        End If
    rs.MoveNext
    Else
    rs.MoveNext
    End If
Wend
rs.Close
End Sub

Public Sub cmdPDelete_Click()
Set rs = db.OpenRecordset("Products")
            txtUnitPrice.Text = ""
            txtQuantity.Text = ""
            txtTotalPrice.Text = ""
            txtDescription.Text = ""
            txtDate.Text = ""
While Not rs.EOF
    If frmMain.lvwProducts.SelectedItem.Text = rs!ProductID Then
    rs.Delete
    cmdPDelete.Enabled = False
    cmdPEdit.Enabled = False
    ProductView
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            rs.MoveLast
        End If
    rs.MoveNext
    Else
    rs.MoveNext
    End If
Wend
rs.Close
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub cmdClose_Click()
End
End Sub

Public Sub ProductView()
Set rs = db.OpenRecordset("Products")
frmMain.lvwProducts.ListItems.Clear
    While Not rs.EOF
    Set lstmain2 = frmMain.lvwProducts.ListItems.Add(, , rs!ProductID)
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        frmMain.lvwProducts.Enabled = True
        frmMain.cmdPDelete.Enabled = True
        frmMain.cmdPEdit.Enabled = True
    End If
End Sub

Public Sub ContactView()
Set rs = db.OpenRecordset("Contacts")
frmMain.lvwContacts.ListItems.Clear
    While Not rs.EOF
    Set lstmain1 = frmMain.lvwContacts.ListItems.Add(, , rs!ContactID)
    rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        MsgBox "There are no records on the database"
    Else
        frmMain.lvwContacts.Enabled = True
        frmMain.cmdCDelete.Enabled = True
        frmMain.cmdCEdit.Enabled = True
    End If
End Sub



