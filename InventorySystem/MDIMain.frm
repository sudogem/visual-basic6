VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "nventory System"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13605
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList submenu_ImageList1 
      Left            =   12450
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   104
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0418
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox contanerMiddle 
      Align           =   1  'Align Top
      BackColor       =   &H000066FF&
      BorderStyle     =   0  'None
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   13605
      TabIndex        =   2
      Top             =   2925
      Visible         =   0   'False
      Width           =   13605
   End
   Begin VB.PictureBox containerBottom 
      Align           =   1  'Align Top
      BackColor       =   &H005A5A5A&
      BorderStyle     =   0  'None
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   13605
      TabIndex        =   1
      Top             =   2970
      Visible         =   0   'False
      Width           =   13605
   End
   Begin VB.PictureBox containerTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      FillColor       =   &H002F3540&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   13605
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13605
      Begin MSComctlLib.ImageList submenu_ImageList10 
         Left            =   11310
         Top             =   630
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   150
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":0BA7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1183
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList9 
         Left            =   11310
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   150
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1B50
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":20EA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList8 
         Left            =   11310
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   110
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":2A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":2EB1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList7 
         Left            =   11880
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   110
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":364B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":3AAE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox ReportsContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2160
         Picture         =   "MDIMain.frx":4298
         ScaleHeight     =   405
         ScaleWidth      =   10995
         TabIndex        =   9
         Top             =   2490
         Visible         =   0   'False
         Width           =   10995
         Begin VB.Image SupplyReports 
            Height          =   375
            Left            =   6450
            MouseIcon       =   "MDIMain.frx":486E
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":4B78
            Top             =   0
            Width           =   1650
         End
         Begin VB.Image DailySales 
            Height          =   375
            Left            =   4800
            MouseIcon       =   "MDIMain.frx":4F81
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":528B
            Top             =   0
            Width           =   1650
         End
      End
      Begin MSComctlLib.ImageList submenu_ImageList6 
         Left            =   11880
         Top             =   630
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   110
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":56DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":5BBE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList5 
         Left            =   11880
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   104
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":641B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":681F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox InventorySubmenuContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   2220
         Picture         =   "MDIMain.frx":6F9E
         ScaleHeight     =   435
         ScaleWidth      =   11205
         TabIndex        =   8
         Top             =   2490
         Visible         =   0   'False
         Width           =   11205
         Begin VB.Image DeleteItemsAndPackages 
            Height          =   375
            Left            =   4530
            MouseIcon       =   "MDIMain.frx":7574
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":787E
            Top             =   0
            Width           =   2250
         End
         Begin VB.Image EditItemsAndPackages 
            Height          =   375
            Left            =   2250
            MouseIcon       =   "MDIMain.frx":7E4A
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":8154
            Top             =   0
            Width           =   2250
         End
         Begin VB.Image InventoryUpdate 
            Height          =   375
            Left            =   8310
            MouseIcon       =   "MDIMain.frx":86DE
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":89E8
            Top             =   0
            Width           =   1650
         End
         Begin VB.Image Supply 
            Height          =   375
            Left            =   6840
            MouseIcon       =   "MDIMain.frx":8EB8
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":91C2
            Top             =   0
            Width           =   1560
         End
         Begin VB.Image AddItemsAndPackages 
            Height          =   375
            Left            =   240
            MouseIcon       =   "MDIMain.frx":95B6
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":98C0
            Top             =   0
            Width           =   2250
         End
      End
      Begin MSComctlLib.ImageList submenu_ImageList4 
         Left            =   12450
         Top             =   1770
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   150
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":9E6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":A428
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList3 
         Left            =   12450
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   104
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":ADF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":B225
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList submenu_ImageList2 
         Left            =   12450
         Top             =   630
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   104
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":B9D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":BDD2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox UsersContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2250
         ScaleHeight     =   345
         ScaleWidth      =   11325
         TabIndex        =   7
         Top             =   2490
         Visible         =   0   'False
         Width           =   11325
         Begin VB.Image DeleteUser 
            Height          =   375
            Left            =   3840
            MouseIcon       =   "MDIMain.frx":C542
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":C84C
            Top             =   0
            Width           =   1560
         End
         Begin VB.Image EditUser 
            Height          =   375
            Left            =   2340
            MouseIcon       =   "MDIMain.frx":CC71
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":CF7B
            Top             =   0
            Width           =   1560
         End
         Begin VB.Image AddUser 
            Height          =   375
            Left            =   810
            MouseIcon       =   "MDIMain.frx":D36A
            MousePointer    =   99  'Custom
            Picture         =   "MDIMain.frx":D674
            Top             =   0
            Width           =   1560
         End
         Begin VB.Image Image2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   -30
            Picture         =   "MDIMain.frx":DA7C
            Top             =   0
            Width           =   10095
         End
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   13020
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   125
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":E052
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":10AFF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   13020
         Top             =   630
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   125
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1385C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1696A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   13020
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   125
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":19EC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1CE65
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   13020
         Top             =   1770
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   125
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1FFC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":22B18
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "}"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12210
         TabIndex        =   12
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "{"
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
         Left            =   11190
         TabIndex        =   11
         Top             =   2040
         Width           =   105
      End
      Begin VB.Label logout 
         BackStyle       =   0  'Transparent
         Caption         =   "LOGOUT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   11310
         MouseIcon       =   "MDIMain.frx":25C20
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2070
         Width           =   915
      End
      Begin VB.Image DefaultSubMenuContainer 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2220
         Picture         =   "MDIMain.frx":25F2A
         Top             =   2490
         Width           =   10095
      End
      Begin VB.Image BuyMenu 
         Height          =   2070
         Left            =   9030
         MouseIcon       =   "MDIMain.frx":26500
         MousePointer    =   99  'Custom
         Picture         =   "MDIMain.frx":2680A
         Top             =   300
         Width           =   1875
      End
      Begin VB.Image ReportsMenu 
         Height          =   2070
         Left            =   6990
         MouseIcon       =   "MDIMain.frx":292A7
         MousePointer    =   99  'Custom
         Picture         =   "MDIMain.frx":295B1
         Top             =   300
         Width           =   1875
      End
      Begin VB.Image InventoryMenu 
         Height          =   2070
         Left            =   4980
         MouseIcon       =   "MDIMain.frx":2C6AF
         MousePointer    =   99  'Custom
         Picture         =   "MDIMain.frx":2C9B9
         Top             =   300
         Width           =   1875
      End
      Begin VB.Image UsersMenu 
         Height          =   2070
         Left            =   2880
         MouseIcon       =   "MDIMain.frx":2F946
         MousePointer    =   99  'Custom
         Picture         =   "MDIMain.frx":2FC50
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "sub-menu  >>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   195
         Left            =   1020
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   750
         Left            =   210
         Picture         =   "MDIMain.frx":32791
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   2130
         TabIndex        =   5
         Top             =   1710
         Width           =   225
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S    Y    S     T    E    M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   510
         TabIndex        =   4
         Top             =   1980
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTORY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         TabIndex        =   3
         Top             =   1710
         Width           =   1635
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim userMenu_clicked As Boolean
Dim inventoryMenu_clicked As Boolean
Dim reportsMenu_clicked As Boolean
Dim buyMenu_clicked As Boolean

Dim submenu_adduser As Boolean
Dim submenu_edituser As Boolean
Dim submenu_deleteuser As Boolean

Dim submenu_addItemsAndInventory As Boolean
Dim submenu_editItemsAndInventory As Boolean
Dim submenu_deleteItemsAndInventory As Boolean
Dim submenu_supply As Boolean
Dim inventory_update As Boolean
Dim daily_sales As Boolean
Dim supply_reports As Boolean

Private Sub logout_Click()
    Call HideSubmenuForms
    Call HideContainers
    Call MDIForm_Load
End Sub

Private Sub MDIForm_Load()
    Call CentreFormOnMdi(formLogin, MainWindow)
    formLogin.Show
    'Call ShowContainers
    
    'For Main Menu
    userMenu_clicked = False
    inventoryMenu_clicked = False
    reportsMenu_clicked = False
    buyMenu_clicked = False
    
    'For Sub Menu
    submenu_adduser = False
    submenu_edituser = False
    submenu_deleteuser = False
    
    submenu_addItemsAndInventory = False
    submenu_editItemsAndInventory = False
    submenu_deleteItemsAndInventory = False
    submenu_supply = False
    inventory_update = False
    daily_sales = False
    supply_reports = False
End Sub


'#### SUB-MENUS ####
'START USERS SUB MENU
Private Sub AddUser_Click()
    submenu_edituser = False
    submenu_deleteuser = False
    
    'If submenu_adduser = True Then
    '    submenu_adduser = False
    'Else
        submenu_adduser = True
    'End If
    
    If (submenu_adduser) Then
        Call AddUserClicked
        AddUser.Picture = submenu_ImageList1.ListImages(2).Picture
            Call HideSubmenuForms
            formAddUser.Show
    End If

End Sub

Private Sub EditUser_Click()
    submenu_adduser = False
    submenu_deleteuser = False
    
    'If submenu_edituser = True Then
    '    submenu_edituser = False
    'Else
        submenu_edituser = True
    'End If
    
    If (submenu_edituser) Then
        Call EditUserClicked
        EditUser.Picture = submenu_ImageList2.ListImages(2).Picture
            Call HideSubmenuForms
            formEditUser.Show
    End If
End Sub

Private Sub DeleteUser_Click()
    submenu_adduser = False
    submenu_edituser = False
    
    'If submenu_deleteuser = True Then
    '    submenu_deleteuser = False
    'Else
        submenu_deleteuser = True
    'End If
    
    If (submenu_deleteuser) Then
        Call DeleteUserClicked
        DeleteUser.Picture = submenu_ImageList3.ListImages(2).Picture
            Call HideSubmenuForms
            formDeleteUser.Show
    End If
End Sub
'END USERS SUB MENU

'START INVENTORY SUB MENU
Private Sub AddItemsAndPackages_Click()
    submenu_supply = False
    inventory_update = False
    
    'If submenu_addItemsAndInventory = True Then
    '    submenu_addItemsAndInventory = False
    'Else
        submenu_addItemsAndInventory = True
    'End If
    
    If (submenu_addItemsAndInventory) Then
        Call AddItemsAndPackagesClicked
        AddItemsAndPackages.Picture = submenu_ImageList4.ListImages(2).Picture
            Call HideSubmenuForms
            formAddItemsAndPackages.Show
    End If
End Sub

Private Sub EditItemsAndPackages_Click()
    submenu_supply = False
    inventory_update = False
    
    'If submenu_editItemsAndInventory = True Then
    '    submenu_editItemsAndInventory = False
    'Else
        submenu_editItemsAndInventory = True
    'End If
    
    If (submenu_editItemsAndInventory) Then
        Call EditItemsAndPackagesClicked
        EditItemsAndPackages.Picture = submenu_ImageList9.ListImages(2).Picture
            Call HideSubmenuForms
            formEditItemsAndPackages.Show
    End If
End Sub

Private Sub DeleteItemsAndPackages_Click()
    submenu_supply = False
    inventory_update = False
    
    'If submenu_deleteItemsAndInventory = True Then
    '    submenu_deleteItemsAndInventory = False
    'Else
        submenu_deleteItemsAndInventory = True
    'End If
    
    If (submenu_deleteItemsAndInventory) Then
        Call DeleteItemsAndPackagesClicked
        DeleteItemsAndPackages.Picture = submenu_ImageList10.ListImages(2).Picture
            Call HideSubmenuForms
            formDeleteItemsAndPackages.Show
    End If
End Sub


Private Sub Supply_Click()
    submenu_addItemsAndInventory = False
    inventory_update = False
    
    'If submenu_supply = True Then
    '    submenu_supply = False
    'Else
        submenu_supply = True
    'End If
    
    If (submenu_supply) Then
        Call SupplyClicked
        Supply.Picture = submenu_ImageList5.ListImages(2).Picture
            Call HideSubmenuForms
            formSupply.Show
    End If
End Sub

Private Sub InventoryUpdate_Click()
    submenu_addItemsAndInventory = False
    submenu_supply = False
    
    'If inventory_update = True Then
    '    inventory_update = False
    'Else
        inventory_update = True
    'End If
    
    If (inventory_update) Then
        Call InventoryUpdateClicked
        InventoryUpdate.Picture = submenu_ImageList6.ListImages(2).Picture
            Call HideSubmenuForms
            formInventoryUpdate.Show
    End If
End Sub
'END INVENTORY SUB MENU

'START REPORTS SUB MENU
Private Sub DailySales_Click()
    supply_reports = False
    
    'If daily_sales = True Then
    '    daily_sales = False
    'Else
        daily_sales = True
    'End If
    
    If (daily_sales) Then
        Call DailySalesClicked
        DailySales.Picture = submenu_ImageList7.ListImages(2).Picture
    End If
    Unload formSupplierReport
    formSalesReport.Show
End Sub

Private Sub SupplyReports_Click()
    daily_sales = False
    
    'If supply_reports = True Then
    '    supply_reports = False
    'Else
        supply_reports = True
    'End If
    
    If (supply_reports) Then
        Call SupplyReportsClicked
        SupplyReports.Picture = submenu_ImageList8.ListImages(2).Picture
    End If
    Unload formSalesReport
    formSupplierReport.Show
End Sub
'END REPORTS SUB MENU
'#### END SUB-MENUS ####



'#### START MAIN MENU #####
Private Sub UsersMenu_Click()
    'FOR USERS MENU
    inventoryMenu_clicked = False
    reportsMenu_clicked = False
    buyMenu_clicked = False
            
    If userMenu_clicked = True Then
        userMenu_clicked = False
    Else
        userMenu_clicked = True
    End If
       
    If (userMenu_clicked) Then
        Call SetAllSubmenuToDefault
        Call UsersMenuClicked
        UsersMenu.Picture = ImageList1.ListImages(2).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            UsersContainer.Visible = True
    Else
        Call SetAllSubmenuToDefault
        UsersMenu.Picture = ImageList1.ListImages(1).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            UsersContainer.Visible = False
    End If
End Sub

Private Sub InventoryMenu_Click()
    'FOR INVENTORY MENU
    userMenu_clicked = False
    reportsMenu_clicked = False
    buyMenu_clicked = False
    
    If inventoryMenu_clicked = True Then
        inventoryMenu_clicked = False
    Else
        inventoryMenu_clicked = True
    End If
    
    If (inventoryMenu_clicked) Then
        Call SetAllSubmenuToDefault
        Call InventoryMenuClicked
        InventoryMenu.Picture = ImageList2.ListImages(2).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            InventorySubmenuContainer.Visible = True
    Else
        InventoryMenu.Picture = ImageList2.ListImages(1).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            InventorySubmenuContainer.Visible = False
    End If
End Sub

Private Sub ReportsMenu_Click()
    'FOR REPORTS MENU
    userMenu_clicked = False
    inventoryMenu_clicked = False
    buyMenu_clicked = False
    
    If reportsMenu_clicked = True Then
        reportsMenu_clicked = False
    Else
        reportsMenu_clicked = True
    End If
    
    If (reportsMenu_clicked) Then
        Call SetAllSubmenuToDefault
        Call ReportsMenuClicked
        ReportsMenu.Picture = ImageList3.ListImages(2).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            ReportsContainer.Visible = True
    Else
        ReportsMenu.Picture = ImageList3.ListImages(1).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            ReportsContainer.Visible = False
    End If
End Sub

Private Sub BuyMenu_Click()
    'FOR BUY MENU
    userMenu_clicked = False
    inventoryMenu_clicked = False
    reportsMenu_clicked = False
    
    If buyMenu_clicked = True Then
        buyMenu_clicked = False
    Else
        buyMenu_clicked = True
    End If
    
    If (buyMenu_clicked) Then
        Call SetAllSubmenuToDefault
        Call BuyMenuClicked
        BuyMenu.Picture = ImageList4.ListImages(2).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
            formBuy.Show
    Else
        BuyMenu.Picture = ImageList4.ListImages(1).Picture
            Call HideAllSubmenuContainer
            Call HideSubmenuForms
    End If
    
End Sub
'#### END MAIN MENU ####


