VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainfrmPrinc 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "principal form: adding and assigning teacher"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrinViewform.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrinViewform.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrinViewform.frx":05B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrinViewform.frx":070C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrinViewform.frx":0A34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   3120
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   5503
      ButtonWidth     =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   2
            TextSave        =   "3/24/2002"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   2
            TextSave        =   "6:05 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   540
      TabIndex        =   3
      Top             =   2580
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&View Report"
      Height          =   495
      Left            =   540
      TabIndex        =   2
      Top             =   1980
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search teacher"
      Height          =   555
      Left            =   540
      TabIndex        =   1
      Top             =   1320
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add teacher"
      Height          =   495
      Left            =   540
      TabIndex        =   0
      Top             =   720
      Width           =   2115
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddteacher 
         Caption         =   "Add Teacher"
      End
      Begin VB.Menu mnuSearchteacher 
         Caption         =   "Search Teacher"
      End
      Begin VB.Menu jghj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnurpt 
      Caption         =   "&Report"
      Begin VB.Menu mnuStudgrdrpt 
         Caption         =   "Show Student grade report"
      End
   End
   Begin VB.Menu mntols 
      Caption         =   "&Tools"
      Begin VB.Menu mnuchangepw 
         Caption         =   "Change Principal Password"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About FleXGrade..."
      End
   End
End
Attribute VB_Name = "frmMainfrmPrinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 frmAddTeacher.Show vbModal
End Sub

Private Sub Command2_Click()
 frmSearchinst.Show vbModal
End Sub

Private Sub Command5_Click()
MsgBox "Salamat sa pag-gamit sa among program....yeheyyyy!", vbOKOnly + vbInformation, "FleXGrade"
End
End Sub

Private Sub mnuAddteacher_Click()
 frmAddTeacher.Show vbModal
End Sub

