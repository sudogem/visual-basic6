VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainfrmStaff 
   Caption         =   "form for staff"
   ClientHeight    =   3570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4965
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   6297
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnustudinfo 
         Caption         =   "Student Information"
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnutols 
      Caption         =   "&Tools"
      Begin VB.Menu mnudel 
         Caption         =   "&Delete Student Record"
      End
      Begin VB.Menu mnuchangestaffpw 
         Caption         =   "&Change Staff Password..."
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMainfrmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnudel_Click()
 frmdelstudlog.Show vbModal
End Sub

Private Sub mnuexit_Click()
 End
End Sub

Private Sub mnustudinfo_Click()
 frmStaff.Show
End Sub
