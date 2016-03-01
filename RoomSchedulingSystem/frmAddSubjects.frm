VERSION 5.00
Begin VB.Form frmAddSubjects 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Height          =   480
      Left            =   7440
      TabIndex        =   8
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton cmdAddSubjSummer 
      Caption         =   "Add Subjects"
      Height          =   525
      Left            =   6000
      TabIndex        =   7
      Top             =   3600
      Width           =   2715
   End
   Begin VB.CommandButton cmdAddSubjSecSem 
      Caption         =   "Add Subjects"
      Height          =   525
      Left            =   3180
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdAddSubjFirstSem 
      Caption         =   "Add Subjects"
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Width           =   2715
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3150
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cboYearLevel 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   765
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Summer"
      Height          =   255
      Left            =   6060
      TabIndex        =   11
      Top             =   1140
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Second Semester"
      Height          =   195
      Left            =   3180
      TabIndex        =   10
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "First Semester"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Year level :"
      Height          =   255
      Left            =   315
      TabIndex        =   0
      Top             =   825
      Width           =   975
   End
End
Attribute VB_Name = "frmAddSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddSubjFirstSem_Click()
 frmAddSubjects.Show
End Sub
Private Sub cmdAddSubjSecSem_Click()
  frmAddSubjects.Show
End Sub
Private Sub cmdAddSubjSummer_Click()
  frmAddSubjects.Show
End Sub
Private Sub Command4_Click()
 Unload Me
End Sub
