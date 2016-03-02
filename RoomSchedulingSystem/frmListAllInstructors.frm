VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListAllInstructors 
   BackColor       =   &H00C0C0FF&
   Caption         =   "list of all instructors"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   420
      Left            =   240
      TabIndex        =   4
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   540
      Left            =   2640
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   525
      Left            =   7680
      TabIndex        =   2
      Top             =   4500
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "List of all Instructors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   1
      Top             =   180
      Width           =   3375
   End
End
Attribute VB_Name = "frmListAllInstructors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
 frmAddEditInstructor.Show
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdEdit_Click()
 frmAddEditInstructor.Show
End Sub

Private Sub Form_Load()
 
 'PopulateFlexGrid "select * from tblInstructor", MSFlexGrid1
 'PopulateFlexGrid1 "select * from tblInstructor", DataGrid1
 
End Sub
