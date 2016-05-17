VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Brought to us by:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   1680
      Picture         =   "frmAbout.frx":A06C
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   120
      Picture         =   "frmAbout.frx":C246
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   3480
      MouseIcon       =   "frmAbout.frx":FB14
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":10956
      Top             =   4200
      Width           =   1380
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me
End Sub
