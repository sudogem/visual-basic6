VERSION 5.00
Begin VB.Form formError 
   BorderStyle     =   0  'None
   Caption         =   "Error"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   Picture         =   "frmError.frx":0000
   ScaleHeight     =   2175
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1920
      Picture         =   "frmError.frx":1A95
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label errorLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "formError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub errorMessage(message As String)
errorLabel.Caption = message
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


