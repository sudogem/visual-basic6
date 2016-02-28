VERSION 5.00
Begin VB.Form frmhelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   435
      Left            =   5460
      TabIndex        =   1
      Top             =   5460
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   7275
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
