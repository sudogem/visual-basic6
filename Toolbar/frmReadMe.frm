VERSION 5.00
Begin VB.Form frmReadMe 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Read Me"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data file locations must be set to: C:\KDJDev\KDJCode\Examples\Toolbar"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pre-requisites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reference Microsoft ActiveX Data Object 2.0 Library (or a later version)"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frmReadMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

