VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9180
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "List of Subjects"
      Height          =   3075
      Left            =   6120
      TabIndex        =   8
      Top             =   1140
      Width           =   4395
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   4155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subjects"
      Height          =   3075
      Left            =   240
      TabIndex        =   1
      Top             =   1140
      Width           =   5715
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1260
         TabIndex        =   13
         Top             =   2580
         Width           =   1035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4380
         TabIndex        =   12
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   2100
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   1095
         Left            =   1260
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Units :"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Description :"
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Subjects :"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
   End
   Begin VB.Label Label1 
      Caption         =   "BSCS OF Curriculum 99 First  Year, FirstSem"
      Height          =   555
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
