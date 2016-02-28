VERSION 5.00
Begin VB.Form frmdelstudlog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Student Record"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmdelstud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter Staff Password to confirmed deletion."
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3855
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1260
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Username :"
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Password :"
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   780
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmdelstudlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If StrComp(tempUsername, txtUsername.Text, vbTextCompare) = 0 Then
  MsgBox "ok...username"
   If StrComp(tempPassword, txtPassword.Text, vbTextCompare) = 0 Then
       MsgBox "ok password"
      Else
       MsgBox "bad pw"
   End If
Else
     MsgBox "bad user"
End If
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub
