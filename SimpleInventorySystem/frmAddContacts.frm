VERSION 5.00
Begin VB.Form frmAddContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Contact"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contact Information"
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtCel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cel.no.:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Phone no.:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contact ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset

Private Sub Form_Load()
Dim x
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\datas.mdb")
Set rs = db.OpenRecordset("Contacts")

If rs.RecordCount = 0 Then
txtID.Text = "1"
Else
rs.MoveLast
txtID.Text = rs!ContactID + 1
End If
End Sub

Private Sub cmdAdd_Click()
Set rs = db.OpenRecordset("Contacts")
If txtID.Text = "" Or _
    txtName.Text = "" Or _
    txtAddress.Text = "" Or _
    txtPhone.Text = "" Or _
    txtCel.Text = "" Or _
    txtEmail.Text = "" Then
    MsgBox "Please fill all the boxes."
Else
    rs.AddNew
    rs!ContactID = txtID.Text
    rs!Name = txtName.Text
    rs!Address = txtAddress.Text
    rs!PhoneNo = txtPhone.Text
    rs!Celno = txtCel.Text
    rs!Email = txtEmail.Text
    rs.Update
    Limpyo
    If rs.RecordCount = 0 Then
        txtID.Text = "1"
    Else
        rs.MoveLast
        txtID.Text = rs!ContactID + 1
    End If
    frmMain.ContactView
    Unload Me
End If
rs.Close
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Function Limpyo()
txtName.Text = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtCel.Text = ""
txtEmail.Text = ""
End Function

Private Sub txtCel_Change()
Dim dig$, i, digi$, digits$
If txtCel.Text <> "" Then
    dig$ = Mid(txtCel.Text, Len(txtCel.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtCel.Text) - 1
            digi$ = Mid(txtCel.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtCel.Text = digits$
        txtCel.SelStart = Len(txtCel.Text)
    End If
End If
End Sub

Private Sub txtPhone_Change()
Dim dig$, i, digi$, digits$
If txtPhone.Text <> "" Then
    dig$ = Mid(txtPhone.Text, Len(txtPhone.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtPhone.Text) - 1
            digi$ = Mid(txtPhone.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtPhone.Text = digits$
        txtPhone.SelStart = Len(txtPhone.Text)
    End If
End If
End Sub
