VERSION 5.00
Begin VB.Form frmAddProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Product"
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
   Begin VB.Frame Frame1 
      Caption         =   "Product Information"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtUnitPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtTotalPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtDate 
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
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Product Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Total Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddProducts"
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
Set rs = db.OpenRecordset("Products")
txtDate.Text = Date
End Sub

Private Sub cmdAdd_Click()
Set rs = db.OpenRecordset("Products")
If txtID.Text = "" Or _
    txtUnitPrice.Text = "" Or _
    txtQuantity.Text = "" Or _
    txtTotalPrice.Text = "" Or _
    txtDescription.Text = "" Or _
    txtDate.Text = "" Then
    MsgBox "Please fill all the boxes."
Else
    rs.AddNew
    rs!ProductID = txtID.Text
    rs!UnitPrice = txtUnitPrice.Text
    rs!Quantity = txtQuantity.Text
    rs!TotalPrice = txtTotalPrice.Text
    rs!Description = txtDescription.Text
    rs!Date = txtDate.Text
    rs.Update
    Limpyo
    frmMain.ProductView
    Unload Me
End If

rs.Close
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Function Limpyo()
txtID.Text = ""
txtUnitPrice.Text = ""
txtQuantity.Text = ""
txtTotalPrice.Text = ""
txtDescription.Text = ""
End Function


Private Sub txtQuantity_Change()
Dim dig$, i, digi$, digits$
If txtQuantity.Text <> "" Then
    dig$ = Mid(txtQuantity.Text, Len(txtQuantity.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtQuantity.Text) - 1
            digi$ = Mid(txtQuantity.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtQuantity.Text = digits$
        txtQuantity.SelStart = Len(txtQuantity.Text)
    End If
End If
txtTotalPrice.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub txtUnitPrice_Change()
Dim dig$, i, digi$, digits$
If txtUnitPrice.Text <> "" Then
    dig$ = Mid(txtUnitPrice.Text, Len(txtUnitPrice.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtUnitPrice.Text) - 1
            digi$ = Mid(txtUnitPrice.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        txtUnitPrice.Text = digits$
        txtUnitPrice.SelStart = Len(txtUnitPrice.Text)
    End If
End If
txtTotalPrice.Text = Val(txtUnitPrice.Text) * Val(txtQuantity.Text)
End Sub
