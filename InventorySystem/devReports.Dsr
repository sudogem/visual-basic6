VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} devReports 
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13860
   _ExtentX        =   24448
   _ExtentY        =   15769
   FolderFlags     =   1
   TypeLibGuid     =   "{EB615F02-E797-476E-A9F9-64A0BFD5C59B}"
   TypeInfoGuid    =   "{0224221D-908A-44F8-AC34-E751573718D1}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "devReports"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\shared\inventory\database\InventoryDB.mdb;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   4
   BeginProperty Recordset1 
      CommandName     =   "cmdSalesReport_Mas"
      CommDispId      =   1002
      RsDispId        =   1019
      CommandText     =   "SELECT out_mas.* FROM out_mas"
      ActiveConnectionName=   "devReports"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "transid"
         Caption         =   "transid"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "date"
         Caption         =   "date"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "buyer"
         Caption         =   "buyer"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "contact"
         Caption         =   "contact"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "total"
         Caption         =   "total"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdSalesReport_Det"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT out_det.* FROM out_det"
      ActiveConnectionName=   "devReports"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdSalesReport_mas"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "transid"
         Caption         =   "transid"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "productid"
         Caption         =   "productid"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qty"
         Caption         =   "qty"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "sellprice"
         Caption         =   "sellprice"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "total"
         Caption         =   "total"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "transid"
         ChildField      =   "transid"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdSupplierReport_Mas"
      CommDispId      =   1020
      RsDispId        =   1023
      CommandText     =   "SELECT * FROM in_mas"
      ActiveConnectionName=   "devReports"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "transid"
         Caption         =   "transid"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "date"
         Caption         =   "date"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "supplier"
         Caption         =   "supplier"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "contact"
         Caption         =   "contact"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "cmdSupplierReport_Det"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "SELECT * FROM in_det"
      ActiveConnectionName=   "devReports"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdSupplierReport_Mas"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "det_id"
         Caption         =   "det_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "transid"
         Caption         =   "transid"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "itemid"
         Caption         =   "itemid"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qty"
         Caption         =   "qty"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "cost"
         Caption         =   "cost"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "transid"
         ChildField      =   "transid"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "devReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rsCommand1_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub rscmdSalesReport_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

