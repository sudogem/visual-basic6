Attribute VB_Name = "Module3"
Option Explicit

Private Function formatStringforSQL(str As String) As String
    formatStringforSQL = Replace(str, "'", "''")
End Function

Public Sub openConnection()
On Error GoTo errhandler
    Dim sConnString As String
    sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath
    If Not conn.State = adStateOpen Then
        conn.Open sConnString
    End If
    Exit Sub
errhandler:
    Set conn = Nothing
End Sub

Public Sub closeConnection()
On Error Resume Next
    conn.Close
    Set conn = Nothing
End Sub

Public Function getUserInfo(conn As ADODB.Connection, userid As Long) As ADODB.Recordset
    Dim strquery As String
    strquery = "SELECT * FROM users WHERE userid=" & userid
    Set getUserInfo = conn.Execute(strquery)
End Function

Public Function AddUser(conn As ADODB.Connection, usr As String, pass As String, fn As String, ln As String, add As String, cnum As String, utype As String) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    
    'set status flag to true or -1
    strquery = "INSERT INTO users([username], [password], [firstname], [lastname], [address], [contact_number], [utype], [status_flag]) VALUES ('" & usr & "','" & pass & "','" & fn & "','" & ln & "','" & add & "','" & cnum & "', '" & utype & "',TRUE)"
    'MsgBox strquery
    conn.Execute strquery
    AddUser = True
    Exit Function
errhandler:
    'MsgBox Err.Description
    AddUser = False
End Function

Public Function EditUser(conn As ADODB.Connection, userid As Long, usr As String, pass As String, fn As String, ln As String, add As String, cnum As String, utype As String) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    
    'set status flag to true or -1
    strquery = "UPDATE users SET [username]='" & usr & "', [password]='" & pass & "', [firstname]='" & fn & "', [lastname]='" & ln & "', [address]='" & add & "', [contact_number]='" & cnum & "', [utype]='" & utype & "' WHERE userid=" & userid
    'MsgBox strquery
    conn.Execute strquery
    EditUser = True
    Exit Function
errhandler:
    'MsgBox Err.Description
    EditUser = False
End Function

Public Function DeleteUser(conn As ADODB.Connection, usrID As Long) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    
    'set status flag to true or -1
    strquery = "DELETE FROM users WHERE userid=" & usrID
    'MsgBox strquery
    conn.Execute strquery
    DeleteUser = True
    Exit Function
errhandler:
    'MsgBox Err.Description
    DeleteUser = False
End Function


Public Function addItem(conn As ADODB.Connection, itemid As String, itemType As String, desc As String, quantity As Long, price As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    strquery = "INSERT INTO item([itemid],[itype],[desc],[qty],[price]) VALUES('" & formatStringforSQL(itemid) & "','" & formatStringforSQL(itemType) & "','" & formatStringforSQL(desc) & "'," & quantity & "," & price & ")"
    'MsgBox strquery
    conn.Execute strquery
    addItem = True
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    addItem = False
End Function

Public Function editItem(conn As ADODB.Connection, itemIDOld As String, itemid As String, itemType As String, desc As String, quantity As Long, price As Double)
On Error GoTo errhandler
    Dim strquery As String
    strquery = "UPDATE item SET [itemid]='" & itemid & "',[itype]='" & formatStringforSQL(itemType) & "',[desc]='" & formatStringforSQL(desc) & "',[qty]=" & quantity & ",[price]=" & price & " WHERE itemid='" & itemIDOld & "'"
    'MsgBox strquery
    conn.Execute strquery
    editItem = True
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    editItem = False
End Function

Public Function deleteItem(conn As ADODB.Connection, itemid As String) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    conn.BeginTrans
    strquery = "DELETE FROM package WHERE itemid='" & itemid & "'"
    'MsgBox strquery
    conn.Execute strquery
    strquery = "DELETE FROM item WHERE itemid='" & itemid & "'"
    'MsgBox strquery
    conn.Execute strquery
    deleteItem = True
    conn.CommitTrans
    Exit Function
errhandler:
    conn.RollbackTrans
    'MsgBox Err.Number & ": " & Err.Description
    deleteItem = False
End Function

Public Function addPackage(conn As ADODB.Connection, packageid As String, itemid As String, pname As String, qty As Long, price As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    strquery = "INSERT INTO package(packageid, itemid, name, qty, price) VALUES('" & formatStringforSQL(packageid) & "','" & formatStringforSQL(itemid) & "','" & formatStringforSQL(pname) & "'," & qty & "," & price & ")"
    'MsgBox strquery
    conn.Execute strquery
    addPackage = True
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    addPackage = False
End Function

Public Function editPackage(conn As ADODB.Connection, oldpackageID As String, packageid As String, itemid As String, pname As String, qty As Long, price As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
          
    strquery = "INSERT INTO package(packageid, itemid, name, qty, price) VALUES('" & formatStringforSQL(packageid) & "','" & formatStringforSQL(itemid) & "','" & formatStringforSQL(pname) & "'," & qty & "," & price & ")"
    'MsgBox strquery
    conn.Execute strquery
    editPackage = True
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    editPackage = False
End Function

Public Function deletePackage(conn As ADODB.Connection, packageid As String) As Boolean
On Error GoTo errhandler
    Dim strquery As String
          
    strquery = "DELETE FROM package WHERE packageid='" & packageid & "'"
    'MsgBox strquery
    conn.Execute strquery
    deletePackage = True
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    deletePackage = False
End Function

Public Function isPackageUsed(conn As ADODB.Connection, packageid As String) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    Dim rs As New ADODB.Recordset
    strquery = "SELECT out_det.productid From out_det WHERE (((out_det.productid)='" & packageid & "'))"
    Set rs = conn.Execute(strquery)
    If rs.BOF = True Or rs.EOF = True Then
        isPackageUsed = False
    Else
        isPackageUsed = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
errhandler:
    isPackageUsed = False
    Set rs = Nothing
End Function

Public Function addOutTransactionMaster(conn As ADODB.Connection, theDate As Date, buyer As String, add As String, contact As String, total As Double, userid As Long) As Double
On Error GoTo errhandler
    Dim strquery As String
    Dim id As Double
    id = CDbl(Now())
    strquery = "INSERT INTO out_mas([transid], [date], [buyer], [address], [total], [contact], [userid]) VALUES(" & id & ",#" & theDate & "#,'" & formatStringforSQL(buyer) & "','" & formatStringforSQL(add) & "'," & total & ",'" & formatStringforSQL(contact) & "'," & userid & ")"
    'MsgBox strquery
    conn.Execute strquery
    addOutTransactionMaster = id
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    addOutTransactionMaster = -1
End Function

Public Function getQuantityOfPackage(conn As ADODB.Connection, packageid As String) As Double
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    strquery = "SELECT qty FROM package WHERE packageid='" & packageid & "'"
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    If rs.EOF = True Or rs.BOF = True Then
        getQuantityOfPackage = -1
    Else
        getQuantityOfPackage = rs!qty
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    Set rs = Nothing
    getQuantityOfPackage = -1
End Function

Public Function getItemIDOfPackage(conn As ADODB.Connection, packageid As String) As String
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    strquery = "SELECT itemid FROM package WHERE packageid='" & packageid & "'"
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    If rs.EOF = True Or rs.BOF = True Then
        getItemIDOfPackage = ""
    Else
        getItemIDOfPackage = rs!itemid & ""
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    Set rs = Nothing
    getItemIDOfPackage = ""
End Function

Public Function getQuantityOfItems(conn As ADODB.Connection, itemid As String) As Double
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    strquery = "SELECT qty FROM item WHERE itemid='" & itemid & "'"
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    If rs.EOF = True Or rs.BOF = True Then
        getQuantityOfItems = -1
    Else
        getQuantityOfItems = rs!qty
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    Set rs = Nothing
    getQuantityOfItems = -1
End Function


Public Function checkItemLevelIfPwedeOutAni(conn As ADODB.Connection, packageid As String, qty As Double) As Boolean
On Error GoTo errhandler
    'check entries
    Dim itemqty As Double
    Dim packageitemqty As Double
    Dim itemid As String
        
    'check if pwede ni
    'get item quantity per package
    packageitemqty = getQuantityOfPackage(conn, packageid)
    'MsgBox packageitemqty
    'get itemid of package
    itemid = getItemIDOfPackage(conn, packageid)
    'get total item qty
    itemqty = getQuantityOfItems(conn, itemid)
    'MsgBox "items: " & itemqty
    
    'check if pwede pa
    If (qty * packageitemqty) <= itemqty Then
        'very good
        checkItemLevelIfPwedeOutAni = True
    Else
        'dili pwede!!!
        checkItemLevelIfPwedeOutAni = False
    End If
    Exit Function
errhandler:
    checkItemLevelIfPwedeOutAni = False
End Function


'make sure to use transactions!
Public Function addOutTransactionDetail(conn As ADODB.Connection, transID As Double, packageid As String, qty As Double, sellprice As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    Dim result As Boolean
    'check sa
    result = checkItemLevelIfPwedeOutAni(conn, packageid, qty)
    'if ok padayon
    If result = False Then
        addOutTransactionDetail = False
        Exit Function
    End If
    
    'add to detail
    
    conn.BeginTrans
    
    strquery = "INSERT INTO out_det(transid,productid, qty, sellprice,total) VALUES(" & transID & ",'" & packageid & "'," & qty & "," & sellprice & "," & (qty * sellprice) & ")"
    'MsgBox strquery
    conn.Execute strquery
    
    'subract from item qty
    Dim itemqty As Double
    Dim itemid As String
    Dim packageitemqty As Double
    itemid = getItemIDOfPackage(conn, packageid)
    itemqty = getQuantityOfItems(conn, itemid)
    packageitemqty = getQuantityOfPackage(conn, packageid)
    
    strquery = "UPDATE item SET qty=" & (itemqty - (qty * packageitemqty)) & " WHERE itemid='" & itemid & "'"
    'MsgBox strquery
    conn.Execute strquery
    
    conn.CommitTrans
    addOutTransactionDetail = True
    Exit Function
errhandler:
    conn.RollbackTrans
    'MsgBox Err.Number & " " & Err.Description
    addOutTransactionDetail = False
End Function

Public Function addInTransactionMaster(conn As ADODB.Connection, theDate As Date, supplier As String, contact As String, userid As Long) As Double
On Error GoTo errhandler
    Dim strquery As String
    Dim id As Double
    Dim d As Date
    d = Now
    id = CDbl(d)
    strquery = "INSERT INTO in_mas([transid], [date], [supplier], [contact], [userid]) VALUES(" & id & ",#" & theDate & "#,'" & formatStringforSQL(supplier) & "','" & formatStringforSQL(contact) & "'," & userid & ")"
    'MsgBox strquery
    conn.Execute strquery
    addInTransactionMaster = id
    Exit Function
errhandler:
    'MsgBox Err.Number & " " & Err.Description
    addInTransactionMaster = -1
End Function

'make sure this belongs in a transaction
Public Function addInTransactionDetail(conn As ADODB.Connection, transID As Double, itemid As String, qty As Long, cost As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    strquery = "INSERT INTO in_det(transid,itemid, qty, cost) VALUES(" & transID & ",'" & itemid & "'," & qty & "," & cost & ")"
    'MsgBox strquery
    conn.Execute strquery
    
    'add to item qty
    Dim itemqty As Double
    itemqty = getQuantityOfItems(conn, itemid)
    
    strquery = "UPDATE item SET qty=" & (itemqty + qty) & " WHERE itemid='" & itemid & "'"
    'MsgBox strquery
    conn.Execute strquery
    
    
    addInTransactionDetail = True
    Exit Function
errhandler:
    MsgBox Err.Number & " " & Err.Description
    addInTransactionDetail = False
End Function

Public Function getItemCodes(conn As ADODB.Connection) As ADODB.Recordset
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT itemid FROM item", conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set getItemCodes = rs
    Exit Function
errhandler:
    'MsgBox Err.Description
    Set getItemCodes = Nothing
End Function

Public Function getPackageCodes(conn As ADODB.Connection) As ADODB.Recordset
On Error GoTo errhandler
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT packageid FROM package", conn, adOpenStatic, adLockOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set getPackageCodes = rs
    Exit Function
errhandler:
    'MsgBox Err.Description
    Set getPackageCodes = Nothing
End Function

Public Function getReorderlevel(conn As ADODB.Connection) As Double
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    getReorderlevel = 0
    strquery = "SELECT * FROM reorder"
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    If rs.EOF = True And rs.BOF = True Then
        'add a reorderlevel
        strquery = "INSERT INTO reorder(reorderlevel) VALUES(2000)"
        conn.Execute strquery
        getReorderlevel = 2000
        Exit Function
    Else
        rs.MoveFirst
        getReorderlevel = rs!reorderlevel
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function setReorderlevel(conn As ADODB.Connection, level As Double) As Boolean
On Error GoTo errhandler
    Dim strquery As String
    Dim rlevel As Double
    rlevel = getReorderlevel(conn)
    If rlevel = level Then
        setReorderlevel = True
        Exit Function
    End If
    
    strquery = "UPDATE reorder SET reorderlevel=" & level
    conn.Execute strquery
    
    setReorderlevel = True
    Exit Function
errhandler:
    'MsgBox Err.Description
    setReorderlevel = False
End Function

Public Function checkLevels(conn As ADODB.Connection) As Boolean
    checkLevels = False
    Dim rs As New ADODB.Recordset
    Dim strquery As String
    Dim level As Double
    Dim mMessage As String
    
    level = getReorderlevel(conn)
    
    strquery = "SELECT * FROM item"
    rs.Open strquery, conn, adOpenStatic, adLockPessimistic, adCmdText
    If rs.EOF = True And rs.BOF = True Then
        Exit Function
    Else
        rs.MoveFirst
        While Not rs.EOF
            If rs!qty <= level Then
                'mMessage rs!itemid & " is below minimum level!"
                'MsgBox rs!itemid & " is below minimum level!", vbOKOnly, "Reorderlevel reached!"
                'formError.errorLabel.Caption = mMessage
                'formError.Show vbModal
                checkLevels = True
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
End Function
