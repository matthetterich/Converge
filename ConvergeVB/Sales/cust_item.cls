VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cust_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"customer"
Option Explicit
Private m_objTransData As Object
Private m_objError As Object
Private m_objCustomer As Object
Private m_objItem As Object
Private m_varCustItemNbr As Variant
Private m_varCustId As Variant
Private m_varCustItemDesc As Variant
Private m_varItemId As Variant
Private m_varByPassGetItemFl As Variant
Private m_varByPassGetCustFl As Variant
Private m_varCustItemPrice As Variant
Private m_varAllowQtyDiscFl As Variant
Private m_varInitialPurchaseDt As Variant
Private m_objRecordset As adodb.Recordset
'

Public Property Set objRecordset(ByVal vData As adodb.Recordset)
    Set m_objRecordset = vData
End Property
Public Property Get objRecordset() As adodb.Recordset
    Set objRecordset = m_objRecordset
End Property
Public Property Let allowQtyDiscFl(ByVal vData As Variant)
    m_varAllowQtyDiscFl = vData
End Property
Public Property Get allowQtyDiscFl() As Variant
    allowQtyDiscFl = m_varAllowQtyDiscFl
End Property
Public Property Let custItemPrice(ByVal vData As Variant)
    m_varCustItemPrice = vData
End Property
Public Property Get custItemPrice() As Variant
    custItemPrice = m_varCustItemPrice
End Property
Public Property Let byPassGetCustFl(ByVal vData As Variant)
    m_varByPassGetCustFl = vData
End Property
Public Property Get byPassGetCustFl() As Variant
    byPassGetCustFl = m_varByPassGetCustFl
End Property
Public Property Let byPassGetItemFl(ByVal vData As Variant)
    m_varByPassGetItemFl = vData
End Property
Public Property Get byPassGetItemFl() As Variant
    byPassGetItemFl = m_varByPassGetItemFl
End Property
Public Property Let itemId(ByVal vData As Variant)
    m_varItemId = vData
End Property
Public Property Get itemId() As Variant
    itemId = m_varItemId
End Property
Public Property Let custItemDesc(ByVal vData As Variant)
    m_varCustItemDesc = vData
End Property
Public Property Get custItemDesc() As Variant
    custItemDesc = m_varCustItemDesc
End Property
Public Property Let custId(ByVal vData As Variant)
    m_varCustId = vData
End Property
Public Property Get custId() As Variant
    custId = m_varCustId
End Property
Public Property Let custItemNbr(ByVal vData As Variant)
    m_varCustItemNbr = vData
End Property
Public Property Get custItemNbr() As Variant
    custItemNbr = m_varCustItemNbr
End Property
Public Property Let initialPurchaseDt(ByVal vData As Variant)
    m_varInitialPurchaseDt = vData
End Property
Public Property Get initialPurchaseDt() As Variant
    initialPurchaseDt = m_varInitialPurchaseDt
End Property
Public Property Set objItem(ByVal vData As Object)
    Set m_objItem = vData
End Property
Public Property Get objItem() As Object
    Set objItem = m_objItem
End Property
Public Property Set objCustomer(ByVal vData As Object)
    Set m_objCustomer = vData
End Property
Public Property Get objCustomer() As Object
    Set objCustomer = m_objCustomer
End Property
Public Property Set objError(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get objError() As Object
    Set objError = m_objError
End Property
Public Property Set objTransData(ByVal vData As Object)
    Set m_objTransData = vData
End Property
Public Property Get objTransData() As Object
    Set objTransData = m_objTransData
End Property

'obsolete Properties
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set m_objRecordset = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = m_objRecordset
End Property

Public Property Let p_allow_qty_disc_fl(ByVal vData As Variant)
    m_varAllowQtyDiscFl = vData
End Property
Public Property Get p_allow_qty_disc_fl() As Variant
    p_allow_qty_disc_fl = m_varAllowQtyDiscFl
End Property
Public Property Let p_cust_item_price(ByVal vData As Variant)
    m_varCustItemPrice = vData
End Property
Public Property Get p_cust_item_price() As Variant
        p_cust_item_price = m_varCustItemPrice
End Property
Public Property Let p_bypass_getcust_fl(ByVal vData As Variant)
    m_varByPassGetCustFl = vData
End Property
Public Property Get p_bypass_getcust_fl() As Variant
    p_bypass_getcust_fl = m_varByPassGetCustFl
End Property
Public Property Let p_bypass_getitem_fl(ByVal vData As Variant)
    m_varByPassGetItemFl = vData
End Property
Public Property Get p_bypass_getitem_fl() As Variant
    p_bypass_getitem_fl = m_varByPassGetItemFl
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    m_varItemId = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = m_varItemId
End Property
Public Property Let p_cust_item_desc(ByVal vData As Variant)
    m_varCustItemDesc = vData
End Property
Public Property Get p_cust_item_desc() As Variant
    p_cust_item_desc = m_varCustItemDesc
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    m_varCustId = vData
End Property
Public Property Get p_cust_id() As Variant
    p_cust_id = m_varCustId
End Property
Public Property Let p_cust_item_nbr(ByVal vData As Variant)
    m_varCustItemNbr = vData
End Property
Public Property Get p_cust_item_nbr() As Variant
    p_cust_item_nbr = m_varCustItemNbr
End Property
Public Property Set o_item(ByVal vData As Object)
    Set m_objItem = vData
End Property
Public Property Get o_item() As Object
    Set o_item = m_objItem
End Property
Public Property Set o_customer(ByVal vData As Object)
    Set m_objCustomer = vData
End Property
Public Property Get o_customer() As Object
    Set o_customer = m_objCustomer
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTransData = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTransData
End Property
Public Property Let allow_qty_disc_fl(ByVal vData As Variant)
    m_varAllowQtyDiscFl = vData
End Property
Public Property Get allow_qty_disc_fl() As Variant
    allow_qty_disc_fl = m_varAllowQtyDiscFl
End Property
Public Property Let cust_item_price(ByVal vData As Variant)
    m_varCustItemPrice = vData
End Property
Public Property Get cust_item_price() As Variant
        cust_item_price = m_varCustItemPrice
End Property
Public Property Let bypass_getcust_fl(ByVal vData As Variant)
    m_varByPassGetCustFl = vData
End Property
Public Property Get bypass_getcust_fl() As Variant
    bypass_getcust_fl = m_varByPassGetCustFl
End Property
Public Property Let bypass_getitem_fl(ByVal vData As Variant)
    m_varByPassGetItemFl = vData
End Property
Public Property Get bypass_getitem_fl() As Variant
    bypass_getitem_fl = m_varByPassGetItemFl
End Property
Public Property Let item_id(ByVal vData As Variant)
    m_varItemId = vData
End Property
Public Property Get item_id() As Variant
    item_id = m_varItemId
End Property
Public Property Let cust_item_desc(ByVal vData As Variant)
    m_varCustItemDesc = vData
End Property
Public Property Get cust_item_desc() As Variant
    cust_item_desc = m_varCustItemDesc
End Property
Public Property Let cust_id(ByVal vData As Variant)
    m_varCustId = vData
End Property
Public Property Get cust_id() As Variant
    cust_id = m_varCustId
End Property
Public Property Let cust_item_nbr(ByVal vData As Variant)
    m_varCustItemNbr = vData
End Property
Public Property Get cust_item_nbr() As Variant
    cust_item_nbr = m_varCustItemNbr
End Property

Public Property Set Item(ByVal vData As Object)
    Set m_objItem = vData
End Property
Public Property Get Item() As Object
    Set Item = m_objItem
End Property
Public Property Set customer(ByVal vData As Object)
    Set m_objCustomer = vData
End Property
Public Property Get customer() As Object
    Set customer = m_objCustomer
End Property
Public Property Set error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get error() As Object
    Set error = m_objError
End Property

Public Property Set trans_data(ByVal vData As Object)
    Set m_objTransData = vData
End Property
Public Property Get trans_data() As Object
    Set trans_data = m_objTransData
End Property

Public Sub GetCustItem()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_item.cls.GetCust_item"
    Dim strSql As String
    Dim dblUpdateQty As Double
    Dim objRecordset As adodb.Recordset

    If IsNull(Me.itemId) = True Then
        strSql = "SELECT * " & _
                 "FROM cust_item " & _
                 "WHERE " & _
                 "    " & fWhereVariable("cust_id", Me.custId, "N", "=") & _
                 " AND " & fWhereVariable("cust_item_nbr", Me.custItemNbr, "S", "=")
    Else
        strSql = "SELECT * " & _
                 "FROM cust_item " & _
                 "WHERE  " & _
                 "    " & fWhereVariable("cust_id", Me.custId, "N", "=") & _
                 " and " & fWhereVariable("item_id", Me.itemId, "N", "=")
    End If
    
    Call Me.objTransData.OpenRecordset(objRecordset, strSql, Me.objError)
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        Me.custId = Null
        Me.itemId = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        Me.custId = objRecordset("cust_id")
        Me.itemId = objRecordset("item_id")
        Me.custItemNbr = objRecordset("cust_item_nbr")
        Me.custItemDesc = objRecordset("cust_item_desc")
        Me.custItemPrice = objRecordset("cust_item_price")
        Me.allowQtyDiscFl = objRecordset("allow_qty_disc_fl")
        Me.initialPurchaseDt = objRecordset("Initial_purchase_dt")
        
        objRecordset.Close
        Set objRecordset = Nothing

        If Me.byPassGetItemFl = "Y" Then
        Else
            Set Me.objItem = CreateObject("converge_item.item")
            Set Me.objItem.o_error = Me.objError
            Set Me.objItem.o_trans_data = Me.objTransData
            Me.objItem.p_item_id = Me.itemId
            Call Me.objItem.getitem
        End If
    
        If Me.byPassGetCustFl = "Y" Then
        Else
            Set Me.objCustomer = CreateObject("converge_sales.customer")
            Set Me.objCustomer.o_error = Me.objError
            Set objCustomer.o_trans_data = Me.objTransData
            Me.objCustomer.p_cust_id = Me.custId
            Call Me.objCustomer.GetCustomer
        End If
    End If
     
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "0200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

    
Public Sub Add()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_item.cls.Add"
    
    Call validate("add")
    Call IssueInsert
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "0400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub Change()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_item.cls.Change"
    
    Call validate("change")
    Call IssueUpdate
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_item.cls.DeleteComment"
    
    Call validate("delete")
    Call IssueDelete
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "0800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub validate(v_strFunctionCd)
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_item.cls.Validate"
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset

'    If v_strFunctionCd = "delete" Then
'        With Me.objError
'            .type_cd = "F"
'            .err_cd = "1000"
'            .routine_nm = strRoutine_nm
'            .message_id = 1002
'        End With
'    End If

    If v_strFunctionCd = "add" And fValuePresent(Me.itemId) = False Then
        If fValuePresent(Me.objItem.p_item_nbr) = False Then
            Me.objError.type_cd = "F"
            Me.objError.err_cd = "1000"
            Me.objError.routine_nm = strRoutine_nm
            Me.objError.message_id = 1050
            Exit Sub
        End If
            
        Me.objItem.p_item_id = Null
        Set Me.objItem.o_error = Me.objError
        Set Me.objItem.o_trans_data = Me.objTransData
        Call Me.objItem.getitem
        
        If Me.objError.type_cd = "E" Or Me.objError.type_cd = "F" Then
            Exit Sub
        Else
            If IsNull(objItem.p_item_id) = True Then
                Me.objError.type_cd = "F"
                Me.objError.err_cd = "1200"
                Me.objError.routine_nm = strRoutine_nm
                Me.objError.message_id = 1045
                Exit Sub
            Else
                Me.itemId = Me.objItem.p_item_id
            End If
        End If
    End If
    
    If v_strFunctionCd = "add" And fValuePresent(Me.custId) = False Then
        If fValuePresent(objCustomer.p_cust_nbr) = False Then
            Me.objError.type_cd = "F"
            Me.objError.err_cd = "1400"
            Me.objError.routine_nm = strRoutine_nm
            Me.objError.message_id = 1075
            Exit Sub
        End If
        
        Me.objCustomer.p_cust_id = Null
        Set Me.objCustomer.o_error = Me.objError
        Set Me.objCustomer.o_trans_data = Me.objTransData
        Call Me.objCustomer.GetCustomer
        
        If Me.objError.type_cd = "E" Or Me.objError.type_cd = "F" Then
            Exit Sub
        Else
            If IsNull(Me.objCustomer.p_cust_id) = True Then
                Me.objError.type_cd = "F"
                Me.objError.err_cd = "1600"
                Me.objError.routine_nm = strRoutine_nm
                Me.objError.message_id = 1080
                Exit Sub
            Else
                Me.custId = Me.objCustomer.p_cust_id
            End If
        End If
    End If
    
    If v_strFunctionCd = "add" Or v_strFunctionCd = "change" Then
        If Len(Me.custItemNbr) > 0 Then
        Else
            Me.objError.type_cd = "F"
            Me.objError.err_cd = "1800"
            Me.objError.routine_nm = strRoutine_nm
            Me.objError.message_id = 1085
            Exit Sub
        End If
    
        If IsEmpty(Me.initialPurchaseDt) = False And IsNull(Me.initialPurchaseDt) = False Then
            If IsDate(Me.initialPurchaseDt) = False Then
                Me.objError.type_cd = "F"
                Me.objError.err_cd = "1087"
                Me.objError.routine_nm = strRoutine_nm
                Me.objError.message_id = 0
                Exit Sub
            End If
        End If
    End If
    
    If v_strFunctionCd = "add" Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM cust_item "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_id", Me.custId, "N", "=")
        strSql = strSql & "and" & fWhereVariable("item_id", Me.itemId, "N", "=")
        strSql = strSql & "and deleted_fl <> 'Y' "
        
        Call Me.objTransData.OpenRecordset(objRecordset, strSql, Me.objError)
        
        If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
            Me.objRecordset.Close
            Set objRecordset = Nothing
            Exit Sub
        End If
    
        If objRecordset.EOF = True Then
        Else
            Me.objError.type_cd = "F"
            Me.objError.err_cd = "2000"
            Me.objError.routine_nm = strRoutine_nm
            Me.objError.message_id = 1060
            Exit Sub
        End If
         
        objRecordset.Close
        Set objRecordset = Nothing
    End If
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "2200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String

    strRoutine_nm = "cust_item.cls.IssueInsert"

                    'delete the cust/item first.  If the user deleted the cust/item
                    'and is not inserting a new one, the old one was only marked as
                    'deleted.  Now, before we can insert the same cust/item, we need
                    'to remove the previous one that may exist.
    strSql = "delete from cust_item " & _
             "where " & _
             "    " & fWhereVariable("cust_id", Me.custId, "N", "=") & _
             "and " & fWhereVariable("item_id", Me.itemId, "N", "=")
    
    Call Me.objTransData.ExecuteDatabase(strSql, dblUpdateQty, objError)
    
    strSql = "insert into cust_item ( "
    strSql = strSql & "  cust_id "
    strSql = strSql & ", item_id "
    strSql = strSql & ", cust_item_nbr "
    strSql = strSql & ", cust_item_desc "
    strSql = strSql & ", cust_item_price "
    strSql = strSql & ", allow_qty_disc_fl "
    strSql = strSql & ", initial_purchase_dt "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(Me.custId, "N")
    strSql = strSql & "," & fInsertVariable(Me.itemId, "N")
    strSql = strSql & "," & fInsertVariable(Me.custItemNbr, "S")
    strSql = strSql & "," & fInsertVariable(Me.custItemDesc, "S")
    strSql = strSql & "," & fInsertVariable(Me.custItemPrice, "N")
    strSql = strSql & "," & fInsertVariable(Me.allowQtyDiscFl, "S")
    strSql = strSql & "," & fInsertVariable(Me.initialPurchaseDt, "D")
    strSql = strSql & "," & fInsertVariable(Me.objTransData.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Me.objTransData.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    Call Me.objTransData.ExecuteDatabase(strSql, dblUpdateQty, objError)
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "2400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strRoutine_nm = "cust_item.cls.IssueUpdate"

   
    strSql = "update cust_item set "
    strSql = strSql & " " & fUpdateVariable("cust_item_nbr", Me.custItemNbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_item_desc", Me.custItemDesc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_item_price", Me.custItemPrice, "N", "=")
    strSql = strSql & "," & fUpdateVariable("allow_qty_disc_fl", Me.allowQtyDiscFl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("initial_purchase_dt", Me.initialPurchaseDt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Me.objTransData.assoc_id, "S", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("cust_id", Me.custId, "N", "=")
    strSql = strSql & "and" & fWhereVariable("item_id", Me.itemId, "N", "=")
    
    Call Me.objTransData.ExecuteDatabase(strSql, dblUpdateQty, Me.objError)

    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "2600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strRoutine_nm = "cust_item.cls.IssueDelete"

'    strSql = "delete from cust_item " & _
'             "where " & _
'             "    " & fWhereVariable("cust_id", Me.custId, "N", "=") & _
'             "and " & fWhereVariable("item_id", Me.itemId, "N", "=")
    strSql = "update cust_item set deleted_fl = 'Y' "
    strSql = strSql & " where " & fWhereVariable("cust_id", Me.custId, "N", "=")
    strSql = strSql & " and" & fWhereVariable("item_id", Me.itemId, "N", "=")
    
    Call Me.objTransData.ExecuteDatabase(strSql, dblUpdateQty, Me.objError)
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "2800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getCustItems()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim objCust_item As cust_item
    Dim strSql As String
    
    strRoutine_nm = "cust_items.cls.GetCustItems"
    
    strSql = "SELECT " & _
             " a.* " & _
             ",b.item_nbr " & _
             ",b.item_desc  " & _
             "FROM cust_item a, item b " & _
             "WHERE " & fWhereVariable("cust_id", Me.custId, "N", "=") & _
             " AND a.item_id = b.item_id " & _
             " AND a.deleted_fl <> 'Y' "
    
    Call Me.objTransData.OpenRecordset(objRecordset, strSql, objError)
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set Me.objRecordset = objRecordset
    End If
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "3000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
Public Sub getItemCusts()
    
    On Error GoTo error_handler
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim objCust_item As cust_item
    Dim strSql As String
    
    strRoutine_nm = "cust_item.cls.GetItemCusts"
    
    strSql = "SELECT " & _
             " a.item_id " & _
             ",a.cust_id " & _
             ",a.cust_item_nbr " & _
             ",a.cust_item_desc " & _
             ",a.cust_item_price " & _
             ",a.initial_purchase_dt " & _
             ",b.cust_nbr " & _
             ",a.allow_qty_disc_fl " & _
             "FROM cust_item a, customer b " & _
             "WHERE " & fWhereVariable("a.item_id", Me.itemId, "N", "=") & _
             "AND a.cust_id = b.cust_id " & _
             "AND a.deleted_fl <> 'Y' " & _
             "ORDER BY b.cust_nbr "
     
    Call Me.objTransData.OpenRecordset(objRecordset, strSql, objError)
    
    If Me.objError.p_type_cd = "E" Or Me.objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set Me.objRecordset = objRecordset
    End If
    
    Exit Sub
error_handler:
    With Me.objError
      .p_type_cd = "F"
      .p_err_cd = "3200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

