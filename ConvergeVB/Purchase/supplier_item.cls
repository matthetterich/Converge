VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "supplier_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"supplier_items"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_supplier_item_nbr As Variant
Private mvarp_supplier_id As Variant
Private mvarp_item_id As Variant
Private mvarp_supplier_item_desc As Variant
Private mvarp_supplier_item_cost As Variant
Private mvarp_last_cost As Variant
Private mvarp_allow_qty_disc_fl As Variant
Private mvarp_min_purchase_qty As Variant
Private mvarp_incr_purchase_qty As Variant
Private mvarp_bypass_getitem_fl As Variant
Private mvarp_bypass_getsupplier_fl As Variant
Private mvaro_supplier As Object
Private mvaro_item As Object
Private mvaro_record_set As ADODB.Recordset
'

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Set o_item(ByVal vData As Object)
    Set mvaro_item = vData
End Property
Public Property Get o_item() As Object
    Set o_item = mvaro_item
End Property
Public Property Set o_supplier(ByVal vData As Object)
    Set mvaro_supplier = vData
End Property
Public Property Get o_supplier() As Object
    Set o_supplier = mvaro_supplier
End Property
Public Property Let p_bypass_getsupplier_fl(ByVal vData As Variant)
    mvarp_bypass_getsupplier_fl = vData
End Property
Public Property Get p_bypass_getsupplier_fl() As Variant
        p_bypass_getsupplier_fl = mvarp_bypass_getsupplier_fl
End Property
Public Property Let p_bypass_getitem_fl(ByVal vData As Variant)
    mvarp_bypass_getitem_fl = vData
End Property
Public Property Get p_bypass_getitem_fl() As Variant
        p_bypass_getitem_fl = mvarp_bypass_getitem_fl
End Property
Public Property Let p_incr_purchase_qty(ByVal vData As Variant)
    mvarp_incr_purchase_qty = vData
End Property
Public Property Get p_incr_purchase_qty() As Variant
        p_incr_purchase_qty = mvarp_incr_purchase_qty
End Property
Public Property Let p_min_purchase_qty(ByVal vData As Variant)
    mvarp_min_purchase_qty = vData
End Property
Public Property Get p_min_purchase_qty() As Variant
        p_min_purchase_qty = mvarp_min_purchase_qty
End Property
Public Property Let p_allow_qty_disc_fl(ByVal vData As Variant)
    mvarp_allow_qty_disc_fl = vData
End Property
Public Property Get p_allow_qty_disc_fl() As Variant
        p_allow_qty_disc_fl = mvarp_allow_qty_disc_fl
End Property
Public Property Let p_supplier_item_cost(ByVal vData As Variant)
    mvarp_supplier_item_cost = vData
End Property
Public Property Get p_supplier_item_cost() As Variant
        p_supplier_item_cost = mvarp_supplier_item_cost
End Property

Public Property Let p_last_cost(ByVal vData As Variant)
    mvarp_last_cost = vData
End Property
Public Property Get p_last_cost() As Variant
        p_last_cost = mvarp_last_cost
End Property

Public Property Let p_supplier_item_desc(ByVal vData As Variant)
    mvarp_supplier_item_desc = vData
End Property
Public Property Get p_supplier_item_desc() As Variant
        p_supplier_item_desc = mvarp_supplier_item_desc
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
End Property
Public Property Let p_supplier_item_nbr(ByVal vData As Variant)
    mvarp_supplier_item_nbr = vData
End Property
Public Property Get p_supplier_item_nbr() As Variant
        p_supplier_item_nbr = mvarp_supplier_item_nbr
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Sub GetSupplierItem()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.GetSupplierItem"
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strSql = "SELECT *      "
    strSql = strSql & "FROM supplier_item "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
    If IsNull(p_item_id) = True Then
        strSql = strSql & "and" & fWhereVariable("supplier_item_nbr", p_supplier_item_nbr, "S", "=")
    Else
        strSql = strSql & "and" & fWhereVariable("item_id", p_item_id, "N", "=")
    End If
    
     Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If objRecordset.EOF = True Then
        p_supplier_id = Null
        p_item_id = Null
        
        ' Close the Record Set
        objRecordset.Close
        Set objRecordset = Nothing
       
    Else
        p_supplier_id = objRecordset("supplier_id")
        p_item_id = objRecordset("item_id")
        p_supplier_item_nbr = objRecordset("supplier_item_nbr")
        p_supplier_item_desc = objRecordset("supplier_item_desc")
        p_supplier_item_cost = objRecordset("supplier_item_cost")
        p_last_cost = objRecordset("last_cost")
        p_allow_qty_disc_fl = objRecordset("allow_qty_disc_fl")
        p_min_purchase_qty = objRecordset("min_purchase_qty")
        p_incr_purchase_qty = objRecordset("incr_purchase_qty")
        
        objRecordset.Close
        Set objRecordset = Nothing
        
        If p_bypass_getitem_fl = "Y" Then
        Else
            Set o_item = CreateObject("converge_item.item")
            Set o_item.o_error = o_error
            Set o_item.o_trans_data = o_trans_data
            o_item.p_item_id = p_item_id
            Call o_item.GetItem
        End If
    
        If p_bypass_getsupplier_fl = "Y" Then
        Else
            Set o_supplier = CreateObject("converge_purchase.supplier")
            Set o_supplier.o_error = o_error
            Set o_supplier.o_trans_data = o_trans_data
            o_supplier.p_supplier_id = p_supplier_id
            Call o_supplier.GetSupplier
        End If
    End If
  
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.Add"
    
    Call Validate("ADD")
    Call IssueInsert
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.Change"
    
    Call Validate("change")
    Call IssueUpdate
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.DeleteComment"
    
    Call IssueDelete
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub Validate(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.Validate"
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset

    If fValuePresent(p_item_id) = False Then
        If fValuePresent(o_item.p_item_nbr) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "0500"
            o_error.p_message_id = 2090
            o_error.p_routine_nm = strRoutine_nm
            Exit Sub
        End If
            
        o_item.p_item_id = Null
        Set o_item.o_error = o_error
        Set o_item.o_trans_data = o_trans_data
        Call o_item.GetItem
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        Else
            If IsNull(o_item.p_item_id) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "0600"
                o_error.p_message_id = 2090
                o_error.p_routine_nm = strRoutine_nm
                Exit Sub
            Else
                p_item_id = o_item.p_item_id
            End If
        End If
    End If
    
    If fValuePresent(p_supplier_id) = False Then
        If fValuePresent(o_supplier.p_supplier_nbr) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "0500"
            o_error.p_message_id = 2100
            o_error.p_routine_nm = strRoutine_nm
            Exit Sub
        End If
        
        o_supplier.p_supplier_id = Null
        Set o_supplier.o_error = o_error
        Set o_supplier.o_trans_data = o_trans_data
        Call o_supplier.GetSupplier
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        Else
            If IsNull(o_supplier.p_supplier_id) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "0600"
                o_error.p_message_id = 2100
                o_error.p_routine_nm = strRoutine_nm
                Exit Sub
            Else
                p_supplier_id = o_supplier.p_supplier_id
            End If
        End If
    End If
    
    If Len(p_supplier_item_nbr) > 0 Then
    Else
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0550"
        o_error.p_message_id = 3311
        o_error.p_routine_nm = strRoutine_nm
        Exit Sub
    End If
    
    If sFunction_cd = "ADD" Then
    
        
        strSql = "SELECT *      "
        strSql = strSql & "FROM supplier_item "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
        strSql = strSql & "and" & fWhereVariable("item_id", p_item_id, "N", "=")
        
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
            
        If objRecordset.EOF = True Then
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "0700"
            o_error.p_message_id = 2150
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_nbr = Err.Number
            o_error.p_desc = Err.Description
            Exit Sub
        End If
            
        objRecordset.Close
        Set objRecordset = Nothing
    End If

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_item.cls.IssueInsert"
    Dim dblUpdatedQty As Double
    Dim strSql As String
    '''''p_supplier_item_cost = FormatNumber(p_supplier_item_cost, 5)

    strSql = "insert into supplier_item ( "
    strSql = strSql & "  supplier_id "
    strSql = strSql & ", item_id "
    strSql = strSql & ", supplier_item_nbr "
    strSql = strSql & ", supplier_item_desc "
    strSql = strSql & ", supplier_item_cost "
    strSql = strSql & ", allow_qty_disc_fl "
    strSql = strSql & ", min_purchase_qty "
    strSql = strSql & ", incr_purchase_qty "
    strSql = strSql & ", last_cost "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_cost, "N")
    strSql = strSql & "," & fInsertVariable(p_allow_qty_disc_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_min_purchase_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_incr_purchase_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_last_cost, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    
   Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier_item.cls.IssueUpdate"
    Dim strSql As String

    strSql = "UPDATE supplier_item set "
    strSql = strSql & " " & fUpdateVariable("supplier_item_nbr", p_supplier_item_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_item_desc", p_supplier_item_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_item_cost", p_supplier_item_cost, "N", "=")
    strSql = strSql & "," & fUpdateVariable("allow_qty_disc_fl", p_allow_qty_disc_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("min_purchase_qty", p_min_purchase_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("incr_purchase_qty", p_incr_purchase_qty, "N", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("item_id", p_item_id, "N", "=")
     
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier_item.cls.IssueDelete"
    Dim strSql As String

    strSql = "delete from supplier_item "
    strSql = strSql & "where "
    strSql = strSql & "   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub GetSupplierItems()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strRoutine_nm = "supplier_item.cls.getSupplierItems"
    
    strSql = "SELECT a.*,  " & _
                    " b.item_nbr,  " & _
                    " b.item_desc " & _
             "  FROM supplier_item a, item b " & _
             " WHERE " & fWhereVariable("supplier_id", p_supplier_id, "N", "=") & _
             "   AND a.item_id = b.item_id " & _
             " ORDER BY b.item_nbr "
                
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
    End If

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub getItemSuppliers()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "supplier_item.cls.GetItemSuppliers"
    
    strSql = "SELECT " & _
                   " a.*, " & _
                   " b.supplier_nbr " & _
             "  FROM supplier_item a, supplier b" & _
             " WHERE " & fWhereVariable("a.item_id", p_item_id, "N", "=") & _
             "   AND a.supplier_id = b.supplier_id " & _
             "   ORDER BY b.supplier_nbr "

    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        'objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub



