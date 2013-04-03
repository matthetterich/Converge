VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_item_id As Variant
Private mvarp_mat_req_id As Variant
Private mvarp_alloc_qty As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_alloc_priority_nbr As Variant
'

Public Property Let p_alloc_priority_nbr(ByVal vData As Variant)
    mvarp_alloc_priority_nbr = vData
End Property
Public Property Get p_alloc_priority_nbr() As Variant
        p_alloc_priority_nbr = mvarp_alloc_priority_nbr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_mat_req_id(ByVal vData As Variant)
    mvarp_mat_req_id = vData
End Property
Public Property Get p_mat_req_id() As Variant
        p_mat_req_id = mvarp_mat_req_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_alloc_qty(ByVal vData As Variant)
    mvarp_alloc_qty = vData
End Property
Public Property Get p_alloc_qty() As Variant
        p_alloc_qty = mvarp_alloc_qty
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
End Property


'//////////////////////////////////////////////////////////////////////////////
'Description:   Sets the item_id that is passed in by a property
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Sets the o_trans_data object up for processing transactions
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Sets the error object up for reporting error encountered
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The procedure will delete all on order inventory allocations
'               for an item and reallocate all material requirements that were allocated prior
'               to the delete
'  Arguments:   None
'//////////////////////////////////////////////////////////////////////////////
Public Sub ReallocateItemOnOrdInv()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    
    strRoutine_nm = "inventory_util.cls.ReallocateItemOnOrdInv"
        
    Call SelOnOrdInvAllocs(p_item_id, objrecordset)
    Call DelOnOrdInvAllocs(p_item_id)
    Call Allocate(objrecordset)
        
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    objrecordset.Close
    Set objrecordset = Nothing
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   The procedure will delete all inventory and on order inventory allocations
'               for an item and reallocate all material requirements that were allocated prior
'               to the delete
'  Arguments:   None
'//////////////////////////////////////////////////////////////////////////////
Public Sub ReallocateItem()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim objOnOrdRecordSet As ADODB.Recordset
    Dim objMatReq As Object
    
    strRoutine_nm = "inventory_util.cls.ReallocateItem"
    
    If Me.p_alloc_priority_nbr >= 0 And Len(Me.p_alloc_priority_nbr) > 0 And (Me.p_alloc_priority_nbr) = True Then
        Set objMatReq = CreateObject("converge_prod.mat_req")
        Call objMatReq.Change( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_mat_req_id _
                , "MR" _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Me.p_alloc_priority_nbr _
                , Empty _
                , Empty _
                )
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
           Exit Sub
        End If
    End If
    
    Call SelectMatReqs(p_item_id, objrecordset)
    Call DelInvAllocs(p_item_id)
    Call DelOnOrdInvAllocs(p_item_id)
    Call Allocate(objrecordset)
        
    objrecordset.Close
    Set objrecordset = Nothing
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelOnOrdInvAllocs(ByVal v_dblItemId As Double, ByRef r_objRecordset As ADODB.Recordset)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    
    strRoutine_nm = "inventory_util.cls.SelOnOrdInvAllocs"
    
    strSql = "SELECT " & _
             " mat_req_id, item_id, inv_loc_id, " & _
             " required_qty, " & _
             " ready_qty, " & _
             " consumed_qty " & _
             " FROM mat_req " & _
             " WHERE " & fWhereVariable("item_id", v_dblItemId, "N", "=") & _
             " AND status_cd in ('OP', 'BO', 'RD', 'CP') AND allocated_fl = 'Y' AND consumed_qty < required_qty " & _
             " order by alloc_priority_nbr desc, required_dt "
           
    Call o_trans_data.OpenRecordset(r_objRecordset, strSql, o_error)
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelectMatReqs(ByVal v_dblItemId As Double, ByRef r_objRecordset As Object)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    
    strRoutine_nm = "inventory_util.cls.SelectMatReqs"
    
    strSql = "SELECT " & _
             " mat_req_id, " & _
             " item_id, " & _
             " inv_loc_id, " & _
             " required_qty, " & _
             " ready_qty, " & _
             " consumed_qty " & _
             " FROM mat_req " & _
             " WHERE " & fWhereVariable("item_id", v_dblItemId, "N", "=") & _
             " AND status_cd in ('OP', 'BO', 'RD', 'CP') AND allocated_fl = 'Y' and consumed_qty < required_qty " & _
             " order by alloc_priority_nbr desc, required_dt "
           
    Call o_trans_data.OpenRecordset(r_objRecordset, strSql, o_error)
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DelOnOrdInvAllocs(ByVal v_dblItemId As Double)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "inventory_util.cls.DelOnOrdInvAllocs"
    
    strSql = "DELETE FROM on_ord_inv_alloc " & _
             "WHERE " & fWhereVariable("item_id", v_dblItemId, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DelInvAllocs(ByVal v_dblItemId As Double)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "inventory_util.cls.DelOnOrdInvAllocs"
    
    strSql = "DELETE FROM inventory_alloc " & _
           " WHERE " & fWhereVariable("item_id", v_dblItemId, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'               o_trans_data
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Allocate(ByRef r_objRecordset As ADODB.Recordset)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objInventoryAlloc As Object
    
    strRoutine_nm = "inventory_util.cls.Allocate"
    
    Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
    Set objInventoryAlloc.o_error = o_error
    Set objInventoryAlloc.o_trans_data = o_trans_data

    Do Until r_objRecordset.EOF = True Or o_error.type_cd = "E" Or o_error.type_cd = "F"
        objInventoryAlloc.p_item_id = r_objRecordset("item_id")
        objInventoryAlloc.p_inv_loc_id = r_objRecordset("inv_loc_id")
        objInventoryAlloc.p_mat_req_id = r_objRecordset("mat_req_id")
        objInventoryAlloc.p_alloc_qty = r_objRecordset("required_qty") - r_objRecordset("consumed_qty")
        Call objInventoryAlloc.Allocate
        r_objRecordset.MoveNext
    Loop
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DelMatReqInvAlloc(ByVal v_dblMat_req_id As Double)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "inventory_util.cls.DelOnOrdInvAllocs"
    
    strSql = "DELETE FROM inventory_alloc " & _
           " WHERE " & fWhereVariable("mat_req_id", v_dblMat_req_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DelMatReqOnOrdAlloc(ByVal v_dblMat_req_id As Double)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "inventory_util.cls.DelMatReqOnOrdInvAlloc"
    
    strSql = "DELETE FROM on_ord_inv_alloc " & _
             " WHERE " & fWhereVariable("mat_req_id", v_dblMat_req_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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




'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   dblItemId
'               o_error
'               o_trans_data
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub AllocateMatReq(ByRef v_varItemId As Variant, ByRef v_varMatReqId As Variant, ByRef v_varInvLocId As Variant, ByRef v_varAllocQty As Variant)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objInventoryAlloc As Object
    
    strRoutine_nm = "inventory_util.cls.AllocateMatReq"
    
    Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
    Set objInventoryAlloc.o_error = o_error
    Set objInventoryAlloc.o_trans_data = o_trans_data

    objInventoryAlloc.p_item_id = v_varItemId
    objInventoryAlloc.p_inv_loc_id = v_varInvLocId
    objInventoryAlloc.p_mat_req_id = v_varMatReqId
    objInventoryAlloc.p_alloc_qty = v_varAllocQty
    Call objInventoryAlloc.Allocate
    
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




