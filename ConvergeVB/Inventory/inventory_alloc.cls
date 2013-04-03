VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_alloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_inventory_alloc_id As Variant
Private mvarp_item_id As Variant
Private mvarp_mat_req_id As Variant
Private mvarp_alloc_qty As Variant
Private mvarp_adj_qty As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_unallocated_qty As Variant
Private mvaro_trans_data As Object
Private mvaro_error As Object
Private mvaro_record_set As ADODB.Recordset
'
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Inventory Allocation Identifier
'               Return to show if any was unallocated
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_unallocated_qty(ByVal vData As Variant)
    mvarp_unallocated_qty = vData
End Property
Public Property Get p_unallocated_qty() As Variant
        p_unallocated_qty = mvarp_unallocated_qty
    
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Inventory Allocation Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_inventory_alloc_id(ByVal vData As Variant)
    mvarp_inventory_alloc_id = vData
End Property
Public Property Get p_inventory_alloc_id() As Variant
        p_inventory_alloc_id = mvarp_inventory_alloc_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for error Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Transaction Data Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Adjustment Quantity
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_adj_qty(ByVal vData As Variant)
    mvarp_adj_qty = vData
End Property
Public Property Get p_adj_qty() As Variant
    p_adj_qty = mvarp_adj_qty
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Allocation Quantity
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_alloc_qty(ByVal vData As Variant)
    mvarp_alloc_qty = vData
End Property
Public Property Get p_alloc_qty() As Variant
    p_alloc_qty = mvarp_alloc_qty
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Inventory Location Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
    p_inv_loc_id = mvarp_inv_loc_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Sales Order Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_mat_req_id(ByVal vData As Variant)
    mvarp_mat_req_id = vData
End Property
Public Property Get p_mat_req_id() As Variant
    p_mat_req_id = mvarp_mat_req_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   and Get for Item Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = mvarp_item_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Allocate inventory and on order inventory
'  Arguments:   None
'      Return:  Unallocate_qty will be used
'//////////////////////////////////////////////////////////////////////////////
Public Sub Allocate()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_alloc.cls.Allocate"
    
    Dim dblAvailableQty As Double
    Dim objInventoryAdj As inventory_adj
    Dim objOnOrdInvAlloc As on_ord_inv_alloc
    Dim objMatReq As Object
    
    Set objOnOrdInvAlloc = New on_ord_inv_alloc
    
                    'deletealocations delete both inv allocs and on ord inv allocs
                    'then it calls mat_req.updatequantities.
    Call DeleteAllocation
    
                    'get the material requirment to determine how much to allocate.
                    'quantities are acurate because the previous call updated the
                    'quantities on the mat req.
    Set objMatReq = CreateObject("converge_prod.mat_req")
    Set objMatReq.o_error = Me.o_error
    Set objMatReq.o_trans_data = Me.o_trans_data
    objMatReq.mat_req_id = Me.p_mat_req_id
    Call objMatReq.getMatReq
                    
                    'delete on order inventory allocation
    Set objOnOrdInvAlloc.o_error = Me.o_error
    Set objOnOrdInvAlloc.o_trans_data = Me.o_trans_data
    objOnOrdInvAlloc.p_mat_req_id = p_mat_req_id
    objOnOrdInvAlloc.p_item_id = p_item_id
    objOnOrdInvAlloc.p_inv_loc_id = p_inv_loc_id
    Call objOnOrdInvAlloc.Deallocate
    
                    'if material requirement is not active
                    'don't allocate any inventory to it
    With objMatReq
        If .status_cd = "CA" Or _
           .status_cd = "CL" Or _
           .status_cd = "OH" Then
            Call .UpdateQuantities( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.p_mat_req_id _
                            )
            Set objMatReq = Nothing
            Exit Sub
        End If
    End With
                    'update the allocated flag to a y so that we
                    'know this material requirement has been allocated, even
                    'if allocated amount is 0.
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
                , Empty _
                , "Y" _
                , Empty _
                )
        
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Me.p_item_id = objMatReq.item_id
    Me.p_alloc_qty = objMatReq.required_qty - objMatReq.consumed_qty
    Me.p_unallocated_qty = 0
    
                    'if nothing needs to be allocated, exit sub.
                    'call update quantities before exiting.  anytime
                    'allocate is called, quantities must be updated
    If Me.p_alloc_qty <= 0 Then
        Call objMatReq.UpdateQuantities( _
                          Me.o_error _
                        , Me.o_trans_data _
                        , Me.p_mat_req_id _
                        )
        Set objMatReq = Nothing
        Exit Sub
    End If
    
    Set objInventoryAdj = New inventory_adj
    dblAvailableQty = objInventoryAdj.AvailableQty( _
                                  Me.p_inv_loc_id _
                                , Me.p_item_id _
                                , Me.o_error _
                                , Me.o_trans_data _
                                )
       
                    'allocate the inventory and set
                    'unallocated quantity.
                    
                    '==> if no inventory is available, set the unallocated
                    '    quantity equal to the requested allocated quantity
                    '==> if the requested quantity is less than the available
                    '    quantity, allocate the requested quantity and set
                    '    unallocated quantity equal to zero
                    '==> if the requested quantity is greater than the available
                    '    quantity, allocate the available quantity and set the
                    '    unallocated quantity equal the requested quantity minus
                    '    the available quantity
    If dblAvailableQty <= 0 Then
        Me.p_unallocated_qty = Me.p_alloc_qty
    Else
        If Me.p_alloc_qty <= dblAvailableQty Then
            Call InsertAllocation(Me.p_alloc_qty)
            Me.p_unallocated_qty = 0
        Else
           Call InsertAllocation(dblAvailableQty)
           Me.p_unallocated_qty = Me.p_alloc_qty - dblAvailableQty
        End If
    End If
    
                    'if the unallocated quantity is greater than
                    'zero, try to allocate on order inventory
    If Me.p_unallocated_qty > 0 Then
        Set objOnOrdInvAlloc.o_error = Me.o_error
        Set objOnOrdInvAlloc.o_trans_data = Me.o_trans_data
        objOnOrdInvAlloc.p_mat_req_id = Me.p_mat_req_id
        objOnOrdInvAlloc.p_item_id = Me.p_item_id
        objOnOrdInvAlloc.p_inv_loc_id = Me.p_inv_loc_id
        objOnOrdInvAlloc.p_alloc_qty = Me.p_unallocated_qty
        Call objOnOrdInvAlloc.Allocate
        Me.p_unallocated_qty = objOnOrdInvAlloc.p_unallocated_qty
    End If
    
                    'now that the allocation is complete, update quantities
                    'on the material requirement.  Anytime an allocated is
                    'requested for a material requirement, update quantities
                    'must be ran
    Call objMatReq.UpdateQuantities( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_mat_req_id _
                    )
    Set objMatReq = Nothing
    Set objOnOrdInvAlloc = Nothing
    
    Exit Sub
error_handler:
    With Me.o_error
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
'Description:   Function to De Allocate inventory and on order inventory
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Deallocate()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_alloc.cls.DeAllocate"
    
    Dim strSql As String, _
        dblUpdateQty As Double, _
        objMatReq As Object, _
        objOnOrdInvAlloc As on_ord_inv_alloc, _
        objInventoryUtils As inventory_util
        
  
                    'delete any current inventory allocations
    Call DeleteAllocation
    
    
                    'delete any current on order inventory allocations
    Set objOnOrdInvAlloc = New on_ord_inv_alloc
    Set objOnOrdInvAlloc.o_error = Me.o_error
    Set objOnOrdInvAlloc.o_trans_data = Me.o_trans_data
    objOnOrdInvAlloc.p_mat_req_id = p_mat_req_id
    objOnOrdInvAlloc.p_item_id = p_item_id
    objOnOrdInvAlloc.p_inv_loc_id = p_inv_loc_id
    Call objOnOrdInvAlloc.Deallocate
    Set objOnOrdInvAlloc = Nothing
    
                    'update the quantities on the material requirement
                    'anytime an allocation is deleted, quantities need
                    'to be updated on the material requirement
    Set objMatReq = CreateObject("converge_prod.mat_req")
    Set objMatReq.o_error = Me.o_error
    Set objMatReq.o_trans_data = Me.o_trans_data
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
                , Empty _
                , "N" _
                , Empty _
                )
    
    Call objMatReq.UpdateQuantities( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_mat_req_id _
                    )
                    
    Set objInventoryUtils = New inventory_util
    objInventoryUtils.p_item_id = Me.p_item_id
    Set objInventoryUtils.o_error = Me.o_error
    Set objInventoryUtils.o_trans_data = Me.o_trans_data
    Call objInventoryUtils.ReallocateItemOnOrdInv
    
    Set objMatReq = Nothing
    
    Exit Sub
error_handler:
    With Me.o_error
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
'Description:   this routine will adjust the current allocation up or down based on
'               the adjustment quantity sent to the method.  If it is being adjusted up
'               available inventory will be validated.  IF not allocation record currently
'               exist, AdjustAllocation will do nothing
'  Arguments:   None

'//////////////////////////////////////////////////////////////////////////////
Public Sub AdjustAllocation()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_alloc.cls.adjustAllocation"
    
    Dim strSql As String, _
        dblAvailableQty As Double, _
        dblAllocQty As Double, _
        dblCurrentAllocQty As Double, _
        dblUpdateQty As Double, _
        objMatReq As Object
                 
    Set objMatReq = CreateObject("converge_prod.mat_req")
                    
                    'update the material requirements quantities
                    'to reflect the most recent inventory adjustment
    Call objMatReq.UpdateQuantities( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_mat_req_id _
                    )
    
                    'get the material requirement's allocated flag
    Set objMatReq.o_error = Me.o_error
    Set objMatReq.o_trans_data = Me.o_trans_data
    objMatReq.p_mat_req_id = p_mat_req_id
    Call objMatReq.getMatReq
    
                    'if the material requirement has been allocated,
                    'send it through the allocation routine again.
                    '==>If inventory has been consumed, the allocation will be
                    'reduced.
                    '==>If inventory has been returned, the allocation will be increased.
    If objMatReq.p_allocated_fl = "Y" Then
        Call Me.Allocate
    End If
    
    Set objMatReq = Nothing
    
    Exit Sub
error_handler:
    With Me.o_error
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
'Description:   Function to Allocate inventory and on order inventory
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub InsertAllocation(ByVal v_dblAllocQty As Double)

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "inventory_alloc.cls.InsertAllocation"

    Me.p_inventory_alloc_id = fGetId("inventory_alloc_id", Me.o_error, Me.o_trans_data)
        
    strSql = " INSERT INTO inventory_alloc ( " & _
                  "  inventory_alloc_id" & _
                  ", inv_loc_id" & _
                  ", item_id" & _
                  ", mat_req_id" & _
                  ", alloc_qty" & _
                  ", updated_assoc_id" & _
                  ", updated_ts" & _
                  ", entered_assoc_id" & _
                  ", entered_ts" & _
                  " ) " & _
                  "VALUES (" & _
                  "  " & fInsertVariable(Me.p_inventory_alloc_id, "N") & _
                  ", " & fInsertVariable(Me.p_inv_loc_id, "S") & _
                  ", " & fInsertVariable(Me.p_item_id, "N") & _
                  ", " & fInsertVariable(Me.p_mat_req_id, "N") & _
                  ", " & fInsertVariable(v_dblAllocQty, "N")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N") & _
                  "," & fInsertVariable("current_ts", "D") & _
                  "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N") & _
                  "," & fInsertVariable("current_ts", "D") & _
                  ")"
        
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, Me.o_error)

    Exit Sub
error_handler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub DeleteAllocation()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "inventory_alloc.cls.DeleteAllocation"

    strSql = "DELETE FROM inventory_alloc " & _
           " WHERE " & fWhereVariable("mat_req_id", Me.p_mat_req_id, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, Me.o_error)
    
    Exit Sub
error_handler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Function fGetAllocatedQty() As Double

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "inventory_alloc.cls.fGetAllocatedQty"

    strSql = "Select alloc_qty FROM inventory_alloc" & _
           "  WHERE " & fWhereVariable("mat_req_id", Me.p_mat_req_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If

    If objrecordset.EOF = True Then
        fGetAllocatedQty = 0
    Else
        fGetAllocatedQty = objrecordset(0)
    End If
     
    Set Me.o_record_set = objrecordset

    Exit Function
error_handler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


'//////////////////////////////////////////////////////////////////////////////
'Description:   Gets the inventory allocation class
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetInventoryAllocs()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_allocs.cls.getInventoryallocs"
    
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT " & _
                " a.inventory_alloc_id, " & _
                " a.alloc_qty, " & _
                " a.mat_req_id, " & _
                " a.inv_loc_id, " & _
                " b.assem_ord_id, " & _
                " b.sales_ord_id, " & _
                " c.item_nbr, " & _
                " c.item_desc, " & _
                " b.required_dt, e.cust_nbr " & _
                " FROM sales_order d LEFT OUTER JOIN customer e ON d.cust_id = e.cust_id RIGHT OUTER Join " & _
                " mat_req b ON d.sales_ord_id = b.sales_ord_id LEFT OUTER JOIN item c ON b.item_id = c.item_id RIGHT OUTER Join inventory_alloc a ON b.mat_req_id = a.mat_req_id " & _
                " WHERE  a.alloc_qty > 0 "
                
    If fValuePresent(p_item_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If fValuePresent(p_inv_loc_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
    If fValuePresent(p_mat_req_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.mat_req_id", p_mat_req_id, "N", "=")
    End If
    
    strSql = strSql & " ORDER BY b.required_dt "
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set Me.o_record_set = objrecordset
    End If

    Exit Sub
error_handler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

    
