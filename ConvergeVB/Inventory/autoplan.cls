VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "autoplan"
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
'

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
'Description:   Selects and delete on order alllocation
'  Arguments:   None
'//////////////////////////////////////////////////////////////////////////////
Public Sub ReallocateItem()
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "purchase_order_item.cls.GetOnOrderInventoryAllocation"
        
    Call SelectOnOrderInventoryAllocation(p_item_id, objRecordset, o_error, o_trans_data)

    Call DeleteOnOrderInventoryAllocation(p_item_id, o_error, o_trans_data)
    
    Call Allocation(p_item_id, objRecordset, o_error, o_trans_data)
        
    Err.Clear
        
End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   intItemId
'               o_error
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelectOnOrderInventoryAllocation(ByVal intItemId As Integer, ByRef o_error As Object, ByRef objRecordset As Object, ByRef o_trans_data As Object)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    ' Define
    Dim strRoutine_nm As String
    Dim strSql As String
    
    ' Initialize
    strRoutine_nm = "autoplan.cls.SelectOnOrderInventoryAllocation"
    
    strSql = "SELECT " & _
                " a.*, " & _
                " b.required_qty, " & _
                " b.ready_qty, " & _
                " b.consumed_qty, " & _
            " FROM on_ord_inv_alloc a, mat_req b" & _
           " WHERE " & fWhereVariable("item_id", intItemId, "N", "=") & _
           "   AND a.mat_req_id = b.mat_req_id "
           
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    Call Err.Clear

End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   intItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DeleteOnOrderInventoryAllocation(ByVal intItemId As Integer, ByRef o_error As Object, ByRef o_trans_data As Object)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    
    ' Define
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim intUpdatedQty As Integer
    
    ' Initialize
    strRoutine_nm = "autoplan.cls.DeleteOnOrderInventoryAllocation"
    
    strSql = "DELETE FROM on_ord_inv_alloc " & _
           " WHERE " & fWhereVariable("item_id", intItemId, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, intUpdatedQty, o_error)
    
    Call Err.Clear

End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   intItemId
'               o_error
'               o_trans_data
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Allocation(ByVal intItemId As Integer, ByRef objRecordset As Object, ByRef o_error As Object, ByRef o_trans_data As Object)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    ' Define
    Dim strRoutine_nm As String
    Dim ObjOnOrdInvAlloc As Object
    Dim objInventoryAlloc As Object
    
    ' Initialize
    Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
    Set objInventoryAlloc.o_error = o_error
    Set objInventoryAlloc.o_trans_data = o_trans_data
    Set ObjOnOrdInvAlloc = CreateObject("converge_inv.on_ord_inv_alloc")
    Set ObjOnOrdInvAlloc.o_error = o_error
    Set ObjOnOrdInvAlloc.o_trans_data = o_trans_data
    strRoutine_nm = "autoplan.cls.DeleteOnOrderInventoryAllocation"
    objInventoryAlloc.p_alloc_fl = "Y"
    ObjOnOrdInvAlloc.p_alloc_fl = "Y"

    Do Until objRecordset.EOF = True Or o_error.type_cd = "E" Or o_error.type_cd = "F"
        If ObjOnOrdInvAlloc.p_alloc_fl = "Y" Then
           If objInventoryAlloc.p_alloc_fl = "N" Then
               ObjOnOrdInvAlloc.p_item_id = objRecordset("item_id")
               ObjOnOrdInvAlloc.p_inv_loc_id = objRecordset("inv_loc_id")
               ObjOnOrdInvAlloc.p_mat_req_id = objRecordset("mat_req_id")
               ObjOnOrdInvAlloc.p_alloc_qty = objRecordset("required_qty") - objRecordset("consumed_qty")
               Call ObjOnOrdInvAlloc.Allocate
           Else
               Call PlanAllocate(objInventoryAlloc, objRecordset("item_id"), objRecordset("inv_loc_id"), objRecordset("mat_req_id"), objRecordset("required_qty") - objRecordset("consumed_qty"))
           End If
        Else
            Call PlanAllocate(objInventoryAlloc, objRecordset("item_id"), objRecordset("inv_loc_id"), objRecordset("mat_req_id"), objRecordset("required_qty") - objRecordset("consumed_qty"))
            ' Call Material Requirment Update
        End If
    Loop
    
End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Allocate inventory and on order inventory
'  Arguments:   objInventoryAlloc
'               intItemId
'               strInvLocId
'               intAllocQty
'               intMatReqId
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub PlanAllocate(ByRef objInventoryAlloc As Object, ByVal intItemId As Integer, ByVal strInvLocId As String, ByVal intMatReqId As Integer, ByVal intAllocQty As Integer)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
  
    On Error Resume Next
    Dim ObjInventoryAdj As Object
    Dim ObjOnOrdInvAlloc As Object
    Dim DblAvailableQty As Double
    Dim DblOnOrderAvailableQty As Double
    Dim dblAllocQty As Double
    Dim strRoutine_nm As String
    
    strRoutine_nm = "inventory_alloc.cls.Allocate"
    Set ObjInventoryAdj = CreateObject("converge_inv.inventory_adj")
    
    objInventoryAlloc.o_error = o_error
    objInventoryAlloc.o_trans_data = o_trans_data
    objInventoryAlloc.p_item_id = intItemId
    objInventoryAlloc.p_inv_loc_id = strInvLocId
    objInventoryAlloc.p_mat_req_id = intMatReqId
    objInventoryAlloc.p_alloc_qty = intAllocQty
    Call objInventoryAlloc.Allocate
    
    If objInventoryAlloc.p_unallocated_qty > 0 Then
       objInventoryAlloc.p_alloc_fl = "N"
        ' Set up the On Order Inventory Allocation
        Set ObjOnOrdInvAlloc = CreateObject("converge_inv.on_ord_inv_alloc")
        ObjOnOrdInvAlloc.o_error = o_error
        ObjOnOrdInvAlloc.o_trans_data = o_trans_data
        ObjOnOrdInvAlloc.p_item_id = intItemId
        ObjOnOrdInvAlloc.p_mat_req_id = intMatReqId
        ObjOnOrdInvAlloc.p_alloc_qty = objInventoryAlloc.p_unallocated_qty
        Call ObjOnOrdInvAlloc.Allocate
    End If

End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Processes the priority allocations
'  Arguments:   None
'//////////////////////////////////////////////////////////////////////////////
Public Sub PriorityAllocation()
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    ' Define
    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim ObjOnOrderRecordSet As adodb.Recordset
    
    ' Initialize
    strRoutine_nm = "purchase_order_item.cls.GetOnOrderInventoryAllocation"
    
    Call SelectInventoryAllocation(p_item_id, objRecordset, o_error, o_trans_data)
    Call SelectOnOrderInventoryAllocation(p_item_id, ObjOnOrderRecordSet, o_error, o_trans_data)

    Call DeleteOnOrderInventoryAllocation(p_item_id, o_error, o_trans_data)
    Call DeleteInventoryAllocation(p_item_id, o_error, o_trans_data)
    
    Call Allocation(p_item_id, objRecordset, o_error, o_trans_data)
    Call Allocation(p_item_id, ObjOnOrderRecordSet, o_error, o_trans_data)
        
    Err.Clear
        
End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   intItemId
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DeleteInventoryAllocation(ByVal intItemId As Integer, ByRef o_error As Object, ByRef o_trans_data As Object)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    ' Define
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim intUpdatedQty As Integer
    
    ' Initialize
    strRoutine_nm = "autoplan.cls.DeleteOnOrderInventoryAllocation"
    
    strSql = "DELETE FROM inventory_alloc " & _
           " WHERE " & fWhereVariable("item_id", intItemId, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, intUpdatedQty, o_error)
    
    Call Err.Clear

End Sub

'//////////////////////////////////////////////////////////////////////////////
'Description:   Selects the on order inventory allocations by receipt_dt
'  Arguments:   intItemId
'               o_error
'               ObjRecordSet
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelectInventoryAllocation(ByVal intItemId As Integer, ByRef objRecordset As Object, ByRef o_error As Object, ByRef o_trans_data As Object)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    On Error Resume Next
    ' Define
    Dim strRoutine_nm As String
    Dim strSql As String
    
    ' Initialize
    strRoutine_nm = "autoplan.cls.SelectInventoryAllocation"
    
    strSql = "SELECT " & _
                " a.*, " & _
                " b.required_qty, " & _
                " b.ready_qty, " & _
                " b.consumed_qty, " & _
            " FROM on_ord_inv_alloc a, mat_req b" & _
           " WHERE " & fWhereVariable("item_id", intItemId, "N", "=") & _
           "   AND a.mat_req_id = b.mat_req_id "
           
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    Call Err.Clear

End Sub
