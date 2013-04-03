VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mat_req"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'/////////////////////////////////////////////////////////////////////////////
'(c) 2000 Gemini Software Solutions Inc. All Rights Reserved.
'
'This work is subject to U.S. and international copyright laws and treaties.
'
'No part of this work may be used, practiced, performed, copied, distributed,
'revised, modified, translated, abridged, condensed, expanded, collected,
'compiled, linked, recast, transformed or adapted without the prior written
'consent of GSS Inc.
'
'Any use or exploitation of this work without authorization could subject the
'perpetrator to criminal and civil liability.
'
'/////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'Description:   This class object is responsible for maintaining Material
'               requirements
'//////////////////////////////////////////////////////////////////////////////

Private m_varMat_req_id             As Variant, _
        m_varItem_id                As Variant, _
        m_varItem_nbr               As Variant, _
        m_varItem_type_cd           As Variant, _
        m_varItem_type_desc         As Variant, _
        m_varInv_loc_id             As Variant, _
        m_varInv_loc_nm             As Variant, _
        m_varRequired_qty           As Variant, _
        m_varRequired_dt            As Variant, _
        m_varRequired_begin_dt      As Variant, _
        m_varRequired_end_dt        As Variant, _
        m_varStatus_cd              As Variant, _
        m_varStatus_desc            As Variant, _
        m_varBack_ordered_qty       As Variant, _
        m_varReady_qty              As Variant, _
        m_varConsumed_qty           As Variant, _
        m_varUnplanned_qty          As Variant
Private m_varSales_ord_id           As Variant, _
        m_varSO_line_nbr            As Variant, _
        m_varAssem_ord_id           As Variant, _
        m_varAC_line_nbr            As Variant, _
        m_varRequired_qty_var_pct   As Variant, _
        m_varUOM_cd                 As Variant, _
        m_varUOM_desc               As Variant, _
        m_varSupplier_id            As Variant, _
        m_varSupplier_nbr           As Variant, _
        m_varAllocated_fl           As Variant, _
        m_varAlloc_priority_nbr     As Variant, _
        m_objError                  As Object, _
        m_objTransData              As Object, _
        m_rstRecordSet              As ADODB.RecordSet
                                        ' The component part being added to the
                                        ' parent part number.
Public Property Set o_trans_data(ByVal v_objTransData As Object)
    Set m_objTransData = v_objTransData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTransData
End Property

Public Property Set RecordSet(ByVal v_rstRecordSet As ADODB.RecordSet)
    Set m_rstRecordSet = v_rstRecordSet
End Property
Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = m_rstRecordSet
End Property

Public Property Set o_record_set(ByVal v_rstRecordSet As ADODB.RecordSet)
    Set m_rstRecordSet = v_rstRecordSet
End Property
Public Property Get o_record_set() As ADODB.RecordSet
    Set o_record_set = m_rstRecordSet
End Property

Public Property Set o_error(ByVal v_objError As Object)
    Set m_objError = v_objError
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property

Public Property Let mat_req_id(ByVal v_varMat_req_id As Variant)
    m_varMat_req_id = v_varMat_req_id
End Property
Public Property Get mat_req_id() As Variant
    mat_req_id = m_varMat_req_id
End Property
Public Property Let alloc_priority_nbr(ByVal vdata As Variant)
    m_varAlloc_priority_nbr = vdata
End Property
Public Property Get alloc_priority_nbr() As Variant
    alloc_priority_nbr = m_varAlloc_priority_nbr
End Property


Public Property Let allocated_fl(ByVal vdata As Variant)
    m_varAllocated_fl = vdata
End Property
Public Property Get allocated_fl() As Variant
    allocated_fl = m_varAllocated_fl
End Property

Public Property Let inv_loc_id(ByVal v_varInv_loc_id As Variant)
    m_varInv_loc_id = v_varInv_loc_id
End Property
Public Property Get inv_loc_id() As Variant
    inv_loc_id = m_varInv_loc_id
End Property
Public Property Let inv_loc_nm(ByVal v_varInv_loc_nm As Variant)
    m_varInv_loc_nm = v_varInv_loc_nm
End Property
Public Property Get inv_loc_nm() As Variant
    inv_loc_id = m_varInv_loc_nm
End Property
Public Property Let item_id(ByVal v_varItem_id As Variant)
    m_varItem_id = v_varItem_id
End Property
Public Property Get item_id() As Variant
    item_id = m_varItem_id
End Property
Public Property Let item_nbr(ByVal v_varItem_nbr As Variant)
    m_varItem_nbr = v_varItem_nbr
End Property
Public Property Get item_nbr() As Variant
    item_nbr = m_varItem_nbr
End Property
Public Property Let item_type_cd(ByVal v_varItem_type_cd As Variant)
    m_varItem_type_cd = v_varItem_type_cd
End Property
Public Property Get item_type_cd() As Variant
    item_type_cd = m_varItem_type_cd
End Property
Public Property Let item_type_desc(ByVal v_varItem_type_desc As Variant)
    m_varItem_type_desc = v_varItem_type_desc
End Property
Public Property Get item_type_desc() As Variant
    item_type_desc = m_varItem_type_desc
End Property
Public Property Let required_qty(ByVal v_varRequired_qty As Variant)
    m_varRequired_qty = v_varRequired_qty
End Property
Public Property Get required_qty() As Variant
    required_qty = m_varRequired_qty
End Property
Public Property Let required_dt(ByVal v_varRequired_dt As Variant)
    m_varRequired_dt = v_varRequired_dt
End Property
Public Property Get required_dt() As Variant
    required_dt = m_varRequired_dt
End Property
Public Property Let required_begin_dt(ByVal v_varRequired_begin_dt As Variant)
    m_varRequired_begin_dt = v_varRequired_begin_dt
End Property
Public Property Get required_begin_dt() As Variant
    required_begin_dt = m_varRequired_begin_dt
End Property
Public Property Let required_end_dt(ByVal v_varRequired_end_dt As Variant)
    m_varRequired_end_dt = v_varRequired_end_dt
End Property
Public Property Get required_end_dt() As Variant
    required_end_dt = m_varRequired_end_dt
End Property
Public Property Let status_cd(ByVal v_varStatus_cd As Variant)
    m_varStatus_cd = v_varStatus_cd
End Property
Public Property Get status_cd() As Variant
    status_cd = m_varStatus_cd
End Property
Public Property Let status_desc(ByVal v_varStatus_desc As Variant)
    m_varStatus_desc = v_varStatus_desc
End Property
Public Property Get status_desc() As Variant
    status_desc = m_varStatus_desc
End Property
Public Property Let back_ordered_qty(ByVal v_varBack_ordered_qty As Variant)
    m_varBack_ordered_qty = v_varBack_ordered_qty
End Property
Public Property Get back_ordered_qty() As Variant
    back_ordered_qty = m_varBack_ordered_qty
End Property
Public Property Let ready_qty(ByVal v_varReady_qty As Variant)
    m_varReady_qty = v_varReady_qty
End Property
Public Property Get ready_qty() As Variant
    ready_qty = m_varReady_qty
End Property
Public Property Let consumed_qty(ByVal v_varConsumed_qty As Variant)
    m_varConsumed_qty = v_varConsumed_qty
End Property
Public Property Get consumed_qty() As Variant
    consumed_qty = m_varConsumed_qty
End Property
Public Property Let unplanned_qty(ByVal v_varUnplanned_qty As Variant)
    m_varUnplanned_qty = v_varUnplanned_qty
End Property
Public Property Get unplanned_qty() As Variant
    unplanned_qty = m_varUnplanned_qty
End Property
Public Property Let sales_ord_id(ByVal v_varSales_ord_id As Variant)
    m_varSales_ord_id = v_varSales_ord_id
End Property
Public Property Get sales_ord_id() As Variant
    sales_ord_id = m_varSales_ord_id
End Property
Public Property Let so_line_nbr(ByVal v_varSo_line_nbr As Variant)
    m_varSO_line_nbr = v_varSo_line_nbr
End Property
Public Property Get so_line_nbr() As Variant
    so_line_nbr = m_varSO_line_nbr
End Property
Public Property Let assem_ord_id(ByVal v_varAssem_ord_id As Variant)
    m_varAssem_ord_id = v_varAssem_ord_id
End Property
Public Property Get assem_ord_id() As Variant
    assem_ord_id = m_varAssem_ord_id
End Property
Public Property Let ac_line_nbr(ByVal v_varAc_line_nbr As Variant)
    m_varAC_line_nbr = v_varAc_line_nbr
End Property
Public Property Get ac_line_nbr() As Variant
    ac_line_nbr = m_varAC_line_nbr
End Property
Public Property Let required_qty_var_pct(ByVal v_varRequired_qty_var_pct As Variant)
    m_varRequired_qty_var_pct = v_varRequired_qty_var_pct
End Property
Public Property Get required_qty_var_pct() As Variant
    required_qty_var_pct = m_varRequired_qty_var_pct
End Property
Public Property Let uom_cd(ByVal v_varUOM_cd As Variant)
    m_varUOM_cd = v_varUOM_cd
End Property
Public Property Get uom_cd() As Variant
    uom_cd = m_varUOM_cd
End Property
Public Property Let uom_desc(ByVal v_varUOM_desc As Variant)
    m_varUOM_desc = v_varUOM_desc
End Property
Public Property Get uom_desc() As Variant
    uom_desc = m_varUOM_desc
End Property
Public Property Let supplier_id(ByVal v_varSupplier_id As Variant)
    m_varSupplier_id = v_varSupplier_id
End Property
Public Property Get supplier_id() As Variant
    supplier_id = m_varSupplier_id
End Property
Public Property Let supplier_nbr(ByVal v_varSupplier_nbr As Variant)
    m_varSupplier_nbr = v_varSupplier_nbr
End Property
Public Property Get supplier_nbr() As Variant
    supplier_nbr = m_varSupplier_nbr
End Property

Public Property Let p_mat_req_id(ByVal v_varMat_req_id As Variant)
    m_varMat_req_id = v_varMat_req_id
End Property
Public Property Get p_mat_req_id() As Variant
    p_mat_req_id = m_varMat_req_id
End Property
Public Property Let p_alloc_priority_nbr(ByVal vdata As Variant)
    m_varAlloc_priority_nbr = vdata
End Property
Public Property Get p_alloc_priority_nbr() As Variant
    p_alloc_priority_nbr = m_varAlloc_priority_nbr
End Property
Public Property Let p_allocated_fl(ByVal vdata As Variant)
    m_varAllocated_fl = vdata
End Property
Public Property Get p_allocated_fl() As Variant
    p_allocated_fl = m_varAllocated_fl
End Property

Public Property Let p_inv_loc_id(ByVal v_varInv_loc_id As Variant)
    m_varInv_loc_id = v_varInv_loc_id
End Property
Public Property Get p_inv_loc_id() As Variant
    p_inv_loc_id = m_varInv_loc_id
End Property
Public Property Let p_inv_loc_nm(ByVal v_varInv_loc_nm As Variant)
    m_varInv_loc_nm = v_varInv_loc_nm
End Property
Public Property Get p_inv_loc_nm() As Variant
    p_inv_loc_id = m_varInv_loc_nm
End Property
Public Property Let p_item_id(ByVal v_varItem_id As Variant)
    m_varItem_id = v_varItem_id
End Property
Public Property Get p_item_id() As Variant
    p_item_id = m_varItem_id
End Property
Public Property Let p_item_nbr(ByVal v_varItem_nbr As Variant)
    m_varItem_nbr = v_varItem_nbr
End Property
Public Property Get p_item_nbr() As Variant
    p_item_nbr = m_varItem_nbr
End Property
Public Property Let p_item_type_cd(ByVal v_varItem_type_cd As Variant)
    m_varItem_type_cd = v_varItem_type_cd
End Property
Public Property Get p_item_type_cd() As Variant
    p_item_type_cd = m_varItem_type_cd
End Property
Public Property Let p_item_type_desc(ByVal v_varItem_type_desc As Variant)
    m_varItem_type_desc = v_varItem_type_desc
End Property
Public Property Get p_item_type_desc() As Variant
    p_item_type_desc = m_varItem_type_desc
End Property
Public Property Let p_required_qty(ByVal v_varRequired_qty As Variant)
    m_varRequired_qty = v_varRequired_qty
End Property
Public Property Get p_required_qty() As Variant
    p_required_qty = m_varRequired_qty
End Property
Public Property Let p_required_dt(ByVal v_varRequired_dt As Variant)
    m_varRequired_dt = v_varRequired_dt
End Property
Public Property Get p_required_dt() As Variant
    p_required_dt = m_varRequired_dt
End Property
Public Property Let p_required_begin_dt(ByVal v_varRequired_begin_dt As Variant)
    m_varRequired_begin_dt = v_varRequired_begin_dt
End Property
Public Property Get p_required_begin_dt() As Variant
    p_required_begin_dt = m_varRequired_begin_dt
End Property
Public Property Let p_required_end_dt(ByVal v_varRequired_end_dt As Variant)
    m_varRequired_end_dt = v_varRequired_end_dt
End Property
Public Property Get p_required_end_dt() As Variant
    p_required_end_dt = m_varRequired_end_dt
End Property
Public Property Let p_status_cd(ByVal v_varStatus_cd As Variant)
    m_varStatus_cd = v_varStatus_cd
End Property
Public Property Get p_status_cd() As Variant
    p_status_cd = m_varStatus_cd
End Property
Public Property Let p_status_desc(ByVal v_varStatus_desc As Variant)
    m_varStatus_desc = v_varStatus_desc
End Property
Public Property Get p_status_desc() As Variant
    p_status_desc = m_varStatus_desc
End Property
Public Property Let p_back_ordered_qty(ByVal v_varBack_ordered_qty As Variant)
    m_varBack_ordered_qty = v_varBack_ordered_qty
End Property
Public Property Get p_back_ordered_qty() As Variant
    p_back_ordered_qty = m_varBack_ordered_qty
End Property
Public Property Let p_ready_qty(ByVal v_varReady_qty As Variant)
    m_varReady_qty = v_varReady_qty
End Property
Public Property Get p_ready_qty() As Variant
    p_ready_qty = m_varReady_qty
End Property
Public Property Let p_consumed_qty(ByVal v_varConsumed_qty As Variant)
    m_varConsumed_qty = v_varConsumed_qty
End Property
Public Property Get p_consumed_qty() As Variant
    p_consumed_qty = m_varConsumed_qty
End Property
Public Property Let p_unplanned_qty(ByVal v_varUnplanned_qty As Variant)
    m_varUnplanned_qty = v_varUnplanned_qty
End Property
Public Property Get p_unplanned_qty() As Variant
    p_unplanned_qty = m_varUnplanned_qty
End Property
Public Property Let p_sales_ord_id(ByVal v_varSales_ord_id As Variant)
    m_varSales_ord_id = v_varSales_ord_id
End Property
Public Property Get p_sales_ord_id() As Variant
    p_sales_ord_id = m_varSales_ord_id
End Property
Public Property Let p_so_line_nbr(ByVal v_varSo_line_nbr As Variant)
    m_varSO_line_nbr = v_varSo_line_nbr
End Property
Public Property Get p_so_line_nbr() As Variant
    p_so_line_nbr = m_varSO_line_nbr
End Property
Public Property Let p_assem_ord_id(ByVal v_varAssem_ord_id As Variant)
    m_varAssem_ord_id = v_varAssem_ord_id
End Property
Public Property Get p_assem_ord_id() As Variant
    p_assem_ord_id = m_varAssem_ord_id
End Property
Public Property Let p_ac_line_nbr(ByVal v_varAc_line_nbr As Variant)
    m_varAC_line_nbr = v_varAc_line_nbr
End Property
Public Property Get p_ac_line_nbr() As Variant
    p_ac_line_nbr = m_varAC_line_nbr
End Property
Public Property Let p_required_qty_var_pct(ByVal v_varRequired_qty_var_pct As Variant)
    m_varRequired_qty_var_pct = v_varRequired_qty_var_pct
End Property
Public Property Get p_required_qty_var_pct() As Variant
    p_required_qty_var_pct = m_varRequired_qty_var_pct
End Property
Public Property Let p_uom_cd(ByVal v_varUOM_cd As Variant)
    m_varUOM_cd = v_varUOM_cd
End Property
Public Property Get p_uom_cd() As Variant
    p_uom_cd = m_varUOM_cd
End Property
Public Property Let p_uom_desc(ByVal v_varUOM_desc As Variant)
    m_varUOM_desc = v_varUOM_desc
End Property
Public Property Get p_uom_desc() As Variant
    p_uom_desc = m_varUOM_desc
End Property
Public Property Let p_supplier_id(ByVal v_varSupplier_id As Variant)
    m_varSupplier_id = v_varSupplier_id
End Property
Public Property Get p_supplier_id() As Variant
    p_supplier_id = m_varSupplier_id
End Property
Public Property Let p_supplier_nbr(ByVal v_varSupplier_nbr As Variant)
    m_varSupplier_nbr = v_varSupplier_nbr
End Property
Public Property Get p_supplier_nbr() As Variant
    p_supplier_nbr = m_varSupplier_nbr
End Property
                                        
Public Sub GetMatReq()
    
    On Error GoTo errHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    
    strRoutine_nm = "mat_req.cls.GetMatReq"
    
    Set objRecordset = New ADODB.RecordSet

    If Len(mat_req_id) > 0 Then
        strSql = "SELECT * " & _
                "FROM mat_req " & _
                " WHERE " & _
                "     mat_req_id = " & mat_req_id
    Else
        If Len(sales_ord_id) > 0 Then
            strSql = "SELECT * " & _
                    "FROM mat_req " & _
                    " WHERE " & _
                    "     sales_ord_id = " & sales_ord_id & _
                    " and so_line_nbr = " & so_line_nbr
        Else
            strSql = "SELECT * " & _
                    "FROM mat_req " & _
                    " WHERE " & _
                    "    assem_ord_id = " & assem_ord_id & _
                    " and ac_line_nbr = " & ac_line_nbr
        End If
    End If

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        Me.mat_req_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        Me.mat_req_id = objRecordset("mat_req_id")
        Me.item_id = objRecordset("item_id")
        Me.item_type_cd = objRecordset("item_type_cd")
        Me.inv_loc_id = objRecordset("inv_loc_id")
        Me.required_qty = objRecordset("required_qty")
        Me.required_dt = objRecordset("required_dt")
        Me.status_cd = objRecordset("status_cd")
        Me.back_ordered_qty = objRecordset("back_ordered_qty")
        Me.ready_qty = objRecordset("ready_qty")
        Me.consumed_qty = objRecordset("consumed_qty")
        Me.sales_ord_id = objRecordset("sales_ord_id")
        Me.so_line_nbr = objRecordset("so_line_nbr")
        Me.assem_ord_id = objRecordset("assem_ord_id")
        Me.ac_line_nbr = objRecordset("ac_line_nbr")
        Me.required_qty_var_pct = objRecordset("required_qty_var_pct")
        Me.uom_cd = objRecordset("uom_cd")
        Me.supplier_id = objRecordset("supplier_id")
        Me.alloc_priority_nbr = objRecordset("alloc_priority_nbr")
        Me.allocated_fl = objRecordset("allocated_fl")
        objRecordset.Close
        Set objRecordset = Nothing
        
        item_nbr = fGetItemNbr(item_id, o_error, o_trans_data)
        inv_loc_nm = fGetInvLocNm(inv_loc_id, o_error, o_trans_data)
        item_type_desc = fGetDesc("item_type_cd", item_type_cd, o_error, o_trans_data)
        status_desc = fGetDesc("mat_req_status_cd", status_cd, o_error, o_trans_data)
        uom_desc = fGetDesc("uom_cd", uom_cd, o_error, o_trans_data)
        supplier_nbr = fGetSupplierNbr(supplier_id, o_error, o_trans_data)
        unplanned_qty = required_qty - (back_ordered_qty + ready_qty + consumed_qty)
        If unplanned_qty < 0 Then
            unplanned_qty = 0
        End If
    End If
        
    Exit Sub
errHandler:
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

Public Sub GetMatReqs()
    
    On Error GoTo errHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    Dim varItemId           As Variant
    
    strRoutine_nm = "mat_req.cls.GetMatReqs"
    
    strSql = "SELECT * from mat_req a, item b, code_desc c " & _
             " WHERE (" & fWhereVariable("c.field_nm", "mat_req_status_cd", "S", "=") & _
               " AND c.code = a.status_cd)" & _
              " AND a.item_id = b.item_id "

    If fValuePresent(item_nbr) = True Then
        varItemId = fGetItemId(item_nbr, o_error, o_trans_data)
        strSql = strSql & " and " & fWhereVariable("a.item_id", varItemId, "N", "=")
    End If
    
    If fValuePresent(item_type_cd) = True Then
        strSql = strSql & " and item_type_cd = '" & item_type_cd & "' "
    End If
    If fValuePresent(sales_ord_id) = True Then
        strSql = strSql & " and sales_ord_id = " & CStr(sales_ord_id) & " "
    End If
    If fValuePresent(assem_ord_id) Then
        strSql = strSql & " and assem_ord_id = " & CStr(assem_ord_id) & " "
    End If
    If fValuePresent(status_cd) Then
        strSql = strSql & " and status_cd = '" & status_cd & "' "
    Else
        strSql = strSql & " and status_cd in ('OP', 'BO', 'RD', 'CP') "
    End If
    If fValuePresent(required_begin_dt) And Not IsNull(required_end_dt) Then
        strSql = strSql & " and " & fWhereVariable("required_dt", required_begin_dt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("required_dt", required_end_dt, "D", "<=")
    End If
    
    If fValuePresent(p_inv_loc_id) Then
        strSql = strSql & " and inv_loc_id = '" & p_inv_loc_id & "' "
    End If
    
    strSql = strSql & " ORDER BY a.item_id, a.required_dt, c.seq_nbr "
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If

    Set RecordSet = objRecordset
    Exit Sub
errHandler:
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

Public Sub addMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            , ByVal v_dblItemID As Double _
            , ByVal v_strInvLocId As String _
            , ByVal v_dblRequiredQty As Double _
            , ByVal v_dteRequiredDt As Date _
            , ByVal v_dblRequiredQtyVarPct As Double _
            , ByVal v_varAllocPriorityNbr As Variant _
            , ByVal v_varSupplierId As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.openAssemblyMatReq"
       
    If Len(v_varSupplierId) < 1 Then
        v_varSupplierId = Me.supplier_id
    End If
    
    If IsNumeric(v_varAllocPriorityNbr) = False Then
        v_varAllocPriorityNbr = 0
    End If
    
    Call validate( _
              r_objError _
            , r_objTransData _
            , "add" _
            , v_strOrderTypeCd _
            , v_varSalesOrdId _
            , v_varSoLineNbr _
            , v_varAssemOrdId _
            , v_varAcLineNbr _
            , v_dblItemID _
            , v_strInvLocId _
            , v_dblRequiredQty _
            , v_dteRequiredDt _
            , v_dblRequiredQtyVarPct _
            , v_varAllocPriorityNbr _
            , v_varSupplierId _
            )
    
    Call IssueInsertData( _
              r_objError _
            , r_objTransData _
            , v_strOrderTypeCd _
            , v_varSalesOrdId _
            , v_varSoLineNbr _
            , v_varAssemOrdId _
            , v_varAcLineNbr _
            , v_dblItemID _
            , v_strInvLocId _
            , v_dblRequiredQty _
            , v_dteRequiredDt _
            , v_dblRequiredQtyVarPct _
            , v_varAllocPriorityNbr _
            , v_varSupplierId _
            )
    
    Exit Sub
errHandler:
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

Public Sub Change( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_dblMatReqId As Variant _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            , ByVal v_varRequiredQty As Variant _
            , ByVal v_varRequiredDt As Variant _
            , ByVal v_varRequiredQtyVarPct As Variant _
            , ByVal v_varAllocPriorityNbr As Variant _
            , ByVal v_varAllocatedFl As Variant _
            , ByVal v_varSupplierId As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.Change"
    
    Dim objInventoryAlloc As Object, _
        objMatReq As mat_req
    
    Set objMatReq = New mat_req
    Set objMatReq.o_error = r_objError
    Set objMatReq.o_trans_data = r_objTransData
    objMatReq.mat_req_id = v_dblMatReqId
    objMatReq.assem_ord_id = v_varAssemOrdId
    objMatReq.ac_line_nbr = v_varAcLineNbr
    objMatReq.sales_ord_id = v_varSalesOrdId
    objMatReq.so_line_nbr = v_varSoLineNbr
    Call objMatReq.GetMatReq
    
                    'if the mat requirement does not exist
                    'don't worry about trying to update it.
    If Len(objMatReq.mat_req_id) < 1 Then
        Exit Sub
    End If
        
    Me.mat_req_id = objMatReq.mat_req_id
    
    If Len(v_varSupplierId) < 1 Then
        v_varSupplierId = Me.supplier_id
    End If
    
    If IsNumeric(v_varAllocPriorityNbr) = False Then
        v_varAllocPriorityNbr = objMatReq.alloc_priority_nbr
    End If
    
    Call validate( _
              r_objError _
            , r_objTransData _
            , "change" _
            , v_strOrderTypeCd _
            , v_varSalesOrdId _
            , v_varSoLineNbr _
            , v_varAssemOrdId _
            , v_varAcLineNbr _
            , Empty _
            , Empty _
            , v_varRequiredQty _
            , v_varRequiredDt _
            , v_varRequiredQtyVarPct _
            , v_varAllocPriorityNbr _
            , v_varSupplierId _
            )
    
    Call IssueUpdateData( _
              r_objError _
            , r_objTransData _
            , objMatReq.mat_req_id _
            , Empty _
            , v_varRequiredQty _
            , v_varAllocPriorityNbr _
            , v_varAllocatedFl _
            , v_varRequiredDt _
            , v_varRequiredQtyVarPct _
            )
    
                    'if the material requirement has been
                    'allocated, and the required quantity
                    'has changed, re allocate the material
                    'requirement
    If objMatReq.allocated_fl = "Y" And _
       (Len(v_varRequiredQty) > 0 Or Len(v_varRequiredQtyVarPct) > 0) Then
        If v_varRequiredQty <> objMatReq.required_qty Then
            Call objMatReq.GetMatReq
                'note, allocate function will call mat_req.updatequantities
            Call allocateMatReq( _
                        r_objError _
                      , r_objTransData _
                      , objMatReq.mat_req_id _
                      , objMatReq.item_id _
                      , objMatReq.inv_loc_id _
                      )
        Else
            If v_varRequiredQtyVarPct <> objMatReq.required_qty_var_pct Then
                Call SetStatus( _
                        r_objError _
                      , r_objTransData _
                      , Me.mat_req_id _
                      )
            End If
        End If
        
    End If

    Exit Sub
errHandler:
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


Public Sub openMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_varMatReqId As Variant _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.openAssemblyMatReq"
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.mat_req_id = v_varMatReqId
    Me.assem_ord_id = v_varAssemOrdId
    Me.ac_line_nbr = v_varAcLineNbr
    Me.sales_ord_id = v_varSalesOrdId
    Me.so_line_nbr = v_varSoLineNbr
    Call Me.GetMatReq
    
    If Me.status_cd = "CA" Then
        With r_objError
            .type_cd = "E"
            .err_cd = "1300"
            .routine_nm = strRoutine_nm
            .err_desc = "Material Requirement has been canceled, it cant be re-opened"
            .message_id = 0
        End With
        Exit Sub
    End If
                                        'call update, but make sure the status code
                                        'is the only thing being updated.
    Call IssueUpdateData( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                , "OP" _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
                
    
    If Me.allocated_fl = "Y" Then
                    'Allocating material requirement will
                    'call updatequantity for the material requirement
        Call allocateMatReq( _
                    r_objError _
                  , r_objTransData _
                  , Me.mat_req_id _
                  , Me.item_id _
                  , Me.inv_loc_id _
                  )
    Else
        Call Me.UpdateQuantities( _
                      r_objError _
                    , r_objTransData _
                    , Me.mat_req_id _
                    )
    End If
    
    Exit Sub
errHandler:
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

            'this method is so that the user can close a material requirement from the
            'screen
Public Sub closeMatReqAsp()
    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.closeMatReqAsp"
    
    Call Me.closeMatReq(Me.o_error, Me.o_trans_data, Me.mat_req_id, "MR", Null, Null, Null, Null)
    
    Exit Sub
errHandler:
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


            'this method is so that the user can open a material requirement from the
            'screen
Public Sub openMatReqAsp()
    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.closeMatReqAsp"
    
    Call Me.openMatReq(Me.o_error, Me.o_trans_data, Me.mat_req_id, "MR", Null, Null, Null, Null)
    
    Exit Sub
errHandler:
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




Public Sub closeMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_varMatReqId As Variant _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.closeMatReq"
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.mat_req_id = v_varMatReqId
    Me.assem_ord_id = v_varAssemOrdId
    Me.ac_line_nbr = v_varAcLineNbr
    Me.sales_ord_id = v_varSalesOrdId
    Me.so_line_nbr = v_varSoLineNbr
    Call Me.GetMatReq
    
    If Me.status_cd = "CA" Then
        With r_objError
            .type_cd = "E"
            .err_cd = "1300"
            .routine_nm = strRoutine_nm
            .err_desc = "Material Requirement has been canceled, it cant be re-opened"
            .message_id = 0
        End With
        Exit Sub
    End If
                    'Remove this edit,  we do not want to prevent
                    'a assembly order from going to complete status
                    'becuase inventory has not been relieved on the
                    'assembly components
'    If Me.consumed_qty < 1 Then
'        With r_objError
'            .type_cd = "E"
'            .err_cd = "1300"
'            .routine_nm = strRoutine_nm
'            .err_desc = "Material Requirement has not consumed any inventory, it can't be closed, only cancelled."
'            .message_id = 0
'        End With
'        Exit Sub
'    End If
                                        'call update, but make sure the status code
                                        'is the only thing being updated.
    Call IssueUpdateData( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                , "CL" _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
    
    Call deallocateMatReq( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                )
                    'DeallocateMatReq calls update quantities.
                    'No need to call it again
    'Call Me.UpdateQuantities( _
    '              r_objError _
    '            , r_objTransData _
    '            , Me.mat_req_id _
    '            )
    
    
    Exit Sub
errHandler:
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


Public Sub cancelMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_varMatReqId As Variant _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.cancelAssemblyMatReq"
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.mat_req_id = v_varMatReqId
    Me.assem_ord_id = v_varAssemOrdId
    Me.ac_line_nbr = v_varAcLineNbr
    Me.sales_ord_id = v_varSalesOrdId
    Me.so_line_nbr = v_varSoLineNbr
    Call Me.GetMatReq
    
    If Me.consumed_qty > 0 Then
        With r_objError
            .type_cd = "E"
            .err_cd = "1300"
            .routine_nm = strRoutine_nm
            .err_desc = "Material Requirement has consumed inventory, it can't be canceled, only closed."
            .message_id = 0
        End With
        Exit Sub
    End If
                                        
                                        'call update, but make sure the status code
                                        'is the only thing being updated.
    Call IssueUpdateData( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                , "CA" _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
    
    Call deallocateMatReq( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                )
                    
                    'DeallocateMatReq calls update quantities.
                    'No need to call it again
    'Call Me.UpdateQuantities( _
    '              r_objError _
    '            , r_objTransData _
    '            , Me.mat_req_id _
    '            )
    
    
    Exit Sub
errHandler:
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


Public Sub holdMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_varMatReqId As Variant _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.cancelAssemblyMatReq"
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.mat_req_id = v_varMatReqId
    Me.assem_ord_id = v_varAssemOrdId
    Me.ac_line_nbr = v_varAcLineNbr
    Me.sales_ord_id = v_varSalesOrdId
    Me.so_line_nbr = v_varSoLineNbr
    Call Me.GetMatReq
    
    If Me.status_cd = "CA" Or Me.status_cd = "CL" Then
        With r_objError
            .type_cd = "E"
            .err_cd = "1300"
            .routine_nm = strRoutine_nm
            .err_desc = "Material Requirement is inactive and can't be put on hold"
            .message_id = 0
        End With
        Exit Sub
    End If
    
                                        'call update, but make sure the status code
                                        'is the only thing being updated.
    Call IssueUpdateData( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                , "OH" _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
    
    Call allocateMatReq( _
                r_objError _
              , r_objTransData _
              , Me.mat_req_id _
              , Me.item_id _
              , Me.inv_loc_id _
              )
                    
                    'allocateMatReq calls update quantities.
                    'No need to call it again
    'Call Me.UpdateQuantities( _
    '              r_objError _
    '            , r_objTransData _
    '            , Me.mat_req_id _
    '            )
    
    Exit Sub
errHandler:
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



Public Sub generatePlan()
    
    On Error GoTo errHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.GeneratePlan"
    
    Dim strSql As String, _
        varSupplier_id As Variant, _
        objItem As Object, _
        dblUnallocatedQty As Double
    
    varSupplier_id = p_supplier_id
    
    Call Me.GetMatReq
    
    Set objItem = CreateObject("converge_item.item")
    Set objItem.o_error = o_error
    Set objItem.o_trans_data = o_trans_data
    objItem.p_item_id = p_item_id
    objItem.getItem
    
    If Len(varSupplier_id) > 0 Then
        If varSupplier_id <> p_supplier_id Then
            p_supplier_id = varSupplier_id
            ModifySupplier
        End If
    End If
        
    dblUnallocatedQty = allocateMatReq( _
                Me.o_error _
              , Me.o_trans_data _
              , Me.mat_req_id _
              , Me.item_id _
              , Me.inv_loc_id _
              )
    
    If dblUnallocatedQty > 0 Then
        If Me.item_type_cd = "PU" Then
            Dim objPurchaseOrder As Object
            Set objPurchaseOrder = CreateObject("converge_purchase.purchase_order")
            Set objPurchaseOrder.o_error = o_error
            Set objPurchaseOrder.o_trans_data = o_trans_data
            objPurchaseOrder.p_po_id = Null
            objPurchaseOrder.p_type_cd = "PO"
            
            If p_supplier_id > 0 Then
                objPurchaseOrder.p_supplier_id = p_supplier_id
            Else
                objPurchaseOrder.p_supplier_id = Null
            End If
            
            objPurchaseOrder.p_inv_loc_id = Me.inv_loc_id
            objPurchaseOrder.p_drop_ship_fl = "N"
            objPurchaseOrder.p_sales_tax_pct = 0
            objPurchaseOrder.p_confirmation_nbr = Empty
            ' Pass it in
            'objPurchaseOrder.p_drop_ship_sales_ord_id = p_sales_ord_id
            objPurchaseOrder.p_entered_dt = Now
            objPurchaseOrder.p_buyer_assoc_id = o_trans_data.p_assoc_id
            objPurchaseOrder.p_buyer_assoc_nbr = Empty
            objPurchaseOrder.p_backorder_cd = Empty
            objPurchaseOrder.p_carrier_nm = Empty
            objPurchaseOrder.p_currency_cd = Empty
            objPurchaseOrder.p_fob_terms_cd = Empty
            objPurchaseOrder.p_freight_cd = Empty
            objPurchaseOrder.p_comm_meth_cd = Empty
            objPurchaseOrder.p_payment_terms_cd = Empty
            objPurchaseOrder.p_transportation_cd = Empty
            Call objPurchaseOrder.AddPOfromRequirement(Me.item_id, dblUnallocatedQty, DateAdd("d", -(objItem.p_lead_days), Me.required_dt))
        Else
            Dim objAssemblyOrder As Object
            Set objAssemblyOrder = CreateObject("converge_prod.assembly_ord")
            Set objAssemblyOrder.o_error = o_error
            Set objAssemblyOrder.o_trans_data = o_trans_data
            objAssemblyOrder.ItemID = Me.item_id
            objAssemblyOrder.ItemNbr = Null
            objAssemblyOrder.LocID = Me.inv_loc_id
            objAssemblyOrder.DueDt = DateAdd("d", -(objItem.p_lead_days), Me.required_dt)
            objAssemblyOrder.OrderedQty = dblUnallocatedQty
            objAssemblyOrder.AddAssemblyOrder
            objAssemblyOrder.openOrder
        End If
        Call allocateMatReq( _
                    Me.o_error _
                  , Me.o_trans_data _
                  , Me.mat_req_id _
                  , Me.item_id _
                  , Me.inv_loc_id _
                  )
    End If
        
    Exit Sub
errHandler:
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


Public Sub ModifySupplier()

    On Error GoTo errHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        intUpdateQty        As Double
    
    strRoutine_nm = "mat_req.cls.ModifySupplier"
    
    strSql = "UPDATE mat_req SET " & _
            " " & fUpdateVariable("supplier_id", supplier_id, "N", "=") & _
            ", " & fUpdateVariable("alloc_priority_nbr", alloc_priority_nbr, "N", "=") & _
            " where " & fWhereVariable("mat_req_id", mat_req_id, "N", "=")
            
    Call o_trans_data.ExecuteDatabase(strSql, intUpdateQty, o_error)

    Exit Sub
errHandler:
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


Private Sub validate( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_strFunctionCd As String _
            , ByVal v_strOrderTypeCd As String _
            , ByRef r_varSalesOrdId As Variant _
            , ByRef r_varSoLineNbr As Variant _
            , ByRef r_varAssemOrdId As Variant _
            , ByRef r_varAcLineNbr As Variant _
            , ByVal v_varItemId As Variant _
            , ByVal v_varInvLocId As Variant _
            , ByVal v_varRequiredQty As Variant _
            , ByVal v_varRequiredDt As Variant _
            , ByVal v_varRequiredQtyVarPct As Variant _
            , ByVal v_varAllocPriorityNbr As Variant _
            , ByVal v_varSupplierId As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.validate"

                'so means this is being called by a sales order
                'material requirement.  ao means assmbly.  MR means
                'we don't know, but we have provided the mat_req_id
    Select Case v_strOrderTypeCd
        Case "SO"
            r_varAssemOrdId = Null
            r_varAcLineNbr = Null
        Case "AO"
            r_varSalesOrdId = Null
            r_varSoLineNbr = Null
        Case "MR"
            r_varSalesOrdId = Null
            r_varSoLineNbr = Null
            r_varAssemOrdId = Null
            r_varAcLineNbr = Null
        Case Else
            With r_objError
                .type_cd = "E"
                .err_cd = "1300"
                .routine_nm = strRoutine_nm
                .err_desc = "Order type code is invalid"
                .message_id = 0
            End With
            Exit Sub
    End Select
    
    If v_strFunctionCd = "add" Then
    
        If IsNull(fGetItemNbr(v_varItemId, r_objError, r_objTransData)) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0100"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
        
        If IsNull(fGetInvLocNm(v_varInvLocId, r_objError, r_objTransData)) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0110"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
        
        If IsNumeric(v_varRequiredQty) = False Or v_varRequiredQty < 0 Then
            o_error.type_cd = "E"
            o_error.err_cd = "0400"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
        
        If IsDate(v_varRequiredDt) = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "0500"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
        
        If IsNumeric(v_varRequiredQtyVarPct) = False Or v_varRequiredQtyVarPct < 0 Then
            o_error.type_cd = "E"
            o_error.err_cd = "0600"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
    
    End If
    
    If v_strFunctionCd = "change" Then
    
        If IsEmpty(v_varRequiredQty) = False Then
            If IsNumeric(v_varRequiredQty) = False Or v_varRequiredQty < 0 Then
                o_error.type_cd = "E"
                o_error.err_cd = "1100"
                o_error.routine_nm = strRoutine_nm
                o_error.message_id = 0
            End If
        End If
        
        If IsEmpty(v_varRequiredDt) = False Then
            If IsDate(v_varRequiredDt) = False Then
                o_error.type_cd = "E"
                o_error.err_cd = "1200"
                o_error.routine_nm = strRoutine_nm
                o_error.message_id = 0
            End If
        End If
        
        If IsEmpty(v_varRequiredQtyVarPct) = False Then
            If IsNumeric(v_varRequiredQtyVarPct) = False Or v_varRequiredQtyVarPct < 0 Then
                o_error.type_cd = "E"
                o_error.err_cd = "1300"
                o_error.routine_nm = strRoutine_nm
                o_error.message_id = 0
            End If
        End If
    End If
    
    Exit Sub
errHandler:
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

Public Sub UpdateQuantities( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_dblMatReqId As Double _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.UpdateQuantities"
    
    Dim dblUpdatedqty As Double
    Dim strSql As String
    
    strSql = "update mat_req set " & _
                " " & fUpdateVariable("ready_qty", getReadyQty(r_objError, r_objTransData, v_dblMatReqId), "N", "=") & _
                "," & fUpdateVariable("back_ordered_qty", getBack_OrderedQty(r_objError, r_objTransData, v_dblMatReqId), "N", "=") & _
                "," & fUpdateVariable("consumed_qty", getConsumedQty(r_objError, r_objTransData, v_dblMatReqId), "S", "=") & _
                " where" & _
                " " & fWhereVariable("mat_req_id", v_dblMatReqId, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
    
    Call SetStatus( _
              r_objError _
            , r_objTransData _
            , v_dblMatReqId _
            )
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub SetStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef r_dblMatReqId As Double _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.SetStatus"

    Dim strSql As String, _
        dblRequired_qty As Double, _
        strStatusCd As String, _
        strOldStatusCd As String, _
        objAssemblyOrder As assembly_ord, _
        objInventoryAlloc As Object
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.mat_req_id = r_dblMatReqId
    Me.GetMatReq
    'Debug.Print Me.mat_req_id & " " & Me.item_nbr & " " & Me.status_cd
    If Me.status_cd = "CA" Or _
       Me.status_cd = "OH" Or _
       Me.status_cd = "CL" Then
        Exit Sub
    End If
    
    dblRequired_qty = Me.required_qty
    If IsNumeric(Me.required_qty_var_pct) = True Then
        If Me.required_qty_var_pct > 0 Then
            dblRequired_qty = Me.required_qty * (1 - Me.required_qty_var_pct)
        End If
    End If
        
    If dblRequired_qty > Me.back_ordered_qty + Me.ready_qty + Me.consumed_qty Then
        strStatusCd = "OP"
    End If
    
    If dblRequired_qty <= Me.back_ordered_qty + Me.ready_qty + Me.consumed_qty Then
        strStatusCd = "BO"
    End If
    
    If dblRequired_qty <= Me.ready_qty + Me.consumed_qty Then
        strStatusCd = "RD"
    End If
    
    If dblRequired_qty <= Me.consumed_qty Then
        strStatusCd = "CP"
    End If
    
    Call IssueUpdateData( _
                  r_objError _
                , r_objTransData _
                , Me.mat_req_id _
                , strStatusCd _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
    
    If IsNull(Me.assem_ord_id) = False Then
        Set objAssemblyOrder = New assembly_ord
        Call objAssemblyOrder.SetStatus( _
                      r_objError _
                    , r_objTransData _
                    , Me.assem_ord_id _
                    )
    End If
    
    Exit Sub
errHandler:
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

Private Sub IssueInsertData( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_strOrderTypeCd As String _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varSoLineNbr As Variant _
            , ByVal v_varAssemOrdId As Variant _
            , ByVal v_varAcLineNbr As Variant _
            , ByVal v_dblItemID As Double _
            , ByVal v_strInvLocId As String _
            , ByVal v_dblRequiredQty As Double _
            , ByVal v_dteRequiredDt As Date _
            , ByVal v_dblRequiredQtyVarPct As Double _
            , ByVal v_dblAllocPriorityNbr As Double _
            , ByVal v_varSupplierId As Variant _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.IssueInsertData"
    
    Dim dblUpdatedqty As Double, _
        strSql As String, _
        dblMatReqId As Double, _
        varItemTypeCd As Variant, _
        varUomCd As Variant
        
    dblMatReqId = fGetId("mat_req_id", r_objError, r_objTransData)
    Call getItemData(r_objError, r_objTransData, v_dblItemID, varUomCd, varItemTypeCd)
    Me.mat_req_id = dblMatReqId
    
    strSql = "insert into mat_req( " & _
                 "  mat_req_id" & _
                 ", allocated_fl " & _
                 ", alloc_priority_nbr " & _
                 ", item_id" & _
                 ", item_type_cd" & _
                 ", inv_loc_id" & _
                 ", required_qty" & _
                 ", required_dt" & _
                 ", status_cd" & _
                 ", back_ordered_qty" & _
                 ", ready_qty" & _
                 ", consumed_qty" & _
                 ", sales_ord_id" & _
                 ", so_line_nbr" & _
                 ", assem_ord_id" & _
                 ", ac_line_nbr" & _
                 ", required_qty_var_pct" & _
                 ", uom_cd" & _
                 ", supplier_id" & _
                 ", updated_assoc_id" & _
                 ", updated_ts" & _
                 ", entered_assoc_id" & _
                 ", entered_ts" & _
                 " ) "
    strSql = strSql & "values (" & _
                 " " & fInsertVariable(dblMatReqId, "N") & _
                 "," & fInsertVariable("N", "S") & _
                 "," & fInsertVariable(v_dblAllocPriorityNbr, "N") & _
                 "," & fInsertVariable(v_dblItemID, "N") & _
                 "," & fInsertVariable(varItemTypeCd, "S") & _
                 "," & fInsertVariable(v_strInvLocId, "S") & _
                 "," & fInsertVariable(v_dblRequiredQty, "N") & _
                 "," & fInsertVariable(v_dteRequiredDt, "D") & _
                 "," & fInsertVariable("OP", "S") & _
                 "," & fInsertVariable(0, "N") & _
                 "," & fInsertVariable(0, "N") & _
                 "," & fInsertVariable(0, "N") & _
                 "," & fInsertVariable(v_varSalesOrdId, "N") & _
                 "," & fInsertVariable(v_varSoLineNbr, "N") & _
                 "," & fInsertVariable(v_varAssemOrdId, "N") & _
                 "," & fInsertVariable(v_varAcLineNbr, "N") & _
                 "," & fInsertVariable(v_dblRequiredQtyVarPct, "N") & _
                 "," & fInsertVariable(varUomCd, "S") & _
                 "," & fInsertVariable(v_varSupplierId, "N") & _
                 "," & fInsertVariable(r_objTransData.p_assoc_id, "N") & _
                 "," & fInsertVariable("current_ts", "D") & _
                 "," & fInsertVariable(r_objTransData.p_assoc_id, "N") & _
                 "," & fInsertVariable("current_ts", "D") & _
                 ")"
  
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub IssueUpdateData( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            , ByVal v_varStatusCd As Variant _
            , ByVal v_varRequiredQty As Variant _
            , ByVal v_varAllocPriorityNbr As Variant _
            , ByVal v_varAllocatedFl As Variant _
            , ByVal v_varRequiredDt As Variant _
            , ByVal v_varRequiredQtyVarPct As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.IssueUpdateData"
    
    Dim dblUpdatedqty As Double, _
        strSql As String
        
    
    strSql = "update mat_req set  " & _
             "  " & fUpdateVariable("status_cd", v_varStatusCd, "S", "=") & _
             ", " & fUpdateVariable("required_qty", v_varRequiredQty, "N", "=") & _
             ", " & fUpdateVariable("required_qty_var_pct", v_varRequiredQtyVarPct, "N", "=") & _
             ", " & fUpdateVariable("alloc_priority_nbr", v_varAllocPriorityNbr, "N", "=") & _
             ", " & fUpdateVariable("allocated_fl", v_varAllocatedFl, "S", "=") & _
             ", " & fUpdateVariable("required_dt", v_varRequiredDt, "D", "=") & _
             " where " & _
             "   mat_req_id = " & v_dblMatReqId

    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
    
    Exit Sub
errHandler:
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


Private Function getItemData( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblItemID As Double _
            , ByRef r_varUomCd As Variant _
            , ByRef r_varTypeCd As Variant _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim dblUpdateQty As Double
    Dim strRoutine_nm As String, _
        objRecordset As ADODB.RecordSet, _
        strSql As String
    
    strRoutine_nm = "mat_req.cls.fGetItemTypeCd"
    
    strSql = "SELECT * FROM item " & _
               " WHERE " & fWhereVariable("item_id", v_dblItemID, "N", "=")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        r_varUomCd = Null
        r_varTypeCd = Null
        Exit Function
    End If
    
    r_varTypeCd = objRecordset("type_cd")
    r_varUomCd = objRecordset("uom_cd")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


Private Function getReadyQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            ) As Double

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String, _
        strSql As String, _
        objRecordset As ADODB.RecordSet
    
    strRoutine_nm = "mat_req.cls.getReadyQty"
    
    Set objRecordset = New ADODB.RecordSet

    strSql = "SELECT alloc_qty  " & _
             "FROM inventory_alloc " & _
             "WHERE         " & _
             "  " & fWhereVariable("mat_req_id", v_dblMatReqId, "N", "=")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objRecordset(0)) = True Or objRecordset.EOF = True Then
        getReadyQty = 0
    Else
        getReadyQty = objRecordset(0)
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


Private Function getBack_OrderedQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            ) As Double

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String, _
        objRecordset As ADODB.RecordSet, _
        strSql As String
    
    strRoutine_nm = "mat_req.cls.getBack_OrderedQty"
    Set objRecordset = New ADODB.RecordSet

    strSql = "SELECT alloc_qty  " & _
            "FROM on_ord_inv_alloc " & _
            "WHERE         " & _
            "  " & fWhereVariable("mat_req_id", v_dblMatReqId, "N", "=")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objRecordset(0)) = True Or objRecordset.EOF = True Then
        getBack_OrderedQty = 0
    Else
        getBack_OrderedQty = objRecordset(0)
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
   
    Exit Function
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function

Private Function getConsumedQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            ) As Double

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String, _
        strSql As String, _
        objRecordset As Object
    
    strRoutine_nm = "mat_req.cls.getShippedQty"
    Set objRecordset = New ADODB.RecordSet

    strSql = "SELECT sum(adj_qty)  "
    strSql = strSql & "FROM inventory_adj "
    strSql = strSql & "WHERE         "
    strSql = strSql & "  " & fWhereVariable("mat_req_id", v_dblMatReqId, "N", "=")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objRecordset(0)) = True Or objRecordset.EOF = True Then
        getConsumedQty = 0
    Else
        getConsumedQty = objRecordset(0) * -1
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


Private Function allocateMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            , ByVal v_dblItemID As Double _
            , ByVal v_strInvLocId As String _
            ) As Double
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.allocateMatReq"
    
    Dim objInventoryAlloc As Object
    Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
    With objInventoryAlloc
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        .p_mat_req_id = v_dblMatReqId
        .p_item_id = v_dblItemID
        .p_inv_loc_id = v_strInvLocId
        .Allocate
        allocateMatReq = .p_unallocated_qty
    End With
    
    Exit Function
errHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function

Private Sub deallocateMatReq( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblMatReqId As Double _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "mat_req.cls.deallocateMatReq"
    
    Dim objInventoryAlloc As Object
    Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
    Set objInventoryAlloc.o_error = r_objError
    Set objInventoryAlloc.o_trans_data = r_objTransData
    objInventoryAlloc.p_mat_req_id = mat_req_id
    objInventoryAlloc.deallocate

    Exit Sub
errHandler:
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



