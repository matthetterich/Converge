VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "purchase_order_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_po_id As Variant
Private mvarp_line_nbr As Variant
Private mvarp_drop_ship_so_line_nbr As Variant
Private mvarp_item_id As Variant
Private mvarp_item_nbr As Variant
Private mvarp_received_qty As Variant
Private mvarp_rejected_qty As Variant
Private mvarp_line_amnt As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_supplier_item_nbr As Variant
Private mvarp_item_cost As Variant
Private mvarp_uom_cd As Variant
Private mvarp_uom_desc As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_purchase_order As purchase_order
Private mvaro_old_purchase_order_item As purchase_order_item
Private mvarp_invoiced_qty As Variant
Private mvarp_ordered_qty As Variant
Private mvarp_type_desc As Variant
Private mvarp_type_cd As Variant
Private mvarp_inspection_req_fl As Variant
Private mvarp_bypass_del_edit_fl As Variant
Private mvarp_item_text As Variant
Private mvaro_record_set As ADODB.Recordset
Private mvaro_item As Object
Private mvarp_item_desc As Variant
Private mvarp_requested_del_dt As Variant
Private mvarp_supplier_nbr As Variant
Private mvarp_start_dt As Variant
Private mvarp_end_dt As Variant
Private mvarp_overrideMinQtyFl As Variant
Private mvarp_recCompleteVarPct As Variant

Public Property Let p_recCompleteVarPct(ByVal vData As Variant)
    mvarp_recCompleteVarPct = vData
End Property
Public Property Get p_recCompleteVarPct() As Variant
        p_recCompleteVarPct = mvarp_recCompleteVarPct
End Property


Public Property Let p_overrideMinQtyFl(ByVal vData As Variant)
    mvarp_overrideMinQtyFl = vData
End Property
Public Property Get p_overrideMinQtyFl() As Variant
        p_overrideMinQtyFl = mvarp_overrideMinQtyFl
End Property

Public Property Let p_start_dt(ByVal vData As Variant)
    mvarp_start_dt = vData
End Property
Public Property Get p_start_dt() As Variant
        p_start_dt = mvarp_start_dt
End Property

Public Property Let p_end_dt(ByVal vData As Variant)
    mvarp_end_dt = vData
End Property
Public Property Get p_end_dt() As Variant
        p_end_dt = mvarp_end_dt
End Property

Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property

Public Property Let p_requested_del_dt(ByVal vData As Variant)
    mvarp_requested_del_dt = vData
End Property
Public Property Get p_requested_del_dt() As Variant
        p_requested_del_dt = mvarp_requested_del_dt
End Property

Public Property Let p_bypass_del_edit_fl(ByVal vData As Variant)
    mvarp_bypass_del_edit_fl = vData
End Property
Public Property Get p_bypass_del_edit_fl() As Variant
        p_bypass_del_edit_fl = mvarp_bypass_del_edit_fl
End Property

Public Property Let p_item_desc(ByVal vData As Variant)
    mvarp_item_desc = vData
End Property
Public Property Get p_item_desc() As Variant
        p_item_desc = mvarp_item_desc
End Property

Public Property Let p_item_text(ByVal vData As Variant)
    mvarp_item_text = vData
End Property
Public Property Get p_item_text() As Variant
    p_item_text = mvarp_item_text
End Property

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

Public Property Let p_inspection_req_fl(ByVal vData As Variant)
    mvarp_inspection_req_fl = vData
End Property
Public Property Get p_inspection_req_fl() As Variant
        p_inspection_req_fl = mvarp_inspection_req_fl
End Property

Public Property Let p_type_cd(ByVal vData As Variant)
    mvarp_type_cd = vData
End Property
Public Property Get p_type_cd() As Variant
        p_type_cd = mvarp_type_cd
End Property

Public Property Let p_type_desc(ByVal vData As Variant)
    mvarp_type_desc = vData
End Property
Public Property Get p_type_desc() As Variant
        p_type_desc = mvarp_type_desc
End Property

Public Property Let p_ordered_qty(ByVal vData As Variant)
    mvarp_ordered_qty = vData
End Property
Public Property Get p_ordered_qty() As Variant
        p_ordered_qty = mvarp_ordered_qty
End Property

Public Property Let p_invoiced_qty(ByVal vData As Variant)
    mvarp_invoiced_qty = vData
End Property
Public Property Get p_invoiced_qty() As Variant
        p_invoiced_qty = mvarp_invoiced_qty
End Property
Public Property Set o_purchase_order(ByVal vData As purchase_order)
    Set mvaro_purchase_order = vData
End Property
Public Property Get o_purchase_order() As purchase_order
    Set o_purchase_order = mvaro_purchase_order
End Property


Public Property Set o_old_purchase_order_item(ByVal vData As purchase_order_item)
    Set mvaro_old_purchase_order_item = vData
End Property
Public Property Get o_old_purchase_order_item() As purchase_order_item
    Set o_old_purchase_order_item = mvaro_old_purchase_order_item
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

Public Property Let p_uom_desc(ByVal vData As Variant)
    mvarp_uom_desc = vData
End Property
Public Property Get p_uom_desc() As Variant
        p_uom_desc = mvarp_uom_desc
End Property

Public Property Let p_uom_cd(ByVal vData As Variant)
    mvarp_uom_cd = vData
End Property
Public Property Get p_uom_cd() As Variant
        p_uom_cd = mvarp_uom_cd
End Property

Public Property Let p_item_cost(ByVal vData As Variant)
    mvarp_item_cost = vData
End Property
Public Property Get p_item_cost() As Variant
        p_item_cost = mvarp_item_cost
End Property

Public Property Let p_supplier_item_nbr(ByVal vData As Variant)
    mvarp_supplier_item_nbr = vData
End Property
Public Property Get p_supplier_item_nbr() As Variant
        p_supplier_item_nbr = mvarp_supplier_item_nbr
End Property

Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property

Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    mvarp_payment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
        p_payment_terms_desc = mvarp_payment_terms_desc
End Property
Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    mvarp_payment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
        p_payment_terms_cd = mvarp_payment_terms_cd
End Property
Public Property Let p_line_amnt(ByVal vData As Variant)
    mvarp_line_amnt = vData
End Property
Public Property Get p_line_amnt() As Variant
        p_line_amnt = mvarp_line_amnt
End Property
Public Property Let p_rejected_qty(ByVal vData As Variant)
    mvarp_rejected_qty = vData
End Property
Public Property Get p_rejected_qty() As Variant
        p_rejected_qty = mvarp_rejected_qty
End Property
Public Property Let p_received_qty(ByVal vData As Variant)
    mvarp_received_qty = vData
End Property
Public Property Get p_received_qty() As Variant
        p_received_qty = mvarp_received_qty
End Property
Public Property Let p_item_nbr(ByVal vData As Variant)
    mvarp_item_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
        p_item_nbr = mvarp_item_nbr
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property
Public Property Let p_line_nbr(ByVal vData As Variant)
    mvarp_line_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
        p_line_nbr = mvarp_line_nbr
End Property
Public Property Let p_drop_ship_so_line_nbr(ByVal vData As Variant)
    mvarp_drop_ship_so_line_nbr = vData
End Property
Public Property Get p_drop_ship_so_line_nbr() As Variant
        p_drop_ship_so_line_nbr = mvarp_drop_ship_so_line_nbr
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
        p_po_id = mvarp_po_id
End Property
Public Sub GetPurchaseOrderItem()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.GetPurchaseOrderItem"
  
    Call SelectPurchaseOrderItem

    Exit Sub
errHandler:
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
Public Sub Add()
           
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.Add"
    
    Call getPurchaseOrder
    Call Validate("add")
    Call UpdateOrderedQty
    Call IssueInsertData
    Call IssueUpdateLastCost
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
    Else
        Call changeOnOrderInventory( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_po_id _
                    , Me.p_line_nbr _
                    , Me.p_ordered_qty _
                    , 0 _
                    , Me.o_purchase_order.p_drop_ship_fl _
                    , Me.p_status_cd _
                    , Me.o_purchase_order.p_inv_loc_id _
                    , Me.p_item_id _
                    )
    End If
    Exit Sub
errHandler:
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

Public Sub Change()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.Change"
    
    Call getPurchaseOrder
    Call getOldPurchaseOrderItem
    Call Validate("change")
    Call UpdateOrderedQty
    Call IssueUpdateData
    
    If Len(Me.p_item_cost) > 0 And _
       Me.o_old_purchase_order_item.p_item_cost <> Me.p_item_cost Then
        Call updateAffectedLots(Me.o_error, Me.o_trans_data, Me.p_po_id, Me.p_line_nbr, Me.p_item_cost)
    End If
    
    Call UpdateQuantities( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_line_nbr _
                )
    
    Exit Sub
errHandler:
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

Public Sub Delete()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.Delete"
    
    Call getPurchaseOrder
    Call getOldPurchaseOrderItem
    Call Validate("delete")
    Call IssueDeleteData
    
    Call changeOnOrderInventory( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_line_nbr _
                , 0 _
                , 0 _
                , Me.o_purchase_order.p_drop_ship_fl _
                , "CA" _
                , Me.o_purchase_order.p_inv_loc_id _
                , Me.o_old_purchase_order_item.p_item_id _
                )
    
    Exit Sub
errHandler:
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

Public Sub closeOrderItem()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.closeorderitem"

    p_status_cd = "CL"
    Call ChangeStatus
    
    Exit Sub
errHandler:
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
'Description:   This function will change the item's status equal to the status
'               of the purchase order record.
'
'Arguments:     none
'//////////////////////////////////////////////////////////////////////////////
Public Sub ChangeStatus()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.changestatus"
    
    Dim objRecordSet As ADODB.Recordset, _
        strSql As String, _
        dblUpdatedQty As Double
    
    Call getPurchaseOrder
    Call getOldPurchaseOrderItem
    
    If Me.p_status_cd = "OP" Or _
       Me.p_status_cd = "CP" Or _
       Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "ST" Or _
       Me.p_status_cd = "OH" Or _
       Me.p_status_cd = "RS" Or _
       Me.p_status_cd = "NC" Or _
       Me.p_status_cd = "CA" Then
    Else
        With Me.o_error
            .p_type_cd = "F"
            .p_err_cd = "0200"
            .p_routine_nm = strRoutine_nm
            .p_message_id = 2010
        End With
        Exit Sub
    End If
    
    If Me.p_status_cd = "OP" Then
        If IsNull(Me.o_old_purchase_order_item.p_requested_del_dt) = True Then
            With Me.o_error
                .p_type_cd = "F"
                .p_err_cd = "0200"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1026
            End With
            Exit Sub
        End If
    End If
    
    If Me.p_status_cd = "OP" Then
        Me.p_status_cd = Me.o_purchase_order.p_status_cd
    End If
    
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_line_nbr _
                , Me.p_status_cd _
                )

    Call SetStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_line_nbr _
                )
    
    Call Me.GetPurchaseOrderItem
    
    Call changeOnOrderInventory( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_line_nbr _
                , Me.p_ordered_qty _
                , Me.p_received_qty _
                , Me.o_purchase_order.p_drop_ship_fl _
                , Me.p_status_cd _
                , Me.o_purchase_order.p_inv_loc_id _
                , Me.o_old_purchase_order_item.p_item_id _
                )
    Exit Sub
errHandler:
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
'Description:   This procedure checks the sets the received qty for purchase
'               Order Items.
'
'Arguments:
'//////////////////////////////////////////////////////////////////////////////
Public Sub UpdateQuantities( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_dblPoId As Double _
            , ByRef v_dblLineNbr As Double _
            )
        
    On Error GoTo errHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.IssueUpdateQuantities"
    
    Dim strSql As String, _
        dblUpdatedQty As Double, _
        dblRecievedQty As Double
        
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_po_id = v_dblPoId
    Me.p_line_nbr = v_dblLineNbr
    Call Me.GetPurchaseOrderItem
    Call getPurchaseOrder
    
    dblRecievedQty = getReceivedQty( _
                              r_objError _
                            , r_objTransData _
                            , v_dblPoId _
                            , v_dblLineNbr _
                            , Me.o_purchase_order.p_drop_ship_fl _
                            , Me.o_purchase_order.p_drop_ship_sales_ord_id _
                            , Me.p_drop_ship_so_line_nbr _
                            )
    
    strSql = " UPDATE purchase_order_item set " & _
                " " & fUpdateVariable("received_qty", dblRecievedQty, "N", "=") & _
                "   WHERE " & fWhereVariable("po_id", v_dblPoId, "N", "=") & _
                "   AND " & fWhereVariable("line_nbr", v_dblLineNbr, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedQty, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Call SetStatus( _
              r_objError _
            , r_objTransData _
            , v_dblPoId _
            , v_dblLineNbr _
            )
    
    Call changeOnOrderInventory( _
              r_objError _
            , r_objTransData _
            , v_dblPoId _
            , v_dblLineNbr _
            , Me.p_ordered_qty _
            , dblRecievedQty _
            , Me.o_purchase_order.p_drop_ship_fl _
            , Me.p_status_cd _
            , Me.o_purchase_order.p_inv_loc_id _
            , Me.p_item_id _
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure checks the sets the status for purchase order
'               items.  The quantity's are check and the status is changed
'               accordingly in table(purchase_order_item).
'
'Arguments:     VarPo_id, varLine_nbr,
'//////////////////////////////////////////////////////////////////////////////
Private Sub SetStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblPoId As Double _
            , ByVal v_dblLineNbr As Double _
            )

    On Error GoTo errHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.SetStatus"
    
    Dim strStatusCd As String, _
        strSql As String, _
        dblUpdatedQty As Double, _
        objPurchaseOrder As purchase_order
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_po_id = v_dblPoId
    Me.p_line_nbr = v_dblLineNbr
    Me.GetPurchaseOrderItem
    
    If Me.p_status_cd = "PL" Or _
       Me.p_status_cd = "OH" Or _
       Me.p_status_cd = "OP" Or _
       Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "RS" Or _
       Me.p_status_cd = "NC" Or _
       Me.p_status_cd = "CA" Then
        Exit Sub
    End If
    
    strStatusCd = "ST"
    
    If Me.p_received_qty >= Me.p_ordered_qty Then
        strStatusCd = "CP"
    End If
    
                    'find out if this is a drop ship po.  If it is, get the drop ship
                    'sales order and find out what the ship complete variance pct is.
                    'if we are with-in variance, call the poi complete.
    If strStatusCd = "ST" Then
        Set objPurchaseOrder = New purchase_order
        Set objPurchaseOrder.o_error = r_objError
        Set objPurchaseOrder.o_trans_data = r_objTransData
        objPurchaseOrder.p_po_id = Me.p_po_id
        Call objPurchaseOrder.getPurchaseOrder
        If objPurchaseOrder.p_drop_ship_fl = "Y" Then
            Dim objSoi As Object
            Set objSoi = CreateObject("converge_sales.sales_order_item")
            Set objSoi.o_error = r_objError
            Set objSoi.o_trans_data = r_objTransData
            objSoi.p_sales_ord_id = objPurchaseOrder.p_drop_ship_sales_ord_id
            objSoi.p_line_nbr = Me.p_drop_ship_so_line_nbr
            Call objSoi.GetSalesOrderLineItem
            If IsNumeric(objSoi.p_ship_complete_var_pct) = True Then
                If objSoi.p_ship_complete_var_pct < 1 And objSoi.p_ship_complete_var_pct > 0 Then
                    If Me.p_received_qty >= (Me.p_ordered_qty * (1 - objSoi.p_ship_complete_var_pct)) Then
                       strStatusCd = "CP"
                    End If
                End If
            End If
        Else
            If (Me.p_received_qty > Me.p_ordered_qty * (1 - Me.p_recCompleteVarPct)) Then
                strStatusCd = "CP"
            End If
        End If
    End If
    
    If Me.p_status_cd = strStatusCd Then
        Exit Sub
    End If
    
    Call IssueUpdateStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblPoId _
                , v_dblLineNbr _
                , strStatusCd _
                )
    
    Me.p_status_cd = strStatusCd
    
    Set objPurchaseOrder = New purchase_order
    Call objPurchaseOrder.SetStatus( _
                              r_objError _
                            , r_objTransData _
                            , v_dblPoId _
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



Private Sub Validate(sFunction_cd)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.Validate"
      
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If IsEmpty(Me.p_recCompleteVarPct) = False Then
            If IsNumeric(Me.p_recCompleteVarPct) = False Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "0100"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_desc = "received complete variance percentage must be numeric"
                Me.o_error.p_message_id = 3429
                Exit Sub
            Else
                If Me.p_recCompleteVarPct > 1 Or Me.p_recCompleteVarPct < 0 Then
                    Me.o_error.p_type_cd = "E"
                    Me.o_error.p_err_cd = "0100"
                    Me.o_error.p_routine_nm = strRoutine_nm
                    Me.o_error.p_message_id = 3429
                    Me.o_error.p_desc = "received complete variance percentage must be between 0 and 1"
                    Exit Sub
                End If
            End If
        Else
                    'if the poi is being added and the variance pct was not provided,
                    'default the value of the variance pct
            If sFunction_cd = "add" Then
                    'drop shipments use the variance pct from the sales order
                If Me.o_purchase_order.p_drop_ship_fl = "Y" Then
                    Me.p_recCompleteVarPct = 0
                Else
                    Me.p_recCompleteVarPct = 0.1
                End If
            End If
        End If
    End If
    
    If Len(p_requested_del_dt) > 0 Then
        If IsDate(p_requested_del_dt) = False Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0100"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1027
        End If
    End If
    
    If sFunction_cd = "add" Then
        If Me.o_purchase_order.p_status_cd = "ST" Or _
           Me.o_purchase_order.p_status_cd = "CA" Or _
           Me.o_purchase_order.p_status_cd = "CL" Or _
           Me.o_purchase_order.p_status_cd = "NC" Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0110"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2146
            Exit Sub
        End If
        
        If fValuePresent(p_po_id) = False Then
            Me.o_error.p_type_cd = "F"
            Me.o_error.p_err_cd = "0100"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2136
            Exit Sub
        End If
    
        If p_item_id > 0 Then
        Else
            Me.o_error.p_type_cd = "F"
            Me.o_error.p_err_cd = "0100"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2090
            Exit Sub
        End If
        
        Set o_item = CreateObject("converge_item.item")
        Set o_item.o_error = Me.o_error
        Set o_item.o_trans_data = Me.o_trans_data
        o_item.p_item_id = p_item_id
        o_item.GetItem
        
        If IsNull(o_item.p_item_id) = True Then
            Me.o_error.p_type_cd = "F"
            Me.o_error.p_err_cd = "0090"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2090
            Exit Sub
        End If
        
        If o_item.p_type_cd = "PU" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0105"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2145
            Exit Sub
        End If
        p_uom_cd = o_item.p_uom_cd
        If fValuePresent(p_supplier_item_nbr) = False Then
            p_supplier_item_nbr = o_item.p_item_nbr
        End If
    End If
    
    If sFunction_cd = "change" Then
        If o_purchase_order.p_status_cd = "ST" Or _
           o_purchase_order.p_status_cd = "CA" Or _
           o_purchase_order.p_status_cd = "CP" Or _
           o_purchase_order.p_status_cd = "NC" Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0110"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2146
            Exit Sub
        End If
        
        If IsNull(p_supplier_item_nbr) = True Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0120"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2110
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "delete" Then
        If (o_purchase_order.p_status_cd = "PL") Or p_bypass_del_edit_fl = "Y" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0110"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2147
            Exit Sub
        End If
    End If
        
    Exit Sub
errHandler:
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

Private Sub SelectPurchaseOrderItem()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim lv_amount As Variant
    Dim objRecordSet As ADODB.Recordset
    
    strRoutine_nm = "purchase_order_item.cls.SelectPurchaseOrderItem"
    
    strSql = "SELECT * "
    strSql = strSql & "FROM purchase_order_item "
    strSql = strSql & " WHERE "
    strSql = strSql & " " & fWhereVariable("po_id", p_po_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objRecordSet, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    Else
        If objRecordSet.EOF = True Then
            objRecordSet.Close
            Set objRecordSet = Nothing
            p_line_nbr = Null
        Else
            p_po_id = objRecordSet("po_id")
            p_line_nbr = objRecordSet("line_nbr")
            p_drop_ship_so_line_nbr = objRecordSet("drop_ship_so_line_nbr")
            p_item_desc = objRecordSet("item_desc")
            p_type_cd = objRecordSet("type_cd")
            p_supplier_item_nbr = objRecordSet("supplier_item_nbr")
            p_item_id = objRecordSet("item_id")
            If objRecordSet("override_min_qty_fl") = "Y" Then
                Me.p_overrideMinQtyFl = True
            Else
                Me.p_overrideMinQtyFl = False
            End If
            p_recCompleteVarPct = objRecordSet("rec_complete_var_pct")
            p_ordered_qty = objRecordSet("ordered_qty")
            p_item_cost = objRecordSet("item_cost")
            p_item_cost = p_item_cost
            p_uom_cd = objRecordSet("uom_cd")
            p_status_cd = objRecordSet("status_cd")
            p_payment_terms_cd = objRecordSet("payment_terms_cd")
            p_inspection_req_fl = objRecordSet("inspection_req_fl")
            p_requested_del_dt = objRecordSet("requested_del_dt")
            p_received_qty = objRecordSet("received_qty")
            p_rejected_qty = objRecordSet("rejected_qty")
            p_invoiced_qty = objRecordSet("invoiced_qty")
            lv_amount = p_ordered_qty * p_item_cost
            p_line_amnt = lv_amount
            p_item_text = objRecordSet("item_tx")
            
            objRecordSet.Close
            Set objRecordSet = Nothing
            
            ' Function Calls
            p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, Me.o_error, Me.o_trans_data)
            p_status_desc = fGetDesc("purchase_order_item_status_cd", p_status_cd, Me.o_error, Me.o_trans_data)
            p_type_desc = fGetDesc("item_type_cd", p_type_cd, Me.o_error, Me.o_trans_data)
            p_uom_desc = fGetDesc("uom_cd", p_uom_cd, Me.o_error, Me.o_trans_data)
            p_item_nbr = fGetItemNbr(p_item_id, Me.o_error, Me.o_trans_data)
    
        End If
    End If
     
    Exit Sub
errHandler:
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

Private Sub IssueInsertData()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    
    strRoutine_nm = "purchase_order_item.cls.IssueInsertData"
    
    p_line_nbr = fGetLineNbr(p_po_id)
    p_received_qty = 0
    p_rejected_qty = 0
    p_invoiced_qty = 0
    
    strSql = "insert into purchase_order_item( "
    strSql = strSql & "  po_id"
    strSql = strSql & ", line_nbr"
    strSql = strSql & ", drop_ship_so_line_nbr"
    strSql = strSql & ", item_desc"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", supplier_item_nbr"
    strSql = strSql & ", item_id"
    strSql = strSql & ", override_min_qty_fl"
    strSql = strSql & ", ordered_qty"
    strSql = strSql & ", item_cost"
    strSql = strSql & ", uom_cd"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", inspection_req_fl"
    strSql = strSql & ", requested_del_dt"
    strSql = strSql & ", received_qty"
    strSql = strSql & ", rejected_qty"
    strSql = strSql & ", invoiced_qty"
    strSql = strSql & ", item_tx"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & ", rec_complete_var_pct"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_po_id, "N")
    strSql = strSql & "," & fInsertVariable(p_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_drop_ship_so_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_item_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    If Me.p_overrideMinQtyFl = True Then
        strSql = strSql & "," & fInsertVariable("Y", "S")
    Else
        strSql = strSql & "," & fInsertVariable("N", "S")
    End If
    strSql = strSql & "," & fInsertVariable(p_ordered_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_item_cost, "N")
    strSql = strSql & "," & fInsertVariable(p_uom_cd, "S")
    strSql = strSql & ",'PL'"
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_inspection_req_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_requested_del_dt, "D")
    strSql = strSql & "," & fInsertVariable(0, "N")
    strSql = strSql & "," & fInsertVariable(0, "N")
    strSql = strSql & "," & fInsertVariable(0, "N")
    strSql = strSql & "," & fInsertVariable(p_item_text, "S")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Me.p_recCompleteVarPct, "N")
    strSql = strSql & ")"
  
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
   
    Exit Sub
errHandler:
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

Private Sub IssueUpdateLastCost()
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double

    strRoutine_nm = "purchase_order_item.cls.IssueUpdateLastCost"
    
    strSql = " UPDATE item SET " & _
             " " & fUpdateVariable("last_cost", p_item_cost, "N", "=") & _
             " WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
             
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
    
    Exit Sub
errHandler:
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

Private Sub IssueUpdateData()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double

    strRoutine_nm = "purchase_order_item.cls.IssueUpdateData"
    
    strSql = "UPDATE purchase_order_item set "
    strSql = strSql & " " & fUpdateVariable("item_desc", p_item_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_item_nbr", p_supplier_item_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("requested_del_dt", p_requested_del_dt, "D", "=")
    
    If Me.p_overrideMinQtyFl = True Then
        strSql = strSql & "," & fUpdateVariable("override_min_qty_fl", "Y", "S", "=")
    Else
        strSql = strSql & "," & fUpdateVariable("override_min_qty_fl", "N", "S", "=")
    End If
    
    strSql = strSql & "," & fUpdateVariable("ordered_qty", p_ordered_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_cost", p_item_cost, "N", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("inspection_req_fl", p_inspection_req_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("invoiced_qty", p_invoiced_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_tx", p_item_text, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Me.o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & "," & fUpdateVariable("rec_complete_var_pct", Me.p_recCompleteVarPct, "N", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("po_id", p_po_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
    
    Exit Sub
errHandler:
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

Private Sub IssueDeleteData()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "purchase_order_item.cls.IssueDeleteData"

    strSql = "delete from purchase_order_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("po_id", p_po_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
   
    Exit Sub
errHandler:
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

Private Function fGetLineNbr(ByVal v_varPo_id As Variant)
    
    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim objRecordSet As ADODB.Recordset
    Dim strRoutine_nm As String
    Dim dblUpdatedQty As Double
    Dim strSql As String
    
    strRoutine_nm = "purchase_order_item.cls.fGetLineNbr"
    
    strSql = "SELECT max(line_nbr) + 1 FROM purchase_order_item " & _
             " WHERE " & fWhereVariable("po_id", v_varPo_id, "N", "=")
               
    Call Me.o_trans_data.OpenRecordset(objRecordSet, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Function
    End If
        
    If IsNull(objRecordSet(0)) = True Then
        fGetLineNbr = 1
    Else
        fGetLineNbr = objRecordSet(0)
    End If
    
    objRecordSet.Close
    Set objRecordSet = Nothing
    
    Exit Function
errHandler:
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
'Description:   This procedure checks the gets the received qty from the
'               inventory table.
'
'Arguments:
'//////////////////////////////////////////////////////////////////////////////
Private Function getReceivedQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblPoId As Double _
            , ByVal v_dblLineNbr As Double _
            , ByVal v_strDropShipFl As String _
            , ByVal v_varDropShipSalesOrdId As Variant _
            , ByVal v_varDropShipSoLineNbr As Variant _
            ) As Double

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.getReceivedQty"
    
    Dim objRecordSet As ADODB.Recordset, _
        strSql As String
    
    If v_strDropShipFl = "N" Then
        strSql = " SELECT sum(adj_qty) " & _
                 " FROM inventory_adj " & _
                 " WHERE " & fWhereVariable("po_id", v_dblPoId, "N", "=") & _
                 " AND " & fWhereVariable("po_line_nbr", v_dblLineNbr, "N", "=")
    Else
        strSql = " SELECT sum(loaded_qty) " & _
                 " FROM shipment_item " & _
                 " WHERE " & fWhereVariable("sales_ord_id", v_varDropShipSalesOrdId, "N", "=") & _
                 " AND " & fWhereVariable("line_nbr", v_varDropShipSoLineNbr, "N", "=")
    End If
        
    Call r_objTransData.OpenRecordset(objRecordSet, strSql, r_objError)
    
    If IsNull(objRecordSet(0)) = True Then
        getReceivedQty = 0
    Else
        getReceivedQty = objRecordSet(0)
    End If
     
    objRecordSet.Close
    Set objRecordSet = Nothing
    
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

Public Sub getPurchaseOrderItems()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_items.cls.getPurchaseOrderItems"
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
    
    strSql = " SELECT " & _
             "      purchase_order_item.po_id " & _
             "     ,purchase_order_item.line_nbr " & _
             "     ,purchase_order_item.status_cd " & _
             "     , purchase_order_item.item_cost  " & _
             "     ,status.code_desc as status_desc " & _
             "     ,item.item_nbr " & _
             "     ,purchase_order_item.item_desc " & _
             "     ,purchase_order_item.requested_del_dt " & _
             "     ,purchase_order_item.ordered_qty " & _
             "     ,purchase_order_item.override_min_qty_fl " & _
             "     ,purchase_order_item.received_qty " & _
             "     ,supplier.supplier_nbr " & _
             "     ,purchase_order_item.item_cost * purchase_order_item.ordered_qty as line_am " & _
             " FROM purchase_order_item " & _
             " LEFT OUTER JOIN code_desc status ON " & _
             "     purchase_order_item.status_cd = status.code " & _
             " AND status.field_nm = 'purchase_order_item_status_cd' " & _
             " LEFT OUTER JOIN item ON " & _
             "     purchase_order_item.item_id = item.item_id " & _
             " LEFT OUTER JOIN purchase_order ON " & _
             "     purchase_order_item.po_id = purchase_order.po_id " & _
             " LEFT OUTER JOIN supplier ON " & _
             "     purchase_order.supplier_id = supplier.supplier_id " & _
             " WHERE 1 = 1 "
             
    If fValuePresent(p_po_id) = True Then
        strSql = strSql & "  AND " & fWhereVariable("purchase_order_item.po_id", p_po_id, "N", "=")
    End If
    
    If fValuePresent(p_status_cd) = True Then
        strSql = strSql & "  AND " & fWhereVariable("purchase_order_item.status_cd", p_status_cd, "S", "=")
    End If
    
    If fValuePresent(p_item_nbr) = True Then
        strSql = strSql & "  AND " & fWhereVariable("item.item_nbr", p_item_nbr, "S", "=")
    End If
    
    If fValuePresent(p_supplier_nbr) = True Then
        strSql = strSql & "  AND " & fWhereVariable("supplier.supplier_nbr", p_supplier_nbr, "S", "=")
    End If
    
    If fValuePresent(p_start_dt) = True Then
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("purchase_order_item.requested_del_dt", p_start_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("purchase_order_item.requested_del_dt", p_end_dt, "D", "<=") & ")"
        Else
            strSql = strSql & "  AND " & fWhereVariable("purchase_order_item.requested_del_dt", p_start_dt, "D", ">=")
        End If
    Else
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND " & fWhereVariable("purchase_order_item.requested_del_dt", p_end_dt, "D", "<=")
        End If
    End If
        
    strSql = strSql & " ORDER BY purchase_order_item.requested_del_dt "
        
    Call Me.o_trans_data.OpenRecordset(objRecordSet, strSql, Me.o_error)

    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Set objRecordSet = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordSet
    End If

    Exit Sub
errHandler:
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

Private Sub IssueUpdateStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblPoId As Double _
            , ByVal v_dblLineNbr As Double _
            , ByVal v_strStatusCd As String _
            )
    
    On Error GoTo errHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_items.cls.IssueUpdateStatus"
    
    Dim strSql As String, _
        dblUpdatedQty As Double
    
    strSql = "UPDATE purchase_order_item set " & _
           "  " & fUpdateVariable("status_cd", v_strStatusCd, "S", "=") & _
           " ," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=") & _
           " ," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
           " WHERE " & fWhereVariable("po_id", v_dblPoId, "N", "=") & _
           " and " & fWhereVariable("line_nbr", v_dblLineNbr, "N", "=") & _
           " and " & fWhereVariable("status_cd", "CA", "S", "<>")
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedQty, r_objError)
    
    If dblUpdatedQty = 0 Then
        r_objError.p_type_cd = "F"
        r_objError.p_err_cd = "0200"
        r_objError.p_routine_nm = strRoutine_nm
        r_objError.p_message_id = 1005
        Exit Sub
    End If

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


' make sure ordered qty > the min purchase qty.  make sure it is in a valid increment
Private Sub UpdateOrderedQty()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.UpdateOrderedQty"
    
    If o_purchase_order.p_drop_ship_fl = "Y" Or Me.p_overrideMinQtyFl = True Then
        Exit Sub
    End If
    
    Dim intIncrQty As Double
    Dim dblIncrQty As Double
    Dim objSuppItem As supplier_item
    
    Set objSuppItem = New supplier_item
    Set objSuppItem.o_error = Me.o_error
    Set objSuppItem.o_trans_data = Me.o_trans_data
    objSuppItem.p_item_id = p_item_id
    objSuppItem.p_supplier_id = o_purchase_order.p_supplier_id
    Call objSuppItem.GetSupplierItem
    If IsNull(objSuppItem.p_item_id) = True Then
        Set o_item = CreateObject("converge_item.item")
        Set o_item.o_error = Me.o_error
        Set o_item.o_trans_data = Me.o_trans_data
        o_item.p_item_id = p_item_id
        Call o_item.GetItem
        If IsNull(o_item.p_po_min_qty) = False Then
            If p_ordered_qty < o_item.p_po_min_qty Then
                p_ordered_qty = o_item.p_po_min_qty
            End If
        End If
        If IsNull(o_item.p_po_increment_qty) = True Then
            intIncrQty = 1
        Else
            intIncrQty = o_item.p_po_increment_qty
        End If
    Else
        If IsNull(objSuppItem.p_min_purchase_qty) = False Then
            If p_ordered_qty < objSuppItem.p_min_purchase_qty Then
                p_ordered_qty = objSuppItem.p_min_purchase_qty
            End If
        End If
        If IsNull(objSuppItem.p_incr_purchase_qty) = True Then
            intIncrQty = 1
        Else
            intIncrQty = objSuppItem.p_incr_purchase_qty
        End If
    End If
    
    If intIncrQty > 0 Then
        dblIncrQty = p_ordered_qty / intIncrQty
    Else
        dblIncrQty = p_ordered_qty
    End If
    
    If dblIncrQty > Int(dblIncrQty) Then
        p_ordered_qty = (Int(dblIncrQty) + 1) * intIncrQty
    End If
    
    Exit Sub
errHandler:
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

Private Sub getPurchaseOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.getPurchaseOrder"
    
    If Me.o_purchase_order Is Nothing Then
        Set Me.o_purchase_order = New purchase_order
        Set Me.o_purchase_order.o_error = Me.o_error
        Set Me.o_purchase_order.o_trans_data = Me.o_trans_data
        Me.o_purchase_order.p_po_id = Me.p_po_id
        Me.o_purchase_order.p_bypass_nm_addr_fl = "Y"
        Me.o_purchase_order.getPurchaseOrder
    End If

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub getOldPurchaseOrderItem()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.getOldPurchaseOrderItem"

    Set Me.o_old_purchase_order_item = New purchase_order_item
    Set Me.o_old_purchase_order_item.o_error = Me.o_error
    Set Me.o_old_purchase_order_item.o_trans_data = Me.o_trans_data
    Me.o_old_purchase_order_item.p_po_id = Me.p_po_id
    Me.o_old_purchase_order_item.p_line_nbr = Me.p_line_nbr
    Call Me.o_old_purchase_order_item.GetPurchaseOrderItem
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub addOnOrderInventory(ByVal v_varOn_Ord_qty As Variant)
    
    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim objOnOrdInvAdj As Object
    
    strRoutine_nm = "purchase_order_item.cls.addOnOrderInventory"
    
    If o_purchase_order.p_drop_ship_fl = "Y" Then
        Exit Sub
    End If
    
    Set objOnOrdInvAdj = CreateObject("converge_inv.on_ord_inv_adj")
    Set objOnOrdInvAdj.o_error = Me.o_error
    Set objOnOrdInvAdj.o_trans_data = Me.o_trans_data
    With objOnOrdInvAdj
        .p_po_id = p_po_id
        .p_on_ord_qty = v_varOn_Ord_qty
        .p_po_line_nbr = p_line_nbr
        .p_assem_ord_id = Null
        .p_inv_loc_id = o_purchase_order.p_inv_loc_id
        .p_item_id = p_item_id
        .Add
    End With
    Set objOnOrdInvAdj = Nothing

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub changeOnOrderInventory( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByVal v_dblPoId As Double _
                , ByVal v_dblLineNbr As Double _
                , ByVal v_dblOrderedQty As Double _
                , ByVal v_dblReceivedQty As Double _
                , ByVal v_strDropShipFl As String _
                , ByVal v_strStatusCd As String _
                , ByVal v_strInvLocId As String _
                , ByVal v_dblItemId As Double _
                )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.changeOnOrderInventory"
    
    Dim objOnOrdInvAdj As Object
    
    If v_strDropShipFl = "Y" Then
        Exit Sub
    End If
    
    Set objOnOrdInvAdj = CreateObject("converge_inv.on_ord_inv_adj")
    With objOnOrdInvAdj
        Set .o_error = Me.o_error
        Set .o_trans_data = Me.o_trans_data
        .p_po_id = v_dblPoId
        If v_strStatusCd = "CL" Or _
           v_strStatusCd = "CP" Or _
           v_strStatusCd = "CA" Then
            .p_on_ord_qty = 0
        Else
            .p_on_ord_qty = v_dblOrderedQty - v_dblReceivedQty
            If .p_on_ord_qty < 1 Then
                .p_on_ord_qty = 0
            End If
        End If
        .p_po_line_nbr = v_dblLineNbr
        .p_assem_ord_id = Null
        .p_inv_loc_id = v_strInvLocId
        .p_item_id = v_dblItemId
        .Change
    End With
    Set objOnOrdInvAdj = Nothing

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

' when a purchase order items cost is changed, all the lots that have been recieved on that
' poi must have there cost updated also.  This routine gets a list of lots recieved
' on the poi, then updates there cost.
Private Sub updateAffectedLots( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblPoId As Double _
            , ByVal v_dblLineNbr As Double _
            , ByVal v_dblItemCost As Double _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item.cls.updateAffectedLots"
    
    Dim objLot As Object
    Dim rstLots As ADODB.Recordset
    Set objLot = CreateObject("converge_inv.lot")
    
    Set objLot.o_error = r_objError
    Set objLot.o_trans_data = r_objTransData
    objLot.p_po_id = v_dblPoId
    objLot.p_po_line_nbr = v_dblLineNbr
    objLot.p_item_nbr = Null
    objLot.p_po_receipt_id = Null
    objLot.p_assem_ord_id = Null
    objLot.p_prod_run_id = Null
    objLot.p_phy_inv_id = Null
    objLot.p_supplier_lot_id = Null
    Call objLot.getLots
    
    Set rstLots = objLot.o_record_set
    
    Do While rstLots.EOF = False And _
        r_objError.type_cd <> "E" And _
        r_objError.type_cd <> "F"

        Set objLot.o_error = r_objError
        Set objLot.o_trans_data = r_objTransData
        objLot.p_lot_id = rstLots("lot_id")
        objLot.p_unit_cost = v_dblItemCost
        objLot.p_status_cd = rstLots("status_cd")
        objLot.p_supplier_lot_id = Empty
        Set objLot.o_record_set_input = Nothing
        Call objLot.Change
        Call rstLots.MoveNext
    Loop
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub



