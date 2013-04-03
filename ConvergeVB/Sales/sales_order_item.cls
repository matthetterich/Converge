VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_order_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_varSales_ord_id As Variant
Private m_varLine_nbr As Variant
Private m_vardrop_ship_po_Line_nbr As Variant
Private m_varSec_leg_shipped_qty As Variant
Private m_varCust_item_nbr As Variant
Private m_varItem_id As Variant
Private m_varItem_desc As Variant
Private m_varItem_tx As Variant
Private m_varItem_nbr As Variant
Private m_varItem_price As Variant
Private m_varItem_type_cd As Variant
Private m_varItem_type_desc As Variant
Private m_varOrdered_qty As Variant
Private m_varPayment_terms_cd As Variant
Private m_varPayment_terms_desc As Variant
Private m_varline_am As Variant
Private m_varRequested_del_dt As Variant
Private m_varPlanned_ship_dt As Variant
Private m_varPlanned_ship_qty As Variant
Private m_varShip_complete_var_pct As Variant
Private m_varShipped_qty As Variant
Private m_varStatus_cd As Variant
Private m_varStatus2_cd As Variant
Private m_varStatus3_cd As Variant
Private m_varStatus4_cd As Variant
Private m_varStatus_desc As Variant
Private m_varSupplier_id As Variant
Private m_varSupplier_nbr As Variant
Private m_varType_cd As Variant
Private m_varType_desc As Variant
Private m_varUom_cd As Variant
Private m_varUom_desc As Variant
Private m_varDrop_ship_fl As Variant
Private m_varAlloc_priority_nbr As Variant
Private m_varp_cust_id As Variant
Private m_varp_cust_nbr As Variant
Private m_varp_start_dt As Variant
Private m_varp_end_dt As Variant
Private m_varp_bo_qty_fl As Variant
Private m_varp_cust_po_nbr As Variant
Private m_objError As Object
Private m_objTrans_data As Object
Private m_objMatReq As Object
Private m_objItem As Object
Private m_objSupplier As Object
Private m_objSalesOrder As sales_order
Private m_objSalesOrderItem As sales_order_item
Private m_objRecordset As adodb.Recordset
Private m_varp_ship_on_dt As Variant
'

Public Property Let p_cust_po_nbr(ByVal vData As Variant)
    m_varp_cust_po_nbr = vData
End Property
Public Property Get p_cust_po_nbr() As Variant
        p_cust_po_nbr = m_varp_cust_po_nbr
End Property

Public Property Let p_bo_qty_fl(ByVal vData As Variant)
    m_varp_bo_qty_fl = vData
End Property
Public Property Get p_bo_qty_fl() As Variant
        p_bo_qty_fl = m_varp_bo_qty_fl
End Property

Public Property Let p_start_dt(ByVal vData As Variant)
    m_varp_start_dt = vData
End Property
Public Property Get p_start_dt() As Variant
        p_start_dt = m_varp_start_dt
End Property

Public Property Let p_drop_ship_fl(ByVal vData As Variant)
    m_varDrop_ship_fl = vData
End Property
Public Property Get p_drop_ship_fl() As Variant
        p_drop_ship_fl = m_varDrop_ship_fl
End Property

Public Property Let p_end_dt(ByVal vData As Variant)
    m_varp_end_dt = vData
End Property
Public Property Get p_end_dt() As Variant
        p_end_dt = m_varp_end_dt
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    m_varp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
        p_cust_id = m_varp_cust_id
End Property
Public Property Set o_sales_order(ByVal vData As sales_order)
    Set m_objSalesOrder = vData
End Property
Public Property Get o_sales_order() As sales_order
        Set o_sales_order = m_objSalesOrder
End Property
Public Property Set o_old_sales_order_item(ByVal vData As sales_order_item)
    Set m_objSalesOrderItem = vData
End Property
Public Property Get o_old_sales_order_item() As sales_order_item
        Set o_old_sales_order_item = m_objSalesOrderItem
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set m_objRecordset = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = m_objRecordset
End Property
Public Property Let p_shipped_qty(ByVal vData As Variant)
    m_varShipped_qty = vData
End Property
Public Property Get p_shipped_qty() As Variant
        p_shipped_qty = m_varShipped_qty
End Property
Public Property Let p_ship_complete_var_pct(ByVal vData As Variant)
    m_varShip_complete_var_pct = vData
End Property
Public Property Get p_ship_complete_var_pct() As Variant
        p_ship_complete_var_pct = m_varShip_complete_var_pct
End Property
Public Property Let p_requested_del_dt(ByVal vData As Variant)
    m_varRequested_del_dt = vData
End Property
Public Property Get p_requested_del_dt() As Variant
        p_requested_del_dt = m_varRequested_del_dt
End Property
Public Property Let p_planned_ship_dt(ByVal vData As Variant)
    m_varPlanned_ship_dt = vData
End Property
Public Property Get p_planned_ship_dt() As Variant
        p_planned_ship_dt = m_varPlanned_ship_dt
End Property
Public Property Let p_planned_ship_qty(ByVal vData As Variant)
    m_varPlanned_ship_qty = vData
End Property
Public Property Get p_planned_ship_qty() As Variant
        p_planned_ship_qty = m_varPlanned_ship_qty
End Property
Public Property Let p_sec_leg_shipped_qty(ByVal vData As Variant)
    m_varSec_leg_shipped_qty = vData
End Property
Public Property Get p_sec_leg_shipped_qty() As Variant
    p_sec_leg_shipped_qty = m_varSec_leg_shipped_qty
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    m_varStatus_desc = vData
End Property
Public Property Get p_status_desc() As Variant
    p_status_desc = m_varStatus_desc
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    m_varStatus_cd = vData
End Property
Public Property Get p_status_cd() As Variant
    p_status_cd = m_varStatus_cd
End Property

Public Property Let p_status2_cd(ByVal vData As Variant)
    m_varStatus2_cd = vData
End Property
Public Property Get p_status2_cd() As Variant
    p_status2_cd = m_varStatus2_cd
End Property

Public Property Let p_status3_cd(ByVal vData As Variant)
    m_varStatus3_cd = vData
End Property
Public Property Get p_status3_cd() As Variant
    p_status3_cd = m_varStatus3_cd
End Property

Public Property Let p_status4_cd(ByVal vData As Variant)
    m_varStatus4_cd = vData
End Property
Public Property Get p_status4_cd() As Variant
    p_status4_cd = m_varStatus4_cd
End Property

Public Property Let p_item_desc(ByVal vData As Variant)
    m_varItem_desc = vData
End Property
Public Property Get p_item_desc() As Variant
    p_item_desc = m_varItem_desc
End Property
Public Property Let p_item_tx(ByVal vData As Variant)
    m_varItem_tx = vData
End Property
Public Property Get p_item_tx() As Variant
    p_item_tx = m_varItem_tx
End Property

Public Property Let p_item_type_cd(ByVal vData As Variant)
    m_varItem_type_cd = vData
End Property
Public Property Get p_item_type_cd() As Variant
    p_item_type_cd = m_varItem_type_cd
End Property
Public Property Let p_item_type_Desc(ByVal vData As Variant)
    m_varItem_type_desc = vData
End Property
Public Property Get p_item_type_Desc() As Variant
    p_item_type_Desc = m_varItem_type_desc
End Property
Public Property Let p_type_desc(ByVal vData As Variant)
    m_varType_desc = vData
End Property
Public Property Get p_type_desc() As Variant
    p_type_desc = m_varType_desc
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    m_varType_cd = vData
End Property
Public Property Get p_type_cd() As Variant
    p_type_cd = m_varType_cd
End Property
Public Property Set o_supplier(ByVal vData As Object)
    Set m_objSupplier = vData
End Property
Public Property Get o_supplier() As Object
    Set o_supplier = m_objSupplier
End Property
Public Property Set o_item(ByVal vData As Object)
    Set m_objItem = vData
End Property
Public Property Get o_item() As Object
    Set o_item = m_objItem
End Property
Public Property Set o_mat_req(ByVal vData As Object)
    Set m_objMatReq = vData
End Property
Public Property Get o_mat_req() As Object
    Set o_mat_req = m_objMatReq
End Property
Public Property Let p_supplier_nbr(ByVal vData As Variant)
    m_varSupplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
    p_supplier_nbr = m_varSupplier_nbr
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    m_varSupplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
    p_supplier_id = m_varSupplier_id
End Property
Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    m_varPayment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
    p_payment_terms_desc = m_varPayment_terms_desc
End Property
Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    m_varPayment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
    p_payment_terms_cd = m_varPayment_terms_cd
End Property
Public Property Let p_line_am(ByVal vData As Variant)
    m_varline_am = vData
End Property
Public Property Get p_line_am() As Variant
    p_line_am = m_varline_am
End Property
Public Property Let p_uom_desc(ByVal vData As Variant)
    m_varUom_desc = vData
End Property
Public Property Get p_uom_desc() As Variant
    p_uom_desc = m_varUom_desc
End Property
Public Property Let p_uom_cd(ByVal vData As Variant)
    m_varUom_cd = vData
End Property
Public Property Get p_uom_cd() As Variant
    p_uom_cd = m_varUom_cd
End Property
Public Property Let p_item_price(ByVal vData As Variant)
    m_varItem_price = vData
End Property
Public Property Get p_item_price() As Variant
    p_item_price = m_varItem_price
End Property
Public Property Let p_ordered_qty(ByVal vData As Variant)
    m_varOrdered_qty = vData
End Property
Public Property Get p_ordered_qty() As Variant
    p_ordered_qty = m_varOrdered_qty
End Property
Public Property Let p_item_nbr(ByVal vData As Variant)
    m_varItem_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
    p_item_nbr = m_varItem_nbr
End Property
Public Property Let p_cust_nbr(ByVal vData As Variant)
    m_varp_cust_nbr = vData
End Property
Public Property Get p_cust_nbr() As Variant
    p_cust_nbr = m_varp_cust_nbr
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    m_varItem_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = m_varItem_id
End Property
Public Property Let p_cust_item_nbr(ByVal vData As Variant)
    m_varCust_item_nbr = vData
End Property
Public Property Get p_cust_item_nbr() As Variant
    p_cust_item_nbr = m_varCust_item_nbr
End Property
Public Property Let p_line_nbr(ByVal vData As Variant)
    m_varLine_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
    p_line_nbr = m_varLine_nbr
End Property
Public Property Let p_drop_ship_po_line_nbr(ByVal vData As Variant)
    m_vardrop_ship_po_Line_nbr = vData
End Property
Public Property Get p_drop_ship_po_line_nbr() As Variant
    p_drop_ship_po_line_nbr = m_vardrop_ship_po_Line_nbr
End Property
Public Property Let p_sales_ord_id(ByVal vData As Variant)
    m_varSales_ord_id = vData
End Property
Public Property Get p_sales_ord_id() As Variant
    p_sales_ord_id = m_varSales_ord_id
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
Public Property Let p_alloc_priority_nbr(ByVal vData As Variant)
    m_varAlloc_priority_nbr = vData
End Property
Public Property Get p_alloc_priority_nbr() As Variant
    p_alloc_priority_nbr = m_varAlloc_priority_nbr
End Property

Public Property Let p_ship_on_dt(ByVal vData As Variant)
    m_varp_ship_on_dt = vData
End Property
Public Property Get p_ship_on_dt() As Variant
        p_ship_on_dt = m_varp_ship_on_dt
End Property



Public Sub GetSalesOrderLineItem()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
        
    strRoutine_nm = "sales_order_item.cls.GetSalesOrderLineItem"
    
    Call SelectSalesOrderItem(objRecordset)
    
    If objRecordset.EOF = True Then
        p_line_nbr = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_sales_ord_id = objRecordset("sales_ord_id")
        p_line_nbr = objRecordset("line_nbr")
        p_drop_ship_po_line_nbr = objRecordset("drop_ship_po_line_nbr")
        p_sec_leg_shipped_qty = objRecordset("sec_leg_shipped_qty")
        p_type_cd = objRecordset("type_cd")
        p_cust_item_nbr = objRecordset("cust_item_nbr")
        p_item_id = objRecordset("item_id")
        p_item_desc = objRecordset("item_desc")
        p_item_tx = objRecordset("item_tx")
        p_ordered_qty = objRecordset("ordered_qty")
        p_item_price = objRecordset("item_price")
        p_item_type_cd = objRecordset("item_type_cd")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_uom_cd = objRecordset("uom_cd")
        p_requested_del_dt = objRecordset("requested_del_dt")
        p_planned_ship_qty = objRecordset("planned_ship_qty")
        p_planned_ship_dt = objRecordset("planned_ship_dt")
        p_status_cd = objRecordset("status_cd")
        p_supplier_id = objRecordset("supplier_id")
        p_shipped_qty = objRecordset("shipped_qty")
        p_line_am = p_ordered_qty * p_item_price
        p_ship_complete_var_pct = objRecordset("ship_complete_var_pct")
        p_ship_on_dt = objRecordset("ship_on_dt")
                
        objRecordset.Close
        Set objRecordset = Nothing
        
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_supplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
        p_uom_desc = fGetDesc("uom_cd", p_uom_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("soi_status_cd", p_status_cd, o_error, o_trans_data)
        p_type_desc = fGetDesc("soi_type_cd", p_type_cd, o_error, o_trans_data)
        p_item_type_Desc = fGetDesc("item_type_cd", p_item_type_cd, o_error, o_trans_data)
        
        If p_type_cd = "MI" Then
            p_item_nbr = "Misc"
        Else
            p_item_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
        End If
        
        Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
        
        If Me.p_type_cd = "SI" And p_status_cd <> "PL" Then
            Set Me.o_mat_req.o_error = o_error
            Set Me.o_mat_req.o_trans_data = o_trans_data
            Me.o_mat_req.mat_req_id = Empty
            Me.o_mat_req.sales_ord_id = p_sales_ord_id
            Me.o_mat_req.so_line_nbr = p_line_nbr
            Me.o_mat_req.assem_ord_id = Empty
            Me.o_mat_req.ac_line_nbr = Empty
            Call Me.o_mat_req.GetMatReq
        Else
            Me.o_mat_req.mat_req_id = Null
        End If
        
    End If

    Exit Sub
ErrHandler:
    With o_error
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
           
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    strRoutine_nm = "sales_order_item.cls.Add"
    
    Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
    
    Call getSalesOrder(p_sales_ord_id)
    Call validate("add")
    Call IssueInsertData
    Call IssueUpdateLastPrice
    
    'if order item was added to a sales order that is already open, automatically open
    'the order item.
    
    If o_sales_order.p_status_cd = "OP" Or _
       o_sales_order.p_status_cd = "PC" Or _
       o_sales_order.p_status_cd = "CP" Then
        Call GetOldSalesOrderItem(p_sales_ord_id, p_line_nbr)
        p_status_cd = "OP"
        Call openOrderItem
        Call Me.o_sales_order.setStatus(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id)
    End If
    
    If Me.o_sales_order.p_status_cd <> "PL" Then
        Call Me.o_sales_order.incrementRevision(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id, Me.o_sales_order, Me.o_trans_data.p_assoc_id, Now, "Added Item")
    End If
    
    Exit Sub
ErrHandler:
    With o_error
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
           
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    strRoutine_nm = "sales_order_item.cls.change"

    Call getSalesOrder(p_sales_ord_id)
    Call GetOldSalesOrderItem(p_sales_ord_id, p_line_nbr)
    Call validate("change")
    Call IssueUpdateData
    Call UpdateMatReq
    If (Me.o_old_sales_order_item.p_ordered_qty <> Me.p_ordered_qty And Len(Me.p_ordered_qty) > 0) Or _
       (Me.o_old_sales_order_item.p_ship_complete_var_pct <> Me.p_ship_complete_var_pct And Len(Me.p_ship_complete_var_pct) > 0) Then
        Call setStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , Me.o_sales_order.p_consignment_fl _
                )
    End If
    
    If Me.o_sales_order.p_status_cd <> "PL" Then
        Call Me.o_sales_order.incrementRevision(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id, Me.o_sales_order, Me.o_trans_data.p_assoc_id, Now, "Changed Item")
    End If
    
    Exit Sub
ErrHandler:
    With o_error
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

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    strRoutine_nm = "sales_order_item.cls.delete"
    
    Call getSalesOrder(p_sales_ord_id)
    Call GetOldSalesOrderItem(p_sales_ord_id, p_line_nbr)
    Call validate("delete")
    Call IssueDelete
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
Public Sub setDropShipPoLineNbr(ByVal v_strStatusCd As String)
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "converge_sales.setDropShipPoLineNbr"
    
    If v_strStatusCd = "CA" Or _
       v_strStatusCd = "CL" Or _
       v_strStatusCd = "OH" Then
    Else
        v_strStatusCd = "PC"
    End If
    
    strSql = "update sales_order_item set "
    strSql = strSql & " " & fUpdateVariable("drop_ship_po_line_nbr", p_drop_ship_po_line_nbr, "N", "=")
    strSql = strSql & ", " & fUpdateVariable("status_cd", v_strStatusCd, "S", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & " and " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    If dblUpdateQty = 0 Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "1000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1013
    End If
    
    Call Me.UpdateQuantities(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id, Me.p_line_nbr)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear

End Sub



Public Sub UpdateQuantities( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef v_dblSalesOrdId As Double _
            , ByRef v_dblLineNbr As Double _
            )
    
    On Error GoTo ErrHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.UpdateQuantities"
    
    Dim objPurchaseOrderItem As Object, _
        strSql As String, _
        dblUpdateQty As Double, _
        varShippedQty As Variant, _
        varPlannedShipQty As Variant
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_sales_ord_id = v_dblSalesOrdId
    Me.p_line_nbr = v_dblLineNbr
    
    Call getSalesOrder(v_dblSalesOrdId)
    
    Call getShippedQty( _
                  r_objError _
                , r_objTransData _
                , v_dblSalesOrdId _
                , v_dblLineNbr _
                , varShippedQty _
                , varPlannedShipQty _
                )
    
    strSql = "update sales_order_item set " & _
                " " & fUpdateVariable("shipped_qty", varShippedQty, "N", "=") & _
                "," & fUpdateVariable("planned_ship_qty", varPlannedShipQty, "N", "=") & _
                "," & fUpdateVariable("sec_leg_shipped_qty", getSecLegShippedQty, "N", "=") & _
                " where " & _
                "      sales_ord_id = " & v_dblSalesOrdId & _
                " AND  line_nbr = " & v_dblLineNbr
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdateQty, r_objError)

    Call setStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblSalesOrdId _
                , v_dblLineNbr _
                , Me.o_sales_order.p_consignment_fl _
                )
                
    If Me.o_sales_order.p_drop_ship_fl = "Y" Then
       Call GetOldSalesOrderItem(v_dblSalesOrdId, v_dblLineNbr)
       If IsNull(Me.o_sales_order.p_drop_ship_po_id) = False And IsNull(Me.o_old_sales_order_item.p_drop_ship_po_line_nbr) = False Then
            Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
            Call objPurchaseOrderItem.UpdateQuantities( _
                                      Me.o_error _
                                    , Me.o_trans_data _
                                    , Me.o_sales_order.p_drop_ship_po_id _
                                    , Me.o_old_sales_order_item.p_drop_ship_po_line_nbr _
                                    )
        End If
    End If
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub changeStatus()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.ChangeStatus"
    
    Dim objRecordset As adodb.Recordset, _
        strStatusCd As String, _
        varOldStatus_cd As Variant
    
    Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
    Call getSalesOrder(Me.p_sales_ord_id)
    Call GetOldSalesOrderItem(Me.p_sales_ord_id, Me.p_line_nbr)
        
    If Me.o_sales_order.p_status_cd = "CL" And Me.p_status_cd <> "CL" Then
       With Me.o_error
            .type_cd = "F"
            .err_cd = "1601"
            .routine_nm = strRoutine_nm
            .err_desc = "Sales Order item status can't be changed because sales order is closed or cancelled"
            .message_id = 0
        End With
    End If
    
    If Me.o_sales_order.p_status_cd = "CA" And _
       Me.p_status_cd <> "CA" Then
       With Me.o_error
            .type_cd = "F"
            .err_cd = "1602"
            .routine_nm = strRoutine_nm
            .err_desc = "Sales Order item status can't be changed because sales order is closed or cancelled"
            .message_id = 0
        End With
    End If
    
    'once an item has been cancelled nothing else can be done to it.
    If Me.o_old_sales_order_item.p_status_cd = "CA" Then
        Exit Sub
    End If
    
    Select Case Me.p_status_cd
        Case "OP"
            Call openOrderItem
        Case "CL"
            Call closeOrderItem
        Case "CA"
            Call cancelOrderItem
        Case "OH"
            Call HoldOrderItem
    End Select
    
    Set Me.o_old_sales_order_item = Nothing
    
    If Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "CA" Then
        Call Me.o_sales_order.setStatus( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_sales_ord_id _
                    )
    End If
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1500"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub openOrderItem()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.OpenOrderItem"
    
    Dim strSql As String
    Dim dblUpdateQty As Double
    Dim objInventoryAlloc As Object
    
    If Me.o_old_sales_order_item Is Nothing Then
        Exit Sub
    End If
    
    If IsDate(Me.o_old_sales_order_item.p_requested_del_dt) = False Then
        o_error.type_cd = "E"
        o_error.err_cd = "1600"
        o_error.routine_nm = strRoutine_nm
        o_error.message_id = 1011
        Exit Sub
    End If
    
    If IsDate(Me.o_old_sales_order_item.p_planned_ship_dt) = False Then
        o_error.type_cd = "E"
        o_error.err_cd = "1800"
        o_error.routine_nm = strRoutine_nm
        o_error.message_id = 1012
        Exit Sub
    End If
    
    If IsNumeric(Me.o_old_sales_order_item.p_ship_complete_var_pct) = False Then
        o_error.type_cd = "E"
        o_error.err_cd = "2000"
        o_error.routine_nm = strRoutine_nm
        o_error.message_id = 1130
        Exit Sub
    End If
    
    If Me.o_old_sales_order_item.p_type_cd = "SI" Then
        If Len(Me.o_old_sales_order_item.p_uom_cd) > 0 Then
        Else
            o_error.type_cd = "F"
            o_error.err_cd = "2100"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            Exit Sub
        End If
            
        If Len(Me.o_old_sales_order_item.p_item_type_cd) > 0 Then
        Else
            o_error.type_cd = "F"
            o_error.err_cd = "2200"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            Exit Sub
        End If
    End If
        
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , "OP" _
                )
    
    'update status.  if the order is being reopened, the status needs to be recalculated
    Call setStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , Me.o_sales_order.p_consignment_fl _
                )
    
    'if this is not a drop ship order and it is a sales item, and sales order is a regular customer or sub
    'sales order, open the material requirement.
    If Me.o_sales_order.p_drop_ship_fl = "N" And Me.o_old_sales_order_item.p_type_cd = "SI" And _
       (Me.o_sales_order.p_type_cd = "SO" Or Me.o_sales_order.p_type_cd = "CO") Then
        
        If Me.o_old_sales_order_item.p_status_cd = "PL" Then
            Call Me.o_mat_req.addMatReq( _
                          Me.o_error _
                        , Me.o_trans_data _
                        , "SO" _
                        , Me.o_old_sales_order_item.p_sales_ord_id _
                        , Me.o_old_sales_order_item.p_line_nbr _
                        , Empty _
                        , Empty _
                        , Me.o_old_sales_order_item.p_item_id _
                        , Me.o_sales_order.p_inv_loc_id _
                        , Me.o_old_sales_order_item.p_ordered_qty _
                        , Me.o_old_sales_order_item.p_planned_ship_dt _
                        , Me.o_old_sales_order_item.p_ship_complete_var_pct _
                        , 0 _
                        , Me.o_old_sales_order_item.p_supplier_id _
                        )
        End If
        
        Call Me.o_mat_req.openMatReq( _
              Me.o_error _
            , Me.o_trans_data _
            , Empty _
            , "SO" _
            , Me.p_sales_ord_id _
            , Me.p_line_nbr _
            , Empty _
            , Empty _
            )
    
        Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
        With objInventoryAlloc
            Set .o_error = Me.o_error
            Set .o_trans_data = Me.o_trans_data
            .p_mat_req_id = Me.o_mat_req.mat_req_id
            .p_item_id = Me.p_item_id
            .p_inv_loc_id = Me.o_sales_order.p_inv_loc_id
            .Allocate
        End With
        Set objInventoryAlloc = Nothing
    End If
    
    
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub closeOrderItem()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "sales_order_item.closeitem"

    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , "CL" _
                )
    
    If Me.o_sales_order.p_drop_ship_fl = "N" And Me.o_old_sales_order_item.p_type_cd = "SI" And _
       (Me.o_sales_order.p_type_cd = "SO" Or Me.o_sales_order.p_type_cd = "CO") Then
        Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
        Call Me.o_mat_req.closeMatReq( _
              Me.o_error _
            , Me.o_trans_data _
            , Empty _
            , "SO" _
            , o_old_sales_order_item.p_sales_ord_id _
            , o_old_sales_order_item.p_line_nbr _
            , Empty _
            , Empty _
            )
    End If

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub cancelOrderItem()

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "sales_order_item.cls.cancelOrder"
    
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , "CA" _
                )
    
    
    If Me.o_sales_order.p_drop_ship_fl = "N" And Me.o_old_sales_order_item.p_type_cd = "SI" And _
       (Me.o_sales_order.p_type_cd = "SO" Or Me.o_sales_order.p_type_cd = "CO") Then
        Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
        Call Me.o_mat_req.cancelMatReq( _
              Me.o_error _
            , Me.o_trans_data _
            , Empty _
            , "SO" _
            , o_old_sales_order_item.p_sales_ord_id _
            , o_old_sales_order_item.p_line_nbr _
            , Empty _
            , Empty _
            )
    End If

    If Me.o_sales_order.p_status_cd <> "PL" Then
        Call Me.o_sales_order.incrementRevision(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id, Me.o_sales_order, Me.o_trans_data.p_assoc_id, Now, "Cancelled Item")
    End If
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "3000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub HoldOrderItem()

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "sales_order_item.cls.HoldOrderItem"
    
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_line_nbr _
                , "OH" _
                )
    
    
    If Me.o_sales_order.p_drop_ship_fl = "N" And Me.o_old_sales_order_item.p_type_cd = "SI" And _
       (Me.o_sales_order.p_type_cd = "SO" Or Me.o_sales_order.p_type_cd = "CO") Then
        Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
        Call Me.o_mat_req.holdMatReq( _
              Me.o_error _
            , Me.o_trans_data _
            , Empty _
            , "SO" _
            , Me.o_old_sales_order_item.p_sales_ord_id _
            , Me.o_old_sales_order_item.p_line_nbr _
            , Empty _
            , Empty _
            )
    End If
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "3200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub UpdateMatReq()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset

    strRoutine_nm = "sales_order_item.cls.updateMatReq"
    
    'if the sales order is open, we must also update the material requirement.
    If o_sales_order.p_drop_ship_fl = "Y" Or _
       Me.o_sales_order.p_type_cd = "BL" Or _
       Me.o_old_sales_order_item.p_status_cd = "PL" Then
        Exit Sub
    End If
        
    Call SelectSalesOrderItem(objRecordset)
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        objRecordset.Close
        Set objRecordset = Nothing
        o_error.type_cd = "F"
        o_error.err_cd = "3400"
        o_error.routine_nm = strRoutine_nm
        o_error.message_id = 0
        Exit Sub
    End If
    
    Me.p_sales_ord_id = objRecordset("sales_ord_id")
    Me.p_line_nbr = objRecordset("line_nbr")
    Me.p_drop_ship_po_line_nbr = objRecordset("drop_ship_po_line_nbr")
    Me.p_type_cd = objRecordset("type_cd")
    Me.p_ordered_qty = objRecordset("ordered_qty")
    Me.p_uom_cd = objRecordset("uom_cd")
    Me.p_requested_del_dt = objRecordset("requested_del_dt")
    Me.p_planned_ship_dt = objRecordset("planned_ship_dt")
    Me.p_status_cd = objRecordset("status_cd")
    Me.p_supplier_id = objRecordset("supplier_id")
    objRecordset.Close
    Set objRecordset = Nothing
    
    'always attempt to update the material requirement, if it
    'does not exist, it will just return control to here
    Set Me.o_mat_req = CreateObject("converge_prod.mat_req")
    Call Me.o_mat_req.Change( _
              Me.o_error _
            , Me.o_trans_data _
            , Empty _
            , "SO" _
            , Me.p_sales_ord_id _
            , Me.p_line_nbr _
            , Empty _
            , Empty _
            , Me.p_ordered_qty _
            , Me.p_planned_ship_dt _
            , Me.p_ship_complete_var_pct _
            , Empty _
            , Empty _
            , Empty _
            )
                
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "3600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub setStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId As Double _
            , ByVal v_dblLineNbr As Double _
            , ByVal v_strConsignmentFl As String _
            )

    On Error GoTo ErrHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.SetStatus"
    
    Dim dblUpdateQty As Double, _
        dblOrdered_qty As Double, _
        strStatusCd As String, _
        strSql As String, _
        objSalesOrder As sales_order
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_sales_ord_id = v_dblSalesOrdId
    Me.p_line_nbr = v_dblLineNbr
    Me.GetSalesOrderLineItem
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
        Exit Sub
    End If
    
    dblOrdered_qty = Me.p_ordered_qty
 
 '==>  if sales order item is in a planned, on hold, cancelled or closed status,
 '     the order must be manually reopened before we calculate the new status.
    If Me.p_status_cd = "PL" Or _
       Me.p_status_cd = "OH" Or _
       Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "CA" Then
        Exit Sub
    End If
    
    If IsNull(Me.p_ship_complete_var_pct) Or Me.p_ship_complete_var_pct = 0 Then
    Else
        dblOrdered_qty = dblOrdered_qty * (1 - Me.p_ship_complete_var_pct)
    End If
    
    strStatusCd = "OP"
    
    'if everything is shipped, change status to shipped complete
    If Me.p_drop_ship_fl = "N" Then
        If Me.p_planned_ship_qty >= dblOrdered_qty Then
            strStatusCd = "PC"
        End If
    Else
        If Me.p_status_cd = "PC" Then
            strStatusCd = "PC"
        End If
    End If
    
    'if everything is shipped, change status to shipped complete
    If Me.p_shipped_qty >= dblOrdered_qty Then
        If v_strConsignmentFl = "N" Then
            strStatusCd = "CP"
        Else
            strStatusCd = "SC"
        End If
    End If
    
    'on consignment orders, if everything is shipped on 1st & 2nd leg shipments, change status to invoiced complete
    If Me.p_shipped_qty >= dblOrdered_qty And Me.p_sec_leg_shipped_qty >= Me.p_shipped_qty Then
        strStatusCd = "CP"
    End If
        
    If Me.p_status_cd <> strStatusCd Then
        strSql = "update sales_order_item set "
        strSql = strSql & fUpdateVariable("status_cd", strStatusCd, "S", "=")
        strSql = strSql & " where "
        strSql = strSql & "   sales_ord_id  = " & v_dblSalesOrdId
        strSql = strSql & " and line_nbr = " & v_dblLineNbr
        Call r_objTransData.ExecuteDatabase(strSql, dblUpdateQty, r_objError)
    End If
    
    Set objSalesOrder = New sales_order
    
    Call objSalesOrder.setStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblSalesOrdId _
                )
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "3800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub validate(ByVal v_strFunction_cd As String)
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.Validate"
        
    If v_strFunction_cd = "add" Then
        If fValuePresent(p_sales_ord_id) = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "4000"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1010
            Exit Sub
        End If
        
        If o_sales_order.p_status_cd = "PL" Or _
           o_sales_order.p_status_cd = "OP" Or _
           o_sales_order.p_status_cd = "PC" Or o_sales_order.p_status_cd = "CP" Or _
           o_sales_order.p_status_cd = "OH" Then
        Else
            o_error.type_cd = "E"
            o_error.err_cd = "4200"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1014
            Exit Sub
        End If
           
        'if item type code is sales item, the item must be on the item table.  if type is not
        'sales item, p_item_id will be set to null
        If p_type_cd = "SI" Then
            Set o_item = CreateObject("converge_item.item")
            Set o_item.o_error = o_error
            Set o_item.o_trans_data = o_trans_data
            If IsNull(p_item_id) = False And p_item_id > 0 Then
                o_item.p_item_id = p_item_id
                o_item.getitem
                If IsNull(o_item.p_item_id) = True Then
                    o_error.type_cd = "E"
                    o_error.err_cd = "4400"
                    o_error.routine_nm = strRoutine_nm
                    o_error.message_id = 1045
                    Exit Sub
                End If
            Else
                o_item.p_item_id = Null
                o_item.p_item_nbr = p_item_nbr
                o_item.getitem
                If IsNull(o_item.p_item_id) = True Then
                    o_error.type_cd = "E"
                    o_error.err_cd = "4600"
                    o_error.routine_nm = strRoutine_nm
                    o_error.message_id = 1045
                    Exit Sub
                End If
                p_item_id = o_item.p_item_id
            End If
            p_uom_cd = o_item.p_uom_cd
            p_item_type_cd = o_item.p_type_cd
            If o_sales_order.p_drop_ship_fl = "Y" Then
                If p_item_type_cd = "PU" Then
                Else
                    o_error.type_cd = "E"
                    o_error.err_cd = "4800"
                    o_error.routine_nm = strRoutine_nm
                    o_error.message_id = 1009
                    Exit Sub
                End If
            End If
        Else
            p_item_id = Null
        End If

        If Len(p_uom_cd) > 0 Then
        Else
            o_error.type_cd = "F"
            o_error.err_cd = "5000"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1016
            Exit Sub
        End If
            
    End If
    
    If v_strFunction_cd = "add" Or v_strFunction_cd = "change" Then
        If fValuePresent(p_item_desc) = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "5200"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1055
            Exit Sub
        End If
    
        If IsNull(p_supplier_id) = True Then
            If IsNull(p_supplier_nbr) = True Then
            Else
               If Len(p_supplier_nbr) > 0 Then
                    p_supplier_id = fGetSupplierId(p_supplier_nbr, o_error, o_trans_data)
                    If IsNull(p_supplier_id) = True Then
                        o_error.type_cd = "E"
                        o_error.err_cd = "5400"
                        o_error.routine_nm = strRoutine_nm
                        o_error.message_id = 1007
                        Exit Sub
                    End If
               End If
            End If
        Else
            If p_supplier_id > 0 Then
                p_supplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
                If IsNull(p_supplier_id) = True Then
                    o_error.type_cd = "E"
                    o_error.err_cd = "5600"
                    o_error.routine_nm = strRoutine_nm
                    o_error.message_id = 1007
                    Exit Sub
                End If
            Else
               If Len(p_supplier_nbr) > 0 Then
                    p_supplier_id = fGetSupplierId(p_supplier_nbr, o_error, o_trans_data)
                    If IsNull(p_supplier_id) = True Then
                        o_error.type_cd = "E"
                        o_error.err_cd = "5800"
                        o_error.routine_nm = strRoutine_nm
                        o_error.message_id = 1007
                        Exit Sub
                    End If
               End If
            End If
        End If
    End If
    
    If v_strFunction_cd = "delete" Then
        'Items can only be deleted if the sales order is still in planned status
        If (o_sales_order.p_status_cd = "PL") And IsNull(o_sales_order.p_sales_ord_id) = False Then
        Else
            o_error.type_cd = "F"
            o_error.err_cd = "6000"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1040
            Exit Sub
        End If
    End If

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsertData()

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order_item.cls.IssueInsertData"
    
    p_line_nbr = fGetLineNbr(p_sales_ord_id)
    If Len(p_ship_complete_var_pct) > 0 Then
    Else
        p_ship_complete_var_pct = 0
    End If
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    strSql = "insert into sales_order_item( "
    strSql = strSql & "  sales_ord_id"
    strSql = strSql & ", line_nbr"
    strSql = strSql & ", sec_leg_shipped_qty"
    strSql = strSql & ", cust_item_nbr"
    strSql = strSql & ", item_id"
    strSql = strSql & ", item_desc"
    strSql = strSql & ", item_tx"
    strSql = strSql & ", item_price"
    strSql = strSql & ", item_type_cd"
    strSql = strSql & ", ordered_qty"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", requested_del_dt"
    strSql = strSql & ", planned_ship_dt"
    strSql = strSql & ", planned_ship_qty"
    strSql = strSql & ", ship_complete_var_pct"
    strSql = strSql & ", shipped_qty"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", supplier_id"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", uom_cd"
    strSql = strSql & ", ship_on_dt"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable(p_cust_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_item_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_item_price, "N")
    strSql = strSql & "," & fInsertVariable(p_item_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_ordered_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_requested_del_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_planned_ship_dt, "D")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable(p_ship_complete_var_pct, "N")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable("PL", "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_uom_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_ship_on_dt, "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateLastPrice()
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order_item.cls.IssueUpdateLastPrice"
    
    strSql = " UPDATE item SET " & _
             " " & fUpdateVariable("last_price", p_item_price, "N", "=") & _
             " WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateData()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order_item.cls.IssueUpdateData"
    
    strSql = "update sales_order_item set "
    strSql = strSql & " " & fUpdateVariable("cust_item_nbr", p_cust_item_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_desc", p_item_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_tx", p_item_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("requested_del_dt", p_requested_del_dt, "S", "=")
    strSql = strSql & "," & fUpdateVariable("planned_ship_dt", p_planned_ship_dt, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & "," & fUpdateVariable("ordered_qty", p_ordered_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_price", p_item_price, "N", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ship_complete_var_pct", p_ship_complete_var_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("ship_on_dt", p_ship_on_dt, "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    strSql = "update sales_order_item set "
    strSql = strSql & " " & fUpdateVariable("type_cd", p_type_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("uom_cd", p_uom_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_id", p_supplier_id, "N", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("status_cd", "PL", "S", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDelete()
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order_item.cls.IssueDeleteData"

    strSql = "delete from sales_order_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub SelectSalesOrderItem(ByRef objRecordset As adodb.Recordset)

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order_item.cls.SelectSalesOrderItem"

    strSql = "SELECT * "
    strSql = strSql & "FROM sales_order_item "
    strSql = strSql & " WHERE "
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Function fGetLineNbr(ByVal varSales_ord_id As Variant)
    
    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim objRecordset As adodb.Recordset
    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strRoutine_nm = "sales_order_item.cls.fGetLineNbr"
    
    strSql = "SELECT max(line_nbr) + 1 FROM sales_order_item " & _
             " WHERE " & fWhereVariable("sales_ord_id", varSales_ord_id, "N", "=")
               
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
        
    If IsNull(objRecordset(0)) = True Then
        fGetLineNbr = 1
    Else
        fGetLineNbr = objRecordset(0)
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Function

Private Sub getShippedQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId As Double _
            , ByVal v_dblLineNbr As Double _
            , ByRef r_varShippedQty As Variant _
            , ByRef r_varPlannedQty As Variant _
            )

    On Error GoTo ErrHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.getShippedQty"
    
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strSql = "SELECT sum(loaded_qty), sum(planned_qty)  "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE         "
    strSql = strSql & "  " & fWhereVariable("sales_ord_id", v_dblSalesOrdId, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", v_dblLineNbr, "N", "=")
    strSql = strSql & "and" & fWhereVariable("ship_type_cd", "2L", "S", "<>")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If IsNull(objRecordset(0)) = True Then
        r_varShippedQty = 0
    Else
        r_varShippedQty = objRecordset(0)
    End If
     
    If IsNull(objRecordset(1)) = True Then
        r_varPlannedQty = 0
    Else
        r_varPlannedQty = objRecordset(1)
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "7300"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Function getSecLegShippedQty() As Double

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order_item.cls.getSecLegShippedQty"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If o_sales_order.p_consignment_fl = "N" Then
        getSecLegShippedQty = 0
        Exit Function
    End If
    
    strSql = "SELECT sum(loaded_qty)  "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE         "
    strSql = strSql & "  " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    strSql = strSql & "and" & fWhereVariable("ship_type_cd", "2L", "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objRecordset(0)) = True Then
        getSecLegShippedQty = 0
    Else
        getSecLegShippedQty = objRecordset(0)
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Function

Private Sub getSalesOrder(ByVal v_varSales_ord_id As Variant)

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    
    strRoutine_nm = "sales_order_item.cls.getSalesOrder"
    
    If o_sales_order Is Nothing Then
        Set o_sales_order = New sales_order
        Set o_sales_order.o_error = o_error
        Set o_sales_order.o_trans_data = o_trans_data
        o_sales_order.p_bypass_nm_addr_fl = "Y"
        o_sales_order.p_sales_ord_id = v_varSales_ord_id
        Call o_sales_order.getSalesOrder
    End If

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub GetOldSalesOrderItem(ByVal v_varSales_ord_id As Variant, ByVal v_varLine_nbr As Variant)

    On Error GoTo ErrHandler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    
    strRoutine_nm = "sales_order_item.cls.GetOldSalesOrderItem"
    
    Set o_old_sales_order_item = New sales_order_item
    Set o_old_sales_order_item.o_error = o_error
    Set o_old_sales_order_item.o_trans_data = o_trans_data
    o_old_sales_order_item.p_sales_ord_id = v_varSales_ord_id
    o_old_sales_order_item.p_line_nbr = v_varLine_nbr
    Call o_old_sales_order_item.GetSalesOrderLineItem

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getSalesOrderLineItems()

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.getSalesOrderLineItems"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = "SELECT a.item_id, a.status_cd , a.line_nbr, a.ordered_qty, " & _
                " a.type_cd, a.item_desc, a.requested_del_dt, a.shipped_qty, " & _
                " b.status_cd as mq_status_cd, b.back_ordered_qty, " & _
                " b.ready_qty, b.consumed_qty, a.planned_ship_dt " & _
                " FROM sales_order_item a left join mat_req b " & _
                " on a.sales_ord_id = b.sales_ord_id " & _
                " and a.line_nbr = b.so_line_nbr " & _
                " WHERE " & _
                "       " & fWhereVariable("a.sales_ord_id", p_sales_ord_id, "N", "=")

    strSql = strSql & " Order by a.requested_del_dt "
    
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
   
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varLineNbr As Variant _
            , ByVal v_varStatusCd As Variant _
            )
    
    On Error GoTo ErrHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orderitem.cls.IssueUpdateStatus"
    
    Dim dblUpdateQty As Double, _
        strSql As String

    strSql = "update sales_order_item set "
    strSql = strSql & fUpdateVariable("status_cd", v_varStatusCd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("sales_ord_id", v_varSalesOrdId, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", v_varLineNbr, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdateQty, r_objError)
    
    If dblUpdateQty = 0 Then
        With r_objError
            .type_cd = "E"
            .err_cd = "8100"
            .routine_nm = strRoutine_nm
            .message_id = 1005
        End With
        Exit Sub
    End If
        
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "8200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getSalesOrderItems()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strCustNbr As String
    Dim strCustItemNbr As String
    Dim strItemNbr As String
    Dim strCustPoNbr As String
    strRoutine_nm = "Sales_Order_item.cls.getSalesOrderItems"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
   strSql = "SELECT  " & _
                " a.sales_ord_id, " & _
                " a.requested_del_dt, " & _
                " a.planned_ship_dt, " & _
                " a.ordered_qty, " & _
                " a.ship_complete_var_pct, " & _
                " a.item_desc, " & _
                " a.shipped_qty, " & _
                " a.type_cd as soi_type_cd, " & _
                " a.status_cd, " & _
                " b.item_nbr, " & _
                " b.item_id, " & _
                " c.type_cd, " & _
                " c.drop_ship_fl, " & _
                " d.cust_nbr, " & _
                " d.cust_id, "
strSql = strSql & " c.cust_po_nbr, " & _
                "f.cust_item_nbr, "
strSql = strSql & " e.status_cd as mq_status_cd, " & _
                " e.ready_qty, " & _
                " e.back_ordered_qty, " & _
                " aa.code_desc as soi_status_desc, " & _
                " cc.code_desc as mr_status_desc, " & _
                " bb.code_desc  as soi_type_desc " & _
                " FROM sales_order_item a " & _
                " left outer join item b on b.item_id = a.item_id" & _
                " left outer join sales_order c on c.sales_ord_id = a.sales_ord_id " & _
                " left outer join customer d on d.cust_id = c.cust_id " & _
                " left outer join cust_item f on d.cust_id = f.cust_id and a.item_id = f.item_id " & _
                " left outer join mat_req e on a.sales_ord_id = e.sales_ord_id and a.line_nbr = e.so_line_nbr " & _
                " left outer join code_desc as aa on aa.code = a.status_cd and aa.field_nm = 'soi_status_cd' " & _
                " left outer join code_desc as bb on bb.code = a.type_cd and bb.field_nm = 'soi_type_cd' " & _
                " left outer join code_desc as cc on cc.code = e.status_cd and cc.field_nm = 'mat_req_status_cd' " & _
                " where 1 = 1 "

    If fValuePresent(p_cust_nbr) = True Then
        If InStr(1, p_cust_nbr, "*", vbBinaryCompare) > 0 Then
            strCustNbr = Left(p_cust_nbr, InStr(1, p_cust_nbr, "*", vbBinaryCompare) - 1) & "%"
            strSql = strSql & " AND d.cust_nbr LIKE '" & strCustNbr & "'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("d.cust_nbr", p_cust_nbr, "S", "=")
        End If
    End If
    
    If fValuePresent(p_cust_po_nbr) = True Then
        If InStr(1, p_cust_po_nbr, "*", vbBinaryCompare) > 0 Then
            strCustPoNbr = Left(p_cust_po_nbr, InStr(1, p_cust_po_nbr, "*", vbBinaryCompare) - 1) & "%"
            strSql = strSql & " AND c.cust_po_nbr LIKE '" & strCustPoNbr & "'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("c.cust_po_nbr", p_cust_po_nbr, "S", "=")
        End If
    End If
    
    If fValuePresent(p_cust_item_nbr) = True Then
        If InStr(1, p_cust_item_nbr, "*", vbBinaryCompare) > 0 Then
            strCustItemNbr = Left(p_cust_item_nbr, InStr(1, p_cust_item_nbr, "*", vbBinaryCompare) - 1) & "%"
            strSql = strSql & " AND f.cust_item_nbr LIKE '" & strCustItemNbr & "'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("f.cust_item_nbr", p_cust_item_nbr, "S", "=")
        End If
    End If
    
    If fValuePresent(p_status_cd) = True Then
        If fValuePresent(p_status2_cd) = True Then
            If fValuePresent(p_status3_cd) = True Then
                If fValuePresent(p_status4_cd) = True Then
                    strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status2_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status3_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status4_cd, "S", "=")
                    strSql = strSql & ")"
                Else
                    strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status2_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status3_cd, "S", "=")
                    strSql = strSql & ")"
                End If
            Else
                strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status2_cd, "S", "=")
                strSql = strSql & ")"
            End If
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
        End If
    End If
    
    If fValuePresent(p_type_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("c.type_cd", p_type_cd, "S", "=")
    End If
    
    If fValuePresent(p_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.sales_ord_id", p_sales_ord_id, "S", "=")
    End If
    
    If fValuePresent(p_item_nbr) = True Then
        If InStr(1, p_item_nbr, "*", vbBinaryCompare) > 0 Then
            strItemNbr = Left(p_item_nbr, InStr(1, p_item_nbr, "*", vbBinaryCompare) - 1) & "%"
            strSql = strSql & " AND b.item_nbr LIKE '" & strItemNbr & "'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("b.item_nbr", p_item_nbr, "S", "=")
        End If
    End If
    
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("c.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
    
    If fValuePresent(p_bo_qty_fl) = True And p_bo_qty_fl = "on" Then
       strSql = strSql & "  AND " & fWhereVariable("e.back_ordered_qty", 0, "N", ">")
    End If
    
    If fValuePresent(p_start_dt) = True Then
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=") & ")"
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", ">=")
        End If
    Else
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=")
        End If
    End If
    
    strSql = strSql & " ORDER BY a.planned_ship_dt"
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getNonDropItemsOpen()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Order_item.cls.getNonDropItemsOpen"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
   strSql = "SELECT  " & _
                " a.sales_ord_id, " & _
                " a.requested_del_dt, " & _
                " a.planned_ship_dt, " & _
                " a.ordered_qty, " & _
                " a.item_desc, " & _
                " a.shipped_qty, " & _
                " a.type_cd as soi_type_cd, " & _
                " a.status_cd, " & _
                " b.item_nbr, " & _
                " b.item_id, " & _
                " c.type_cd, " & _
                " d.cust_nbr, " & _
                " d.cust_id, " & _
                " e.status_cd as mq_status_cd, " & _
                " e.ready_qty, " & _
                " e.back_ordered_qty " & _
                " FROM sales_order_item a " & _
                " inner join item b on b.item_id = a.item_id" & _
                " inner join sales_order c on c.sales_ord_id = a.sales_ord_id " & _
                " inner join customer d on d.cust_id = c.cust_id " & _
                " left join mat_req e on a.sales_ord_id = e.sales_ord_id and a.line_nbr = e.so_line_nbr " & _
                " where 1 = 1 "

    If fValuePresent(p_cust_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("d.cust_id", p_cust_id, "N", "=")
    End If
    
    If fValuePresent(p_status_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
    
    If fValuePresent(p_type_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("c.type_cd", p_type_cd, "S", "=")
    End If
    
    If fValuePresent(p_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.sales_ord_id", p_sales_ord_id, "S", "=")
    End If
    
    If fValuePresent(p_item_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("c.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
    
    If fValuePresent(p_start_dt) = True Then
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", ">=")
            strSql = strSql & " " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=") & ")"
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", ">=")
        End If
    Else
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=")
        End If
    End If
    
    strSql = strSql & " ORDER BY a.planned_ship_dt"
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub



