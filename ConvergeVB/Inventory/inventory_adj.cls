VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_adj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_lot As Object
Private mvaro_location As Object
Private mvaro_bin As Object
Private mvarp_inventory_adj_id As Variant
Private mvarp_item_id As Variant
Private mvarp_item_nbr As Variant
Private mvarp_lot_id As Variant
Private mvarp_inv_adj_id As Variant
Private mvarp_adj_qty As Variant
Private mvarp_ship_id As Variant
Private mvarp_ship_ret_id As Variant
Private mvarp_sales_ord_id As Variant
Private mvarp_so_line_nbr As Variant
Private mvarp_po_receipt_id As Variant
Private mvarp_po_return_id As Variant
Private mvarp_po_id As Variant
Private mvarp_po_line_nbr As Variant
Private mvarp_prod_run_id As Variant
Private mvarp_assem_ord_id As Variant
Private mvarp_ac_line_nbr As Variant
Private mvarp_mat_req_id As Variant
Private mvarp_bin_id As Variant
Private mvarp_new_bin_id As Variant
Private mvarp_phy_inv_id As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_on_hand_qty As Double
Private mvarp_allocated_qty As Double
Private mvarp_available_qty As Double
Private mvarp_move_inv_fl As Boolean
Private mvaro_record_set As ADODB.Recordset
Private mvarp_Comment_tx As Variant

'
Public Property Let p_comment_tx(ByVal inString As Variant)
    mvarp_Comment_tx = inString
End Property
Public Property Get p_comment_tx() As Variant
    p_comment_tx = mvarp_Comment_tx
End Property
Public Property Let p_move_inv_fl(ByVal vData As Variant)
    mvarp_move_inv_fl = vData
End Property
Public Property Get p_move_inv_fl() As Variant
    p_move_inv_fl = mvarp_move_inv_fl
End Property

Public Property Let p_on_hand_qty(ByVal vData As Variant)
    mvarp_on_hand_qty = vData
End Property
Public Property Get p_on_hand_qty() As Variant
    p_on_hand_qty = mvarp_on_hand_qty
End Property

Public Property Let p_allocated_qty(ByVal vData As Variant)
    mvarp_allocated_qty = vData
End Property
Public Property Get p_allocated_qty() As Variant
    p_allocated_qty = mvarp_allocated_qty
End Property

Public Property Let p_available_qty(ByVal vData As Variant)
    mvarp_available_qty = vData
End Property
Public Property Get p_available_qty() As Variant
    p_available_qty = mvarp_available_qty
End Property

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Let p_new_bin_id(ByVal vData As Variant)
    mvarp_new_bin_id = vData
End Property
Public Property Get p_new_bin_id() As Variant
    p_new_bin_id = mvarp_new_bin_id
End Property
Public Property Let p_mat_req_id(ByVal vData As Variant)
    mvarp_mat_req_id = vData
End Property
Public Property Get p_mat_req_id() As Variant
    p_mat_req_id = mvarp_mat_req_id
End Property


Public Property Let p_phy_inv_id(ByVal vData As Variant)
    mvarp_phy_inv_id = vData
End Property
Public Property Get p_phy_inv_id() As Variant
    p_phy_inv_id = mvarp_phy_inv_id
End Property

Public Property Let p_prod_run_id(ByVal vData As Variant)
    mvarp_prod_run_id = vData
End Property
Public Property Get p_prod_run_id() As Variant
    p_prod_run_id = mvarp_prod_run_id
End Property

Public Property Let p_ac_line_nbr(ByVal vData As Variant)
    mvarp_ac_line_nbr = vData
End Property
Public Property Get p_ac_line_nbr() As Variant
    p_ac_line_nbr = mvarp_ac_line_nbr
End Property

Public Property Let p_so_line_nbr(ByVal vData As Variant)
    mvarp_so_line_nbr = vData
End Property
Public Property Get p_so_line_nbr() As Variant
    p_so_line_nbr = mvarp_so_line_nbr
End Property
Public Property Let p_bin_id(ByVal vData As Variant)
    mvarp_bin_id = vData
End Property
Public Property Get p_bin_id() As Variant
    p_bin_id = mvarp_bin_id
End Property
Public Property Let p_po_return_id(ByVal vData As Variant)
    mvarp_po_return_id = vData
End Property
Public Property Get p_po_return_id() As Variant
    p_po_return_id = mvarp_po_return_id
End Property
Public Property Let p_assem_ord_id(ByVal vData As Variant)
    mvarp_assem_ord_id = vData
End Property
Public Property Get p_assem_ord_id() As Variant
    p_assem_ord_id = mvarp_assem_ord_id
End Property
Public Property Let p_po_receipt_id(ByVal vData As Variant)
    mvarp_po_receipt_id = vData
End Property
Public Property Get p_po_receipt_id() As Variant
    p_po_receipt_id = mvarp_po_receipt_id
End Property
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
    p_inv_loc_id = mvarp_inv_loc_id
End Property
Public Property Let p_sales_ord_id(ByVal vData As Variant)
    mvarp_sales_ord_id = vData
End Property
Public Property Get p_sales_ord_id() As Variant
    p_sales_ord_id = mvarp_sales_ord_id
End Property
Public Property Let p_ship_ret_id(ByVal vData As Variant)
    mvarp_ship_ret_id = vData
End Property
Public Property Get p_ship_ret_id() As Variant
    p_ship_ret_id = mvarp_ship_ret_id
End Property
Public Property Let p_ship_id(ByVal vData As Variant)
    mvarp_ship_id = vData
End Property
Public Property Get p_ship_id() As Variant
    p_ship_id = mvarp_ship_id
End Property
Public Property Let p_po_line_nbr(ByVal vData As Variant)
    mvarp_po_line_nbr = vData
End Property
Public Property Get p_po_line_nbr() As Variant
    p_po_line_nbr = mvarp_po_line_nbr
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
    p_po_id = mvarp_po_id
End Property
Public Property Let p_adj_qty(ByVal vData As Variant)
    mvarp_adj_qty = vData
End Property
Public Property Get p_adj_qty() As Variant
    p_adj_qty = mvarp_adj_qty
End Property
Public Property Let p_inv_adj_id(ByVal vData As Variant)
    mvarp_inv_adj_id = vData
End Property
Public Property Get p_inv_adj_id() As Variant
    p_inv_adj_id = mvarp_inv_adj_id
End Property
Public Property Let p_lot_id(ByVal vData As Variant)
    mvarp_lot_id = vData
End Property
Public Property Get p_lot_id() As Variant
    p_lot_id = mvarp_lot_id
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
Public Property Let p_inventory_adj_id(ByVal vData As Variant)
    mvarp_inventory_adj_id = vData
End Property
Public Property Get p_inventory_adj_id() As Variant
    p_inventory_adj_id = mvarp_inventory_adj_id
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
Public Property Set o_lot(ByVal vData As Object)
    Set mvaro_lot = vData
End Property
Public Property Get o_lot() As Object
    Set o_lot = mvaro_lot
End Property
Public Property Set o_bin(ByVal vData As Object)
    Set mvaro_bin = vData
End Property
Public Property Get o_bin() As Object
    Set o_bin = mvaro_bin
End Property
Public Property Set o_location(ByVal vData As Object)
    Set mvaro_location = vData
End Property
Public Property Get o_location() As Object
    Set o_location = mvaro_location
End Property

Public Sub GetInventoryAdj()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    strRoutine_nm = "InventoryAdj.cls.GetInventoryAdj"
    Dim objrecordset As ADODB.Recordset
  
    Call SelectInventoryAdj(p_inv_adj_id, objrecordset)
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
    Else
        If objrecordset.EOF = True Then
            p_inv_adj_id = Null
        Else
            p_inv_adj_id = objrecordset("inv_adj_id")
            p_item_id = objrecordset("item_id")
            p_mat_req_id = objrecordset("mat_req_id")
            p_lot_id = objrecordset("lot_id")
            p_bin_id = objrecordset("bin_id")
            p_inv_loc_id = objrecordset("inv_loc_id")
            p_adj_qty = objrecordset("adj_qty")
            p_ship_id = objrecordset("ship_id")
            p_ship_ret_id = objrecordset("ship_ret_id")
            p_sales_ord_id = objrecordset("sales_ord_id")
            p_so_line_nbr = objrecordset("so_line_nbr")
            p_po_receipt_id = objrecordset("po_receipt_id")
            p_po_return_id = objrecordset("po_return_id")
            p_po_id = objrecordset("po_id")
            p_po_line_nbr = objrecordset("po_line_nbr")
            p_assem_ord_id = objrecordset("assem_ord_id")
            p_ac_line_nbr = objrecordset("ac_line_nbr")
            p_phy_inv_id = objrecordset("phy_inv_id")
            p_prod_run_id = objrecordset("prod_run_id")
        End If
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

Public Sub Add()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objInventoryUtil As Object
    strRoutine_nm = "InventoryAdj.cls.Add"
    
    Call Validate("add")
    Call IssueInsertData
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    If IsNull(Me.p_comment_tx) = False Then
        If Len(Trim(Me.p_comment_tx)) > 0 Then
            Dim objComment As Object
            Set objComment = CreateObject("converge.comment")
            Set objComment.o_error = Me.o_error
            Set objComment.o_trans_data = Me.o_trans_data
            objComment.p_comment_desc = "Inventory Adjustment"
            objComment.p_comment_tx = Me.p_comment_tx
            objComment.p_obj_nm = "phy_inv_adj"
            objComment.p_obj_id = Me.p_inv_adj_id
            objComment.addComment
        End If
    End If
    
    If p_mat_req_id > 0 Then
        Dim objInventoryAlloc As Object
        Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
        Set objInventoryAlloc.o_error = Me.o_error
        Set objInventoryAlloc.o_trans_data = Me.o_trans_data
        objInventoryAlloc.p_mat_req_id = p_mat_req_id
        objInventoryAlloc.p_inv_loc_id = p_inv_loc_id
        objInventoryAlloc.p_adj_qty = p_adj_qty
        Call objInventoryAlloc.AdjustAllocation
    End If
    
    If p_ship_id > 0 Then
        Dim objShipmentItem As Object
        Set objShipmentItem = CreateObject("converge_sales.shipment_item")
        Set objShipmentItem.o_trans_data = Me.o_trans_data
        Set objShipmentItem.o_error = Me.o_error
        objShipmentItem.p_ship_id = p_ship_id
        objShipmentItem.p_sales_ord_id = p_sales_ord_id
        objShipmentItem.p_line_nbr = p_so_line_nbr
        Call objShipmentItem.UpdateQuantities(o_location.p_consignment_fl)
    End If
    
    If p_ship_ret_id > 0 Then
        Dim objSalesOrderItem As Object
        Set objSalesOrderItem = CreateObject("converge_sales.Sales_Order_item")
        Set objSalesOrderItem.o_trans_data = Me.o_trans_data
        Set objSalesOrderItem.o_error = Me.o_error
        objSalesOrderItem.p_sales_ord_id = Me.p_sales_ord_id
        objSalesOrderItem.p_line_nbr = Me.p_so_line_nbr
        Call objSalesOrderItem.UpdateQuantities( _
                                  Me.o_error _
                                , Me.o_trans_data _
                                , Me.p_sales_ord_id _
                                , Me.p_so_line_nbr _
                                )
    End If
    
    If p_po_receipt_id > 0 Or p_po_return_id > 0 Then
        Dim objPurchaseOrderItem As Object
        Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
        Set objPurchaseOrderItem.o_error = Me.o_error
        Set objPurchaseOrderItem.o_trans_data = Me.o_trans_data
        objPurchaseOrderItem.p_po_id = p_po_id
        objPurchaseOrderItem.p_line_nbr = p_po_line_nbr
        Call objPurchaseOrderItem.UpdateQuantities( _
                                  Me.o_error _
                                , Me.o_trans_data _
                                , Me.p_po_id _
                                , Me.p_po_line_nbr _
                                )
        Set objPurchaseOrderItem = Nothing
    
        Set objInventoryUtil = CreateObject("converge_inv.inventory_util")
        Set objInventoryUtil.o_error = Me.o_error
        Set objInventoryUtil.o_trans_data = Me.o_trans_data
        If p_adj_qty > 0 Then
            '   if the po receipt increased inventory, reallocate on order inventory for the item
            '   note: this already happens during the poi update quantities method - matt hetterich 11/7/00
            'objInventoryUtil.p_item_id = p_item_id
            'Call objInventoryUtil.ReallocateItemOnOrdInv
        Else
            ' if the po receipt decreased inventory, reallocate inventory and on order inventory for the item
            objInventoryUtil.p_item_id = p_item_id
            Call objInventoryUtil.ReallocateItem
        End If
        Set objInventoryUtil = Nothing
    End If
    
    ' Production Run
    If p_prod_run_id > 0 And Len(p_ac_line_nbr) = 0 Then
        Dim objProductionRun As Object
        Set objProductionRun = CreateObject("converge_prod.prod_run")
        Call objProductionRun.UpdateQuantities( _
                                  Me.o_error _
                                , Me.o_trans_data _
                                , Me.p_prod_run_id _
                                )
        Set objProductionRun = Nothing
    
        Set objInventoryUtil = CreateObject("converge_inv.inventory_util")
        Set objInventoryUtil.o_error = Me.o_error
        Set objInventoryUtil.o_trans_data = Me.o_trans_data
        If p_adj_qty > 0 Then
                    'Do not call reallocateitemonordinv.  This
                    'already is being called when we call prod run UpdateQuantities
                    'and prodrun.updatequantities call assemblyOrd.updatequantities
            'objInventoryUtil.p_item_id = p_item_id
            'Call objInventoryUtil.ReallocateItemOnOrdInv
        Else
            ' if the po receipt decreased inventory, reallocate inventory and on order inventory for the item
            objInventoryUtil.p_item_id = p_item_id
            Call objInventoryUtil.ReallocateItem
        End If
        Set objInventoryUtil = Nothing
    End If
    
    If p_phy_inv_id > 0 Then
        If p_move_inv_fl = True Then
        Else
            Set objInventoryUtil = CreateObject("converge_inv.inventory_util")
            Set objInventoryUtil.o_error = Me.o_error
            Set objInventoryUtil.o_trans_data = Me.o_trans_data
            Set objInventoryUtil.o_error = Me.o_error
            Set objInventoryUtil.o_trans_data = Me.o_trans_data
            objInventoryUtil.p_item_id = p_item_id
            Call objInventoryUtil.ReallocateItem
        End If
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


Public Sub Move()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objLot As Object
    Dim varFromBinId As Variant
    Dim varToBinId As Variant
    Dim varFromAdjQty As Variant
    Dim varToAdjQty As Variant
    
    strRoutine_nm = "InventoryAdj.cls.Move"
    
    p_move_inv_fl = True
    varFromBinId = p_bin_id
    varToBinId = p_new_bin_id
    varFromAdjQty = p_adj_qty * -1
    varToAdjQty = p_adj_qty
        
    p_adj_qty = varFromAdjQty
    p_bin_id = varFromBinId
    Call Add
    
    p_adj_qty = varToAdjQty
    p_bin_id = varToBinId
    Call Add
    p_move_inv_fl = False
    
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

Public Sub GetQty()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_adj.cls.GetQty"

    p_on_hand_qty = OnHandQty(p_inv_loc_id, p_lot_id, p_bin_id, Me.o_error, Me.o_trans_data)
    
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

Private Sub Validate(ByVal v_strFunction_cd As Variant)
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim rstShipmentItem As ADODB.Recordset
    Dim objrecordset As ADODB.Recordset
    Dim objShipmentItem As Object
    Dim obj2ndLegRs As ADODB.Recordset
    
    strRoutine_nm = "inventory_adj.cls.validate"

    'Set objShipmentItem = New converge_sales.shipment_item
    Set objShipmentItem = CreateObject("converge_sales.shipment_item")
    
    If v_strFunction_cd = "add" Then
        If p_adj_qty = 0 Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "2100"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2203
            Exit Sub
        End If
        
        If IsNull(p_item_id) Then
            p_item_id = fGetItemId(p_item_nbr, Me.o_error, Me.o_trans_data)
        Else
            p_item_nbr = fGetItemNbr(p_item_id, Me.o_error, Me.o_trans_data)
        End If
        
        If IsNull(p_item_id) Or IsNull(p_item_nbr) Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "2110"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1045
            Exit Sub
        End If
        
        'inventory adjustment must be related to one of the following objects
        If p_prod_run_id > 0 Or _
           p_ship_id > 0 Or _
           p_ship_ret_id > 0 Or _
           p_po_receipt_id > 0 Or _
           p_po_return_id > 0 Or _
           p_phy_inv_id Then
        Else
            Me.o_error.p_type_cd = "F"
            Me.o_error.p_err_cd = "2120"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2204
            Exit Sub
        End If
        
        Call validate_location
        Call validate_bin
        Call validate_lot
        
        If p_ship_id > 0 Then
            If p_mat_req_id > 0 Or o_location.p_consignment_fl <> "N" Then
            Else
                Me.o_error.p_type_cd = "F"
                Me.o_error.p_err_cd = "2130"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_desc = "Material requirement must be provided with shipment id"
                Me.o_error.p_message_id = 2205
                Exit Sub
            End If
            
            objShipmentItem.p_ship_id = p_ship_id
            objShipmentItem.p_sales_ord_id = p_sales_ord_id
            objShipmentItem.p_line_nbr = p_so_line_nbr
            Set objShipmentItem.o_error = Me.o_error
            Set objShipmentItem.o_trans_data = Me.o_trans_data
            objShipmentItem.getShipmentItem
            
            ' check to see if the loaded quantity on the 2nd leg shipment is > the loaded quantity on the 1st leg
            If objShipmentItem.p_ship_type_cd = "2L" Then
                strSql = " SELECT SUM(loaded_qty) as tot_loaded " & _
                         " FROM shipment_item " & _
                         " WHERE " & _
                         " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=") & _
                         " AND " & fWhereVariable("line_nbr", p_so_line_nbr, "N", "=") & _
                         " AND " & fWhereVariable("ship_type_cd", "1L", "S", "=")
                
                Call Me.o_trans_data.OpenRecordset(rstShipmentItem, strSql, Me.o_error)
                
                rstShipmentItem.MoveFirst
                                
                strSql = " SELECT SUM(loaded_qty) as tot_loaded " & _
                         " FROM shipment_item " & _
                         " WHERE " & _
                         " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=") & _
                         " AND " & fWhereVariable("line_nbr", p_so_line_nbr, "N", "=") & _
                         " AND " & fWhereVariable("ship_type_cd", "2L", "S", "=")
                
                Call Me.o_trans_data.OpenRecordset(obj2ndLegRs, strSql, Me.o_error)
                obj2ndLegRs.MoveFirst
                
                If obj2ndLegRs("tot_loaded") - CDbl(p_adj_qty) > CDbl(rstShipmentItem("tot_loaded")) Then
                    Me.o_error.p_type_cd = "E"
                    Me.o_error.p_err_cd = "2131"
                    Me.o_error.p_routine_nm = strRoutine_nm
                    Me.o_error.p_message_id = 3426
                    Exit Sub
                End If
                
                rstShipmentItem.Close
                Set rstShipmentItem = Nothing
            End If
            
            If objShipmentItem.p_ship_type_cd = "1L" And p_adj_qty > 0 Then
                strSql = " SELECT SUM(loaded_qty) as tot_loaded " & _
                         " FROM shipment_item " & _
                         " WHERE " & _
                         " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=") & _
                         " AND " & fWhereVariable("line_nbr", p_so_line_nbr, "N", "=") & _
                         " AND " & fWhereVariable("ship_type_cd", "1L", "S", "=")
                
                Call Me.o_trans_data.OpenRecordset(rstShipmentItem, strSql, Me.o_error)
                
                rstShipmentItem.MoveFirst
                                
                strSql = " SELECT SUM(loaded_qty) as tot_loaded " & _
                         " FROM shipment_item " & _
                         " WHERE " & _
                         " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=") & _
                         " AND " & fWhereVariable("line_nbr", p_so_line_nbr, "N", "=") & _
                         " AND " & fWhereVariable("ship_type_cd", "2L", "S", "=")
                
                Call Me.o_trans_data.OpenRecordset(obj2ndLegRs, strSql, Me.o_error)
                obj2ndLegRs.MoveFirst
                
                    'if a first leg shipment is being unloaded, you can't make the 1st leg shipped qty
                    'less than the 2nd leg shipped qty.
                If CDbl(rstShipmentItem("tot_loaded")) - CDbl(p_adj_qty) < obj2ndLegRs("tot_loaded") Then
                    Me.o_error.p_type_cd = "E"
                    Me.o_error.p_err_cd = "2132"
                    Me.o_error.p_routine_nm = strRoutine_nm
                    Me.o_error.p_message_id = 3425
                    Exit Sub
                End If
                
                rstShipmentItem.Close
                Set rstShipmentItem = Nothing
            End If
        End If
            
        If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
            Exit Sub
        End If
        
        'if inventory is being releaved, insure quatity is on hand
        If p_adj_qty < 0 Then
            p_on_hand_qty = OnHandQty(p_inv_loc_id, p_lot_id, p_bin_id, Me.o_error, Me.o_trans_data)
            If p_on_hand_qty < p_adj_qty * -1 Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "2700"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 2206
                Exit Sub
            End If
        End If
        
        'if inventory is being loaded onto a shipment, insure
        'adj qty is available to the sales order.
        If p_adj_qty < 0 And (p_ship_id > 0 Or (p_prod_run_id > 0 And Len(Me.p_ac_line_nbr) > 0)) Then
            p_allocated_qty = AllocatedQty(p_inv_loc_id, p_item_id, p_mat_req_id, Me.o_error, Me.o_trans_data)
            p_available_qty = AvailableQty(p_inv_loc_id, p_item_id, Me.o_error, Me.o_trans_data)
            If p_allocated_qty + p_available_qty < p_adj_qty * -1 Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "2800"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 2207
                Exit Sub
            End If
        End If
        
        'if a shipment is being unloaded, make sure that they are not
        'unloading more than was loaded.
        
        If p_adj_qty > 0 And p_ship_id > 0 Then
            If fLoadedQty(p_ship_id, p_sales_ord_id, p_so_line_nbr) < p_adj_qty Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "2900"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 2208
                Exit Sub
            End If
        End If
    
        If p_adj_qty < 0 And p_po_receipt_id > 0 Then
            If 0 > (fReceivedQty(p_po_receipt_id, p_po_id, p_po_line_nbr, p_lot_id) + p_adj_qty) Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "3000"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 2209
                Exit Sub
            End If
        End If
        
        If p_adj_qty > 0 And p_prod_run_id > 0 And Len(p_ac_line_nbr) > 0 Then
            If 0 < (fReleavedQty(p_prod_run_id, p_assem_ord_id, p_ac_line_nbr, p_lot_id, p_inv_loc_id) + p_adj_qty) Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "3000"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 0
                Exit Sub
            End If
        End If
        
        If p_adj_qty < 0 And p_prod_run_id > 0 And Len(p_ac_line_nbr) = 0 Then
            If 0 > fReceivedQty2(p_prod_run_id, p_assem_ord_id, p_lot_id, p_inv_loc_id) < p_adj_qty Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "3000"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 0
                Exit Sub
            End If
        End If
            
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


Private Sub validate_location()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_adj.cls.validate_location"
    
    Set o_location = CreateObject("converge_inv.inventory_loc")
    Set o_location.o_trans_data = Me.o_trans_data
    Set o_location.o_error = Me.o_error
    o_location.p_inv_loc_id = p_inv_loc_id
    o_location.GetInventoryLoc
    
    If IsNull(o_location.p_inv_loc_id) = True Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "2300"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 1175
        Exit Sub
    End If
    
    If o_location.p_consignment_fl = "Y" Then
        If Len(Me.p_ship_id) < 1 Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "2301"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 0
            Exit Sub
        End If
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


Private Sub validate_bin()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_adj.cls.validate_bin"
    
    ' validate bin.  if inventory location is a consignment location,
    ' a valid bin is not required
    If o_location.p_consignment_fl = "Y" Then
        Exit Sub
    End If
    
    Set o_bin = CreateObject("converge_inv.inventory_bin")
    Set o_bin.o_trans_data = Me.o_trans_data
    Set o_bin.o_error = Me.o_error
    o_bin.p_bin_id = p_bin_id
    o_bin.GetInventoryBin
    
    If IsNull(o_bin.p_bin_id) = True Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "2500"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2202
        Exit Sub
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


Private Sub validate_lot()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_adj.cls.validate_lot"
    
    Set o_lot = CreateObject("converge_inv.lot")
    Set o_lot.o_trans_data = Me.o_trans_data
    Set o_lot.o_error = Me.o_error
    o_lot.p_lot_id = p_lot_id
    o_lot.GetLot
    
    If IsNull(o_lot.p_lot_id) = True Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "2300"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2151
        Exit Sub
    End If
    
    If CInt(o_lot.p_item_id) <> CInt(p_item_id) Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "2301"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2151
        Exit Sub
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


Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "inventory_adj.cls.IssueInsertData"
    
    p_inv_adj_id = fGetId("inv_adj_id", Me.o_error, Me.o_trans_data)
    
    strSql = "insert into inventory_adj ( "
    strSql = strSql & "  inv_adj_id "
    strSql = strSql & ", item_id "
    strSql = strSql & ", lot_id "
    strSql = strSql & ", bin_id "
    strSql = strSql & ", inv_loc_id "
    strSql = strSql & ", mat_req_id "
    strSql = strSql & ", adj_qty "
    strSql = strSql & ", ship_id "
    strSql = strSql & ", ship_ret_id "
    strSql = strSql & ", sales_ord_id "
    strSql = strSql & ", so_line_nbr "
    strSql = strSql & ", po_receipt_id "
    strSql = strSql & ", po_return_id "
    strSql = strSql & ", po_id "
    strSql = strSql & ", po_line_nbr "
    strSql = strSql & ", assem_ord_id "
    strSql = strSql & ", ac_line_nbr "
    strSql = strSql & ", phy_inv_id "
    strSql = strSql & ", prod_run_id "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_inv_adj_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_lot_id, "N")
    strSql = strSql & "," & fInsertVariable(p_bin_id, "S")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_id, "S")
    strSql = strSql & "," & fInsertVariable(p_mat_req_id, "N")
    strSql = strSql & "," & fInsertVariable(p_adj_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_ship_id, "N")
    strSql = strSql & "," & fInsertVariable(p_ship_ret_id, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_so_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_po_receipt_id, "N")
    strSql = strSql & "," & fInsertVariable(p_po_return_id, "N")
    strSql = strSql & "," & fInsertVariable(p_po_id, "N")
    strSql = strSql & "," & fInsertVariable(p_po_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_assem_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_ac_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_phy_inv_id, "N")
    strSql = strSql & "," & fInsertVariable(p_prod_run_id, "N")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    
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


Private Sub SelectInventoryAdj(ByVal varInv_adj_id As Variant, ByRef objrecordset As ADODB.Recordset)
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm
    Dim strSql As String
    strRoutine_nm = "inventory_adj.cls.SelectInventory"

    strSql = "SELECT * " & _
             "FROM inventory_adj " & _
            "WHERE " & fWhereVariable("inv_adj_id", p_inv_adj_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to get the available inventory
'  Arguments:   v_varInv_loc_id
'               v_varItem_id
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Public Function AvailableQty(ByVal v_varInv_loc_id As Variant, ByVal v_varItem_id As Variant, ByRef r_objError As Object, ByRef r_objTrans_data As Object) As Double
    
    On Error GoTo error_handler
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTrans_data
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
        
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.AvailableQty"
    
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    Dim dblInventoryQty As Double
    Dim dblAllocatedQty As Double
    
    strSql = "SELECT sum(a.adj_qty) AS [p_adj_qty] " & _
            "FROM inventory_adj a, lot b " & _
            "WHERE " & _
            "    " & fWhereVariable("a.item_id", v_varItem_id, "N", "=") & _
            "and " & fWhereVariable("a.inv_loc_id", v_varInv_loc_id, "S", "=") & _
            "and a.lot_id = b.lot_id " & _
            "and b.status_cd = 'AV' "

    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_adj_qty")) = True Then
        dblInventoryQty = 0
    Else
        dblInventoryQty = objrecordset("p_adj_qty")
    End If
    
    objrecordset.Close
    
    strSql = "SELECT sum(alloc_qty) as [p_alloc_qty] " & _
            "FROM inventory_alloc " & _
            "WHERE " & _
            "    " & fWhereVariable("item_id", v_varItem_id, "N", "=") & _
            " and " & fWhereVariable("inv_loc_id", v_varInv_loc_id, "S", "=")

    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_alloc_qty")) = True Then
        dblAllocatedQty = 0
    Else
        dblAllocatedQty = objrecordset(0)
    End If
    
    AvailableQty = dblInventoryQty - dblAllocatedQty
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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


Public Function AllocatedQty(ByVal v_varInv_loc_id As Variant, ByVal v_varItem_id As Variant, ByVal v_varMat_req_id As Variant, ByRef r_objError As Object, ByRef r_objTrans_data As Object) As Double
    
    On Error GoTo error_handler
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTrans_data
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
        
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.AllocatedQty"
    
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    
    strSql = "SELECT sum(alloc_qty) as [p_qty] " & _
           "FROM inventory_alloc " & _
           "WHERE " & _
           "    " & fWhereVariable("mat_req_id", v_varMat_req_id, "N", "=") & _
           "and " & fWhereVariable("inv_loc_id", v_varInv_loc_id, "S", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_qty")) = True Then
        AllocatedQty = 0
    Else
        AllocatedQty = objrecordset("p_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Public Function OnHandQty(ByVal v_varInv_loc_id As Variant, ByVal v_varLot_id As Variant, ByVal v_varBin_id As Variant, ByRef r_objError As Object, ByRef r_objTrans_data As Object) As Double
    
    On Error GoTo error_handler
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTrans_data
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
        
    Dim strRoutine_nm
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "inventory_adj.cls.OhHandQty"
    
    strSql = "SELECT sum(adj_qty) as [p_qty] " & _
           "FROM inventory_adj " & _
           "WHERE " & _
           "    " & fWhereVariable("bin_id", v_varBin_id, "S", "=") & _
           "and " & fWhereVariable("inv_loc_id", v_varInv_loc_id, "S", "=") & _
           "and " & fWhereVariable("lot_id", v_varLot_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_qty")) = True Then
        OnHandQty = 0
    Else
        OnHandQty = objrecordset("p_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Private Function fLoadedQty(ByVal varShip_id As Variant, ByVal varSales_ord_id As Variant, ByVal varLine_nbr As Variant) As Double
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.fLoadedQty"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT loaded_qty      "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("ship_id", varShip_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("sales_ord_id", varSales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", varLine_nbr, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If objrecordset.EOF = True Or IsNull(objrecordset("loaded_qty")) = True Then
        fLoadedQty = 0
    Else
        fLoadedQty = objrecordset("loaded_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Private Function fReleavedQty(ByVal varProd_run_id As Variant, ByVal varAssem_ord_id, ByVal varLine_nbr As Variant, ByVal varLot_id As Variant, ByVal varInv_loc_id As Variant) As Double
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.fReleavedQty"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT sum(adj_qty) as releaved_qty     "
    strSql = strSql & " FROM inventory_adj "
    strSql = strSql & " WHERE         "
    strSql = strSql & "     " & fWhereVariable("prod_run_id", varProd_run_id, "N", "=")
    strSql = strSql & " and " & fWhereVariable("assem_ord_id", varAssem_ord_id, "N", "=")
    strSql = strSql & " and " & fWhereVariable("ac_line_nbr", varLine_nbr, "N", "=")
    strSql = strSql & " and " & fWhereVariable("lot_id", varLot_id, "N", "=")
    strSql = strSql & " and " & fWhereVariable("inv_loc_id", varInv_loc_id, "S", "=")
    
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If objrecordset.EOF = True Or IsNull(objrecordset("releaved_qty")) = True Then
        fReleavedQty = 0
    Else
        fReleavedQty = objrecordset("releaved_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Private Function fReceivedQty(ByVal v_varPo_receipt_id As Variant, ByVal v_varPo_id As Variant, ByVal v_varLine_nbr As Variant, ByVal v_varLot_id As Variant) As Double
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.fReceivedQty"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT sum(adj_qty) as [p_received_qty] "
    strSql = strSql & "FROM inventory_adj "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("po_receipt_id", v_varPo_receipt_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("po_id", v_varPo_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("po_line_nbr", v_varLine_nbr, "N", "=")
    strSql = strSql & "and" & fWhereVariable("lot_id", v_varLot_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_received_qty")) = True Then
        fReceivedQty = 0
    Else
        fReceivedQty = objrecordset("p_received_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Private Function fReceivedQty2(ByVal v_varProd_run_id As Variant, ByVal v_varAssem_ord_id As Variant, ByVal v_varLot_id As Variant, v_varInv_loc_id As Variant) As Double
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.fReceivedQty2"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT sum(adj_qty) as [p_received_qty] "
    strSql = strSql & "FROM inventory_adj "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("prod_run_id", v_varProd_run_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("assem_ord_id", v_varAssem_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("lot_id", v_varLot_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("inv_loc_id", v_varInv_loc_id, "S", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_received_qty")) = True Then
        fReceivedQty2 = 0
    Else
        fReceivedQty2 = objrecordset("p_received_qty")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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

Public Sub GetInventoryItems()

    On Error GoTo error_handler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim lngCount As Long
    Dim strItemNbr As String
    Dim objrecordset As ADODB.Recordset
    Dim rstInventoryAlloc As ADODB.Recordset
    Dim rstOnOrdInvAdj As ADODB.Recordset
    Dim rstOnOrdInvAlloc As ADODB.Recordset
    Dim rstInventoryAdj As ADODB.Recordset
   
    strRoutine_nm = "inventory_adj.cls.getInventoryItems"

    strSql = "SELECT " & _
            "  a.item_id " & _
            ",  a.item_id as p_alloc_qty" & _
            ",  a.item_id as p_on_ord_qty" & _
            ",  a.item_id as p_on_ord_alloc_qty" & _
            ", a.item_nbr " & _
            ", b.inv_loc_id as p_adj_loc_id " & _
            ", sum(b.adj_qty) as p_adj_qty " & _
            "FROM item a  " & _
             "left outer JOIN inventory_adj b ON a.item_id = b.item_id  "
    If Len((p_inv_loc_id)) > 0 Then
       strSql = strSql & "  AND " & fWhereVariable("b.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
      
      strSql = strSql & "where 1 = 1 "
    
    If IsNull(p_item_nbr) = True Or p_item_nbr = "" Then
    Else
       lngCount = InStr(1, p_item_nbr, "*", vbBinaryCompare)
       If lngCount > 0 Then
            strItemNbr = Left(p_item_nbr, lngCount - 1)
            strSql = strSql & " AND a.item_nbr LIKE '" & strItemNbr & "%'"
       Else
            strSql = strSql & "  AND " & fWhereVariable("a.item_nbr", p_item_nbr, "S", "=")
       End If
    End If
    
    If IsNull(p_bin_id) = True Or p_bin_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("b.bin_id", p_bin_id, "S", "=")
    End If
    
  
    
    strSql = strSql & " GROUP BY a.item_nbr, a.item_id, b.inv_loc_id "

    Call Me.o_trans_data.OpenRecordset(rstInventoryAdj, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Set objrecordset = New ADODB.Recordset

    'objRecordset.CursorType = adOpenStatic
    'objRecordset.PageSize = 1
    
    objrecordset.Fields.Append "p_adj_qty", adInteger, 100, adFldUpdatable
    objrecordset.Fields.Append "item_id", adInteger, 100, adFldUpdatable
    objrecordset.Fields.Append "p_adj_loc_id", adVarChar, 100, adFldUpdatable
    objrecordset.Fields.Append "item_nbr", adVarChar, 100, adFldUpdatable
    objrecordset.Fields.Append "p_alloc_qty", adInteger, 100, adFldUpdatable
    objrecordset.Fields.Append "p_on_ord_qty", adInteger, 100, adFldUpdatable
    objrecordset.Fields.Append "p_on_ord_alloc_qty", adInteger, 100, adFldUpdatable
    objrecordset.Open
    
    rstInventoryAdj.MoveFirst
    
    If Len(rstInventoryAdj!p_adj_loc_id) = 0 Then
        objrecordset.AddNew
        If Len(p_inv_loc_id) > 0 Then
           objrecordset("p_adj_loc_id") = p_inv_loc_id
        Else
           objrecordset("p_adj_loc_id") = "TIMCO"
        End If
        objrecordset("p_adj_qty") = 0
        objrecordset("p_alloc_qty") = 0
        objrecordset("p_on_ord_alloc_qty") = 0
        objrecordset("p_on_ord_qty") = 0
        objrecordset("item_id") = rstInventoryAdj("item_id")
        objrecordset("item_nbr") = rstInventoryAdj("item_nbr")
        Set o_record_set = objrecordset
        Exit Sub
    End If
    
    Do Until rstInventoryAdj.EOF = True
    
        objrecordset.AddNew
        If IsNull(rstInventoryAdj("p_adj_loc_id")) Then
            If Len(p_inv_loc_id) > 0 Then
               objrecordset("p_adj_loc_id") = p_inv_loc_id
            Else
               objrecordset("p_adj_loc_id") = "TIMCO"
            End If
        Else
            objrecordset("p_adj_loc_id") = rstInventoryAdj("p_adj_loc_id")
        End If
        
        If IsNull(rstInventoryAdj("p_adj_qty")) Then
            objrecordset("p_adj_qty") = 0
        Else
            objrecordset("p_adj_qty") = rstInventoryAdj("p_adj_qty")
        End If
        
        If IsNull(rstInventoryAdj("item_id")) Then
        Else
            objrecordset("item_id") = rstInventoryAdj("item_id")
        End If
        
        If IsNull(rstInventoryAdj("item_nbr")) Then
        Else
            objrecordset("item_nbr") = rstInventoryAdj("item_nbr")
        End If
        
        If Len(objrecordset("p_adj_loc_id")) = 0 Then
            objrecordset("p_adj_loc_id") = "Timco"
        End If
        
        strSql = "SELECT sum(alloc_qty) as p_alloc_qty " & _
                  " FROM inventory_alloc " & _
                "  WHERE " & fWhereVariable("inv_loc_id", objrecordset("p_adj_loc_id"), "S", "=") & _
                "    AND " & fWhereVariable("item_id", objrecordset("item_id"), "N", "=")
        Call Me.o_trans_data.OpenRecordset(rstInventoryAlloc, strSql, Me.o_error)
        
        If Len(rstInventoryAlloc("p_alloc_qty")) > 0 Then
          objrecordset("p_alloc_qty") = rstInventoryAlloc("p_alloc_qty")
        Else
          objrecordset("p_alloc_qty") = 0
        End If
        
        
        strSql = "SELECT sum(alloc_qty) as p_on_ord_alloc_qty " & _
                  " FROM on_ord_inv_alloc " & _
                "  WHERE " & fWhereVariable("inv_loc_id", objrecordset("p_adj_loc_id"), "S", "=") & _
                "    AND " & fWhereVariable("item_id", objrecordset("item_id"), "N", "=")
        Call Me.o_trans_data.OpenRecordset(rstOnOrdInvAlloc, strSql, Me.o_error)
        If Len(rstOnOrdInvAlloc("p_on_ord_alloc_qty")) > 0 Then
            objrecordset("p_on_ord_alloc_qty") = rstOnOrdInvAlloc("p_on_ord_alloc_qty")
        Else
            objrecordset("p_on_ord_alloc_qty") = 0
        End If
            
        strSql = "SELECT sum(on_ord_qty) as p_on_ord_qty " & _
                  " FROM on_ord_inv_adj " & _
                "  WHERE " & fWhereVariable("inv_loc_id", objrecordset("p_adj_loc_id"), "S", "=") & _
                "    AND " & fWhereVariable("item_id", objrecordset("item_id"), "N", "=")
        Call Me.o_trans_data.OpenRecordset(rstOnOrdInvAdj, strSql, Me.o_error)
        If Len(rstOnOrdInvAdj("p_on_ord_qty")) > 0 Then
            objrecordset("p_on_ord_qty") = rstOnOrdInvAdj("p_on_ord_qty")
        Else
            objrecordset("p_on_ord_qty") = 0
        End If
        
        rstInventoryAdj.MoveNext
        
    Loop
    
    Set rstInventoryAlloc = Nothing
    Set rstOnOrdInvAdj = Nothing
    Set rstOnOrdInvAlloc = Nothing
    Set rstInventoryAdj = Nothing
    
    objrecordset.MoveFirst
    Set o_record_set = objrecordset
    
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


Public Sub GetPickList()
    
    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.GetPickList"
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    strSql = "SELECT " & _
             " a.ship_id, " & _
             " a.item_id, " & _
             " a.planned_qty, " & _
             " a.loaded_qty, " & _
             " b.lot_id, " & _
             " b.bin_id, " & _
             " sum(b.adj_qty) AS adj_qty " & _
             " FROM shipment_item a left join inventory_adj b " & _
             " on a.item_id = b.item_id "
    strSql = strSql & " where " & fWhereVariable("a.ship_id", p_ship_id, "N", "=")
    strSql = strSql & "  AND " & fWhereVariable("b.inv_loc_id", p_inv_loc_id, "S", "=")
    strSql = strSql & " GROUP BY " & _
            " a.ship_id, " & _
            " a.item_id, " & _
            " a.planned_qty, " & _
            " a.loaded_qty, " & _
            " b.lot_id, " & _
            " b.bin_id " & _
            " having sum(b.adj_qty) > 0 "
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If

    Set o_record_set = objrecordset

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

Private Function fSumAdjQty(ByVal v_varLot_id As Variant) As Long

    On Error GoTo error_handler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
        Exit Function
    End If
    
    Dim strRoutine_nm
    strRoutine_nm = "inventory_adj.cls.GetSumAdjQty"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT sum(adj_qty) as p_adj_qty"
    strSql = strSql & "FROM inventory_adj "
    strSql = strSql & "WHERE         "
    strSql = strSql & "and" & fWhereVariable("lot_id", v_varLot_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objrecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_adj_qty")) = True Then
        fSumAdjQty = 0
    Else
        fSumAdjQty = objrecordset("p_adj_qty")
    End If

    objrecordset.Close
    Set objrecordset = Nothing

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







