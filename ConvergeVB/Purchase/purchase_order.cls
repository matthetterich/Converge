VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "purchase_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_po_id As Variant
Private mvarp_backorder_cd As Variant
Private mvarp_backorder_desc As Variant
Private mvarp_buyer_assoc_id As Variant
Private mvarp_buyer_assoc_nbr As Variant
Private mvarp_carrier_nm As Variant
Private mvarp_confirmation_nbr As Variant
Private mvarp_currency_cd As Variant
Private mvarp_currency_desc As Variant
Private mvarp_drop_ship_fl As Variant
Private mvarp_drop_ship_desc As Variant
Private mvarp_fob_terms_cd As Variant
Private mvarp_fob_terms_desc As Variant
Private mvarp_freight_cd As Variant
Private mvarp_gross_amnt As Variant
Private mvarp_line_item_qty As Variant
Private mvarp_net_amnt As Variant
Private mvarp_opened_dt As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_revision_nbr As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_supplier_id As Variant
Private mvarp_notif_meth_cd As Variant
Private mvarp_notif_meth_desc As Variant
Private mvarp_supplier_nbr As Variant
Private mvarp_transportation_cd As Variant
Private mvarp_transportation_desc As Variant
Private mvarp_type_cd As Variant
Private mvarp_type_desc As Variant
Private mvarp_move_to_hist_fl As Variant
Private mvarp_order_text As Variant
Private mvaro_error As Object
Private mvaro_record_set As ADODB.Recordset
Private mvaro_trans_data As Object
Private mvaro_old_purchase_order As Object
Private mvarp_comm_meth_cd As Variant
Private mvarp_comm_meth_desc As Variant
Private mvarp_freight_desc As Variant
Private mvaro_stnd_address As Object
Private mvaro_stnd_name As Object
Private mvaro_ship_address As Object
Private mvaro_ship_name As Object
Private mvaro_bill_address As Object
Private mvaro_bill_name As Object
Private mvarp_sales_tax_pct As Variant
Private mvarp_sales_tax_amnt As Variant
Private mvarp_drop_ship_sales_ord_id As Variant
Private mvarp_bypass_nm_addr_fl As Variant
Private mvarp_bypass_desc_fl As Variant
Private mvarp_entered_dt As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_start_dt As Variant
Private mvarp_end_dt As Variant
Private mvarp_status_cd2 As Variant
Private mvarp_status_cd3 As Variant
Private mvarp_status_cd4 As Variant
Private mvarp_conv_from_ord_id As Variant
'

Public Property Let p_conv_from_ord_id(ByVal vData As Variant)
    mvarp_conv_from_ord_id = vData
End Property
Public Property Get p_conv_from_ord_id() As Variant
        p_conv_from_ord_id = mvarp_conv_from_ord_id
End Property


Public Property Let p_move_to_hist_fl(ByVal vData As Variant)
    mvarp_move_to_hist_fl = vData
End Property
Public Property Get p_move_to_hist_fl() As Variant
        p_move_to_hist_fl = mvarp_move_to_hist_fl
End Property
Public Property Let p_status_cd2(ByVal vData As Variant)
    mvarp_status_cd2 = vData
End Property
Public Property Get p_status_cd2() As Variant
        p_status_cd2 = mvarp_status_cd2
End Property
Public Property Let p_status_cd3(ByVal vData As Variant)
    mvarp_status_cd3 = vData
End Property
Public Property Get p_status_cd3() As Variant
        p_status_cd3 = mvarp_status_cd3
End Property
Public Property Let p_status_cd4(ByVal vData As Variant)
    mvarp_status_cd4 = vData
End Property
Public Property Get p_status_cd4() As Variant
        p_status_cd4 = mvarp_status_cd4
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

Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
End Property

Public Property Let p_entered_dt(ByVal vData As Variant)
    mvarp_entered_dt = vData
End Property
Public Property Get p_entered_dt() As Variant
        p_entered_dt = mvarp_entered_dt
End Property
Public Property Let p_bypass_nm_addr_fl(ByVal vData As Variant)
    mvarp_bypass_nm_addr_fl = vData
End Property
Public Property Get p_bypass_nm_addr_fl() As Variant
        p_bypass_nm_addr_fl = mvarp_bypass_nm_addr_fl
End Property
Public Property Let p_bypass_desc_fl(ByVal vData As Variant)
    mvarp_bypass_desc_fl = vData
End Property
Public Property Get p_bypass_desc_fl() As Variant
        p_bypass_desc_fl = mvarp_bypass_desc_fl
End Property
Public Property Let p_drop_ship_sales_ord_id(ByVal vData As Variant)
    mvarp_drop_ship_sales_ord_id = vData
End Property
Public Property Get p_drop_ship_sales_ord_id() As Variant
        p_drop_ship_sales_ord_id = mvarp_drop_ship_sales_ord_id
End Property
Public Property Let p_sales_tax_amnt(ByVal vData As Variant)
    mvarp_sales_tax_amnt = vData
End Property
Public Property Get p_sales_tax_amnt() As Variant
        p_sales_tax_amnt = mvarp_sales_tax_amnt
End Property
Public Property Let p_sales_tax_pct(ByVal vData As Variant)
    mvarp_sales_tax_pct = vData
End Property
Public Property Get p_sales_tax_pct() As Variant
        p_sales_tax_pct = mvarp_sales_tax_pct
End Property

Public Property Let p_order_text(ByVal vData As Variant)
    mvarp_order_text = vData
End Property
Public Property Get p_order_text() As Variant
    p_order_text = mvarp_order_text
End Property

Public Property Set o_stnd_name(ByVal vData As Object)
    Set mvaro_stnd_name = vData
End Property
Public Property Get o_stnd_name() As Object
    Set o_stnd_name = mvaro_stnd_name
End Property

Public Property Set o_ship_name(ByVal vData As Object)
    Set mvaro_ship_name = vData
End Property
Public Property Get o_ship_name() As Object
    Set o_ship_name = mvaro_ship_name
End Property


Public Property Set o_bill_name(ByVal vData As Object)
    Set mvaro_bill_name = vData
End Property
Public Property Get o_bill_name() As Object
    Set o_bill_name = mvaro_bill_name
End Property



Public Property Set o_old_purchase_order(ByVal vData As Object)
    Set mvaro_old_purchase_order = vData
End Property
Public Property Get o_old_purchase_order() As Object
    Set o_old_purchase_order = mvaro_old_purchase_order
End Property

Public Property Set o_stnd_address(ByVal vData As Object)
    Set mvaro_stnd_address = vData
End Property
Public Property Get o_stnd_address() As Object
    Set o_stnd_address = mvaro_stnd_address
End Property

Public Property Set o_ship_address(ByVal vData As Object)
    Set mvaro_ship_address = vData
End Property
Public Property Get o_ship_address() As Object
    Set o_ship_address = mvaro_ship_address
End Property


Public Property Set o_bill_address(ByVal vData As Object)
    Set mvaro_bill_address = vData
End Property
Public Property Get o_bill_address() As Object
    Set o_bill_address = mvaro_bill_address
End Property


Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Let p_freight_desc(ByVal vData As Variant)
    mvarp_freight_desc = vData
End Property
Public Property Get p_freight_desc() As Variant
        p_freight_desc = mvarp_freight_desc
End Property
Public Property Let p_comm_meth_desc(ByVal vData As Variant)
    mvarp_comm_meth_desc = vData
End Property
Public Property Get p_comm_meth_desc() As Variant
        p_comm_meth_desc = mvarp_comm_meth_desc
End Property
Public Property Let p_comm_meth_cd(ByVal vData As Variant)
    mvarp_comm_meth_cd = vData
End Property
Public Property Get p_comm_meth_cd() As Variant
        p_comm_meth_cd = mvarp_comm_meth_cd
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
Public Property Let p_type_desc(ByVal vData As Variant)
    mvarp_type_desc = vData
End Property
Public Property Set p_type_desc(ByVal vData As Variant)
    Set mvarp_type_desc = vData
End Property
Public Property Get p_type_desc() As Variant
        p_type_desc = mvarp_type_desc
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    mvarp_type_cd = vData
End Property
Public Property Set p_type_cd(ByVal vData As Variant)
    Set mvarp_type_cd = vData
End Property
Public Property Get p_type_cd() As Variant
        p_type_cd = mvarp_type_cd
End Property
Public Property Let p_transportation_desc(ByVal vData As Variant)
    mvarp_transportation_desc = vData
End Property
Public Property Get p_transportation_desc() As Variant
        p_transportation_desc = mvarp_transportation_desc
End Property
Public Property Let p_transportation_cd(ByVal vData As Variant)
    mvarp_transportation_cd = vData
End Property
Public Property Get p_transportation_cd() As Variant
        p_transportation_cd = mvarp_transportation_cd
End Property
Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property
Public Property Let p_notif_meth_desc(ByVal vData As Variant)
    mvarp_notif_meth_desc = vData
End Property
Public Property Get p_notif_meth_desc() As Variant
        p_notif_meth_desc = mvarp_notif_meth_desc
End Property
Public Property Let p_notif_meth_cd(ByVal vData As Variant)
    mvarp_notif_meth_cd = vData
End Property
Public Property Get p_notif_meth_cd() As Variant
        p_notif_meth_cd = mvarp_notif_meth_cd
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
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
Public Property Let p_revision_nbr(ByVal vData As Variant)
    mvarp_revision_nbr = vData
End Property
Public Property Get p_revision_nbr() As Variant
        p_revision_nbr = mvarp_revision_nbr
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
Public Property Let p_opened_dt(ByVal vData As Variant)
    mvarp_opened_dt = vData
End Property
Public Property Get p_opened_dt() As Variant
        p_opened_dt = mvarp_opened_dt
End Property
Public Property Let p_net_amnt(ByVal vData As Variant)
    mvarp_net_amnt = vData
End Property
Public Property Get p_net_amnt() As Variant
        p_net_amnt = mvarp_net_amnt
End Property
Public Property Let p_line_item_qty(ByVal vData As Variant)
    mvarp_line_item_qty = vData
End Property
Public Property Get p_line_item_qty() As Variant
        p_line_item_qty = mvarp_line_item_qty
End Property
Public Property Let p_gross_amnt(ByVal vData As Variant)
    mvarp_gross_amnt = vData
End Property
Public Property Get p_gross_amnt() As Variant
        p_gross_amnt = mvarp_gross_amnt
End Property
Public Property Let p_freight_cd(ByVal vData As Variant)
    mvarp_freight_cd = vData
End Property
Public Property Get p_freight_cd() As Variant
        p_freight_cd = mvarp_freight_cd
End Property
Public Property Let p_fob_terms_desc(ByVal vData As Variant)
    mvarp_fob_terms_desc = vData
End Property
Public Property Get p_fob_terms_desc() As Variant
        p_fob_terms_desc = mvarp_fob_terms_desc
End Property
Public Property Let p_fob_terms_cd(ByVal vData As Variant)
    mvarp_fob_terms_cd = vData
End Property
Public Property Get p_fob_terms_cd() As Variant
        p_fob_terms_cd = mvarp_fob_terms_cd
End Property
Public Property Let p_drop_ship_desc(ByVal vData As Variant)
    mvarp_drop_ship_desc = vData
End Property
Public Property Get p_drop_ship_desc() As Variant
        p_drop_ship_desc = mvarp_drop_ship_desc
End Property
Public Property Let p_drop_ship_fl(ByVal vData As Variant)
    mvarp_drop_ship_fl = vData
End Property
Public Property Get p_drop_ship_fl() As Variant
        p_drop_ship_fl = mvarp_drop_ship_fl
End Property
Public Property Let p_currency_desc(ByVal vData As Variant)
    mvarp_currency_desc = vData
End Property
Public Property Get p_currency_desc() As Variant
        p_currency_desc = mvarp_currency_desc
End Property
Public Property Let p_currency_cd(ByVal vData As Variant)
    mvarp_currency_cd = vData
End Property
Public Property Get p_currency_cd() As Variant
        p_currency_cd = mvarp_currency_cd
End Property
Public Property Let p_confirmation_nbr(ByVal vData As Variant)
    mvarp_confirmation_nbr = vData
End Property
Public Property Get p_confirmation_nbr() As Variant
        p_confirmation_nbr = mvarp_confirmation_nbr
End Property
Public Property Let p_carrier_nm(ByVal vData As Variant)
    mvarp_carrier_nm = vData
End Property
Public Property Get p_carrier_nm() As Variant
        p_carrier_nm = mvarp_carrier_nm
End Property
Public Property Let p_buyer_assoc_nbr(ByVal vData As Variant)
    mvarp_buyer_assoc_nbr = vData
End Property
Public Property Get p_buyer_assoc_nbr() As Variant
        p_buyer_assoc_nbr = mvarp_buyer_assoc_nbr
End Property
Public Property Let p_buyer_assoc_id(ByVal vData As Variant)
    mvarp_buyer_assoc_id = vData
End Property
Public Property Get p_buyer_assoc_id() As Variant
        p_buyer_assoc_id = mvarp_buyer_assoc_id
End Property
Public Property Let p_backorder_desc(ByVal vData As Variant)
    mvarp_backorder_desc = vData
End Property
Public Property Get p_backorder_desc() As Variant
        p_backorder_desc = mvarp_backorder_desc
End Property
Public Property Let p_backorder_cd(ByVal vData As Variant)
    mvarp_backorder_cd = vData
End Property
Public Property Get p_backorder_cd() As Variant
        p_backorder_cd = mvarp_backorder_cd
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
        p_po_id = mvarp_po_id
End Property
Public Sub getPurchaseOrder()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
     
    Dim strRoutine_nm As String
    strRoutine_nm = "purchsase_order.cls.GetPurchaseOrder"
    
    If p_po_id > 0 Then
    Else
        p_po_id = Null
        Exit Sub
    End If
    
    Call SelectPurchaseOrder
    
    If IsNull(p_po_id) = True Then
        Set o_stnd_name = CreateObject("converge.name")
        Set o_stnd_address = CreateObject("converge.address")
        o_stnd_name.p_nm_id = Null
        o_stnd_address.p_addr_id = Null
    Else
        If p_bypass_nm_addr_fl = "Y" Then
        Else
            Set o_stnd_name = CreateObject("converge.name")
            Set o_stnd_name.o_trans_data = Me.o_trans_data
            Set o_stnd_name.o_error = Me.o_error
            o_stnd_name.p_nm_id = Null
            o_stnd_name.p_obj_nm = "Purchase_order"
            o_stnd_name.p_obj_id = p_po_id
            o_stnd_name.p_nm_type_cd = "STND"
            o_stnd_name.p_def_nm_fl = "Y"
            Call o_stnd_name.getDefName
                          
            Set o_stnd_address = CreateObject("converge.address")
            Set o_stnd_address.o_trans_data = Me.o_trans_data
            Set o_stnd_address.o_error = Me.o_error
            o_stnd_address.p_addr_id = Null
            o_stnd_address.p_obj_nm = "Purchase_order"
            o_stnd_address.p_obj_id = p_po_id
            o_stnd_address.p_addr_type_cd = "STND"
            o_stnd_address.p_def_addr_fl = "Y"
            Call o_stnd_address.getDefAddress
        
        
            Set o_ship_name = CreateObject("converge.name")
            Set o_ship_name.o_trans_data = Me.o_trans_data
            Set o_ship_name.o_error = Me.o_error
            o_ship_name.p_nm_id = Null
            o_ship_name.p_obj_nm = "Purchase_order"
            o_ship_name.p_obj_id = p_po_id
            o_ship_name.p_nm_type_cd = "SHIP"
            o_ship_name.p_def_nm_fl = "Y"
            Call o_ship_name.getDefName
                          
            Set o_ship_address = CreateObject("converge.address")
            Set o_ship_address.o_trans_data = Me.o_trans_data
            Set o_ship_address.o_error = Me.o_error
            o_ship_address.p_addr_id = Null
            o_ship_address.p_obj_nm = "Purchase_order"
            o_ship_address.p_obj_id = p_po_id
            o_ship_address.p_addr_type_cd = "SHIP"
            o_ship_address.p_def_addr_fl = "Y"
            Call o_ship_address.getDefAddress
        
            Set o_bill_name = CreateObject("converge.name")
            Set o_bill_name.o_trans_data = Me.o_trans_data
            Set o_bill_name.o_error = Me.o_error
            o_bill_name.p_nm_id = Null
            o_bill_name.p_obj_nm = "Purchase_order"
            o_bill_name.p_obj_id = p_po_id
            o_bill_name.p_nm_type_cd = "BILL"
            o_bill_name.p_def_nm_fl = "Y"
            Call o_bill_name.getDefName
                          
            Set o_bill_address = CreateObject("converge.address")
            Set o_bill_address.o_trans_data = Me.o_trans_data
            Set o_bill_address.o_error = Me.o_error
            o_bill_address.p_addr_id = Null
            o_bill_address.p_obj_nm = "Purchase_order"
            o_bill_address.p_obj_id = p_po_id
            o_bill_address.p_addr_type_cd = "BILL"
            o_bill_address.p_def_addr_fl = "Y"
            Call o_bill_address.getDefAddress
        
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

Public Sub Add()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.Add"
    
    Call Validate("add")
    Call IssueInsertData
    Call AddName(o_stnd_name)
    Call AddAddr(o_stnd_address)
    Call AddName(o_ship_name)
    Call AddAddr(o_ship_address)
    Call AddName(o_bill_name)
    Call AddAddr(o_bill_address)
    
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
    strRoutine_nm = "Purchase_order.cls.Change"
    
    If Me.o_old_purchase_order Is Nothing Then
        Set Me.o_old_purchase_order = New purchase_order
        Set Me.o_old_purchase_order.o_trans_data = Me.o_trans_data
        Set Me.o_old_purchase_order.o_error = Me.o_error
        Me.o_old_purchase_order.p_po_id = p_po_id
        Me.o_old_purchase_order.p_bypass_nm_addr_fl = "N"
        Me.o_old_purchase_order.p_bypass_desc_fl = "Y"
        Me.o_old_purchase_order.getPurchaseOrder
    End If
    
    If o_old_purchase_order.p_status_cd = "ST" Then
        Call MoveToHistory
    End If
    
    Call Validate("change")
    Call IssueUpdateData
    
    If p_bypass_nm_addr_fl = "Y" Then
    Else
        If fValuePresent(o_stnd_name.p_nm_id) = True Then
           If fValuePresent(o_stnd_name.p_nm) = True Then
                Call ChangeName(o_stnd_name)
            Else
                Call DeleteName(o_stnd_name)
            End If
        Else
            If fValuePresent(o_stnd_name.p_nm) = True Then
                Call AddName(o_stnd_name)
            End If
        End If
        If fValuePresent(o_stnd_address.p_addr_id) = True Then
            If fValuePresent(o_stnd_address.p_Street_addr) = True And fValuePresent(o_stnd_address.p_City_nm) = True And fValuePresent(o_stnd_address.p_City_nm) = True Then
                Call ChangeAddr(o_stnd_address)
            Else
                Call DeleteAddr(o_stnd_address)
            End If
        Else
            If fValuePresent(o_stnd_address.p_Street_addr) = True And fValuePresent(o_stnd_address.p_City_nm) = True And fValuePresent(o_stnd_name.p_City_nm) = True Then
                Call AddAddr(o_stnd_address)
            End If
        End If
        If fValuePresent(o_ship_name.p_nm_id) = True Then
            If fValuePresent(o_ship_name.p_nm) = True Then
                Call ChangeName(o_ship_name)
            Else
                Call DeleteName(o_ship_name)
            End If
        Else
            If fValuePresent(o_ship_name.p_nm) = True Then
                Call AddName(o_ship_name)
            End If
        End If
        If fValuePresent(o_ship_address.p_addr_id) = True Then
           If fValuePresent(o_ship_address.p_Street_addr) = True And fValuePresent(o_ship_address.p_City_nm) = True And fValuePresent(o_ship_address.p_City_nm) = True Then
                Call ChangeAddr(o_ship_address)
            Else
                Call DeleteAddr(o_ship_address)
            End If
        Else
            If fValuePresent(o_ship_address.p_Street_addr) = True And fValuePresent(o_ship_address.p_City_nm) = True And fValuePresent(o_ship_address.p_City_nm) = True Then
                Call AddAddr(o_ship_address)
            End If
        End If
        If fValuePresent(o_bill_name.p_nm_id) = True Then
           If fValuePresent(o_bill_name.p_nm) = True Then
                Call ChangeName(o_bill_name)
            Else
                Call DeleteName(o_bill_name)
            End If
        Else
            If fValuePresent(o_bill_name.p_nm) = True Then
                Call AddName(o_bill_name)
            End If
        End If
        If fValuePresent(o_bill_address.p_addr_id) = True Then
           If fValuePresent(o_bill_address.p_Street_addr) = True And fValuePresent(o_bill_address.p_City_nm) = True And fValuePresent(o_bill_address.p_City_nm) = True Then
                Call ChangeAddr(o_bill_address)
            Else
                Call DeleteAddr(o_bill_address)
            End If
        Else
            If fValuePresent(o_bill_address.p_Street_addr) = True And fValuePresent(o_bill_address.p_City_nm) = True And fValuePresent(o_bill_address.p_City_nm) = True Then
                Call AddAddr(o_bill_address)
            End If
        End If
    End If
    
    If o_old_purchase_order.p_status_cd = "ST" Then
        p_status_cd = "RS"
        p_move_to_hist_fl = "N"
        Call ChangeStatus
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
Public Sub Delete()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.Delete"
    
    Call getPurchaseOrder
    Call Validate("delete")
    Call IssueDeleteName
    Call IssueDeleteAddr
    Call IssueDeleteItem
    Call IssueDeleteData
    
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

Public Sub AddPOfromRequirement(ByVal varItemID As Variant, ByVal varOrderedQty As Variant, ByVal varRequestedDelDt As Variant)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    'Dim objSalesOrder As Object
    Dim objSupplier As Object
    Dim objPurchaseOrderItem As Object
    Dim strSqlTx As String
    Dim objRecordset As ADODB.Recordset
            
    strRoutine_nm = "Purchase_order.cls.AddPOfromRequirment"

    If varItemID > 0 Then
    Else
       Me.o_error.p_type_cd = "F"
       Me.o_error.p_err_cd = "0500"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 1045
       Exit Sub
    End If
    
    If varOrderedQty > 0 Then
    Else
       Me.o_error.p_type_cd = "F"
       Me.o_error.p_err_cd = "0510"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 2122
       Exit Sub
    End If
    
    'Set objSalesOrder = CreateObject("converge_sales.sales_order")
    'Set objSalesOrder.o_error = Me.o_error
    'Set objSalesOrder.o_trans_data = Me.o_trans_data
    'objSalesOrder.p_sales_ord_id = p_drop_ship_sales_ord_id
    'objSalesOrder.GetSalesOrder
    
    'p_drop_ship_sales_ord_id = Nothing
    
    'p_carrier_nm = objSalesOrder.p_carrier_nm
    
    Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
    Set objPurchaseOrderItem.o_error = Me.o_error
    Set objPurchaseOrderItem.o_trans_data = Me.o_trans_data
    
    Call GetDefaultValues(varItemID, objPurchaseOrderItem)
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

                    'begin -- Matt 8/2/2003
                    'check to see if a po for the supplier in planned status already
                    'exist. if so, add the new item to that po.
    strSqlTx = "select * from purchase_order where supplier_id = " & Me.p_supplier_id & " and status_cd = 'PL' and drop_ship_fl = 'N' order by entered_ts desc "
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSqlTx, Me.o_error)
    If objRecordset.EOF = True Then
        Call Add
    Else
        Me.p_po_id = objRecordset("po_id")
        Me.getPurchaseOrder
    End If
                    'end -- Matt 8/2/2003
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    With objPurchaseOrderItem
        .p_po_id = p_po_id
        .p_item_id = varItemID
        .p_ordered_qty = varOrderedQty
        .p_requested_del_dt = varRequestedDelDt
        .p_inspection_req_fl = "Y"
        .p_payment_terms_cd = Null
        .p_type_cd = "PI"
    End With
   
    Call objPurchaseOrderItem.Add
    
    'Set objSalesOrder = Nothing
    Set objSupplier = Nothing
    Set objPurchaseOrderItem = Nothing
  
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

 Public Sub PlanDropShipSo()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objSalesOrder As Object
    
    strRoutine_nm = "Purchase_order.cls.PlanDropShipSo"

    If p_drop_ship_sales_ord_id > 0 Then
    Else
       Me.o_error.p_type_cd = "F"
       Me.o_error.p_err_cd = "0400"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 2132
       Exit Sub
    End If
    
    If p_po_id > 0 Then
        Set o_old_purchase_order = CreateObject("converge_purchase.purchase_order")
        Set o_old_purchase_order.o_trans_data = Me.o_trans_data
        Set o_old_purchase_order.o_error = Me.o_error
        o_old_purchase_order.p_po_id = p_po_id
        o_old_purchase_order.p_bypass_nm_addr_fl = "N"
        o_old_purchase_order.p_bypass_desc_fl = "Y"
        o_old_purchase_order.getPurchaseOrder
        If IsNull(o_old_purchase_order.p_po_id) = True Then
            Set o_old_purchase_order = Nothing
            p_po_id = 0
        End If
    End If
    
    Set objSalesOrder = CreateObject("converge_sales.sales_order")
    Set objSalesOrder.o_error = Me.o_error
    Set objSalesOrder.o_trans_data = Me.o_trans_data
    objSalesOrder.p_sales_ord_id = p_drop_ship_sales_ord_id
    objSalesOrder.GetSalesOrder
        
    If IsNull(objSalesOrder.p_sales_ord_id) = True Then
       Me.o_error.p_type_cd = "E"
       Me.o_error.p_err_cd = "0600"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 1010
       Exit Sub
    End If
    
    ' This is the main planning control routine.
    ' if po exist and supplier has not changed, update existing PO.
    ' if po exist and supplier has changed, update existing PO.
    ' if no po exist at all, just add a new PO.
    
    If p_po_id > 0 Then
        Dim colOldPOItems As Collection
        Set colOldPOItems = getPOItems(p_po_id)
        p_supplier_id = objSalesOrder.p_drop_ship_supplier_id
        If objSalesOrder.p_drop_ship_supplier_id = o_old_purchase_order.p_supplier_id And _
           (o_old_purchase_order.p_status_cd = "ST" Or _
            o_old_purchase_order.p_status_cd = "RS" Or _
            o_old_purchase_order.p_status_cd = "CP") Then
            
            If o_old_purchase_order.p_status_cd = "ST" Then
                p_status_cd = "RS"
                Call ChangeStatus
                Call IssueDeleteItem
                Call AddLineItems(False, colOldPOItems)
                    'change status to resend again to update the line items
                p_status_cd = "RS"
                Call ChangeStatus
            Else
                Call IssueDeleteItem
                Call AddLineItems(False, colOldPOItems)
                p_status_cd = "RS"
                Call ChangeStatus
            End If
        Else
            Call updateExistingPOs
            Call AddPOfromDropShipSo(objSalesOrder)
            objSalesOrder.p_sales_ord_id = p_drop_ship_sales_ord_id
            objSalesOrder.p_drop_ship_po_id = p_po_id
            Call objSalesOrder.setDropShipPoId
            Call AddLineItems(False, colOldPOItems)
            If o_old_purchase_order.p_status_cd = "RS" Then
                p_status_cd = "RS"
                Call ChangeStatus
            End If
        End If
    Else
        Set colOldPOItems = New Collection
        Call AddPOfromDropShipSo(objSalesOrder)
        Call AddLineItems(True, colOldPOItems)
        objSalesOrder.p_sales_ord_id = p_drop_ship_sales_ord_id
        objSalesOrder.p_drop_ship_po_id = p_po_id
        Call objSalesOrder.setDropShipPoId
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

Private Sub AddPOfromDropShipSo(ByRef r_objSalesOrder As Object)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objSupplier As Object
    Dim objItem As Object
    Dim objInvLoc As Object
    strRoutine_nm = "Purchase_order.cls.AddPOfromDropShipSo"

    Set objSupplier = CreateObject("converge_purchase.supplier")
    Set objSupplier.o_error = Me.o_error
    Set objSupplier.o_trans_data = Me.o_trans_data
    objSupplier.p_supplier_id = r_objSalesOrder.p_drop_ship_supplier_id
    objSupplier.GetSupplier
        
    If IsNull(objSupplier.p_supplier_id) = True Then
       Me.o_error.p_type_cd = "E"
       Me.o_error.p_err_cd = "0700"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 2080
       Exit Sub
    End If
    
    Set objInvLoc = CreateObject("converge_inv.inventory_loc")
    Set objInvLoc.o_error = Me.o_error
    Set objInvLoc.o_trans_data = Me.o_trans_data
    objInvLoc.p_inv_loc_id = r_objSalesOrder.p_inv_loc_id
    objInvLoc.GetDefaultInventoryLoc
    
    If IsNull(objInvLoc.p_inv_loc_id) = True Then
       Me.o_error.p_type_cd = "E"
       Me.o_error.p_err_cd = "0700"
       Me.o_error.p_routine_nm = strRoutine_nm
       Me.o_error.p_message_id = 1175
       Exit Sub
    End If
    
    Set o_stnd_name = CreateObject("converge.name")
    Set o_stnd_name = objSupplier.o_stnd_name
    Set o_stnd_address = CreateObject("converge.address")
    Set o_stnd_address = objSupplier.o_stnd_address
    
    Set o_ship_name = CreateObject("converge.name")
    Set o_ship_name = r_objSalesOrder.o_ship_name
    Set o_ship_address = CreateObject("converge.address")
    Set o_ship_address = r_objSalesOrder.o_ship_address
    o_ship_name.p_nm_type_cd = "SHIP"
    o_ship_address.p_addr_type_cd = "SHIP"
    
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_name = objInvLoc.o_stnd_name
    Set o_bill_address = CreateObject("converge.address")
    Set o_bill_address = objInvLoc.o_stnd_address
    o_bill_name.p_nm_type_cd = "BILL"
    o_bill_address.p_addr_type_cd = "BILL"
    
    p_fob_terms_cd = r_objSalesOrder.p_fob_terms_cd
    p_currency_cd = r_objSalesOrder.p_currency_cd
    p_comm_meth_cd = objSupplier.p_comm_meth_cd
    p_freight_cd = r_objSalesOrder.p_freight_cd
    p_carrier_nm = r_objSalesOrder.p_carrier_nm
    p_transportation_cd = r_objSalesOrder.p_transportation_cd
    p_payment_terms_cd = objSupplier.p_payment_terms_cd
    p_buyer_assoc_id = Me.o_trans_data.p_assoc_id
    p_drop_ship_fl = "Y"
    p_supplier_id = r_objSalesOrder.p_drop_ship_supplier_id
    p_supplier_nbr = Null
    p_inv_loc_id = objInvLoc.p_inv_loc_id
    
    Call Add
    
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

Private Sub AddLineItems(ByVal v_blnInitialAddFl As Boolean, ByRef inPOItems As Collection)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objSalesOrderItem As Object
    Dim objItem As Object
    Dim objSupplierItem As supplier_item
    Dim objPurchaseOrderItem As purchase_order_item
    Dim objRecordset As ADODB.Recordset
    strRoutine_nm = "Purchase_order.cls.AddLineItems"

    strSql = "select * from sales_order_item " & _
             " where " & fWhereVariable("sales_ord_id", p_drop_ship_sales_ord_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        objRecordset.Close
        Set objRecordset = Nothing
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "0800"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2133
        Exit Sub
    End If
    
    Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
    Set objPurchaseOrderItem.o_error = Me.o_error
    Set objPurchaseOrderItem.o_trans_data = Me.o_trans_data
    
    Set objSalesOrderItem = CreateObject("converge_sales.sales_order_item")
    Set objSalesOrderItem.o_error = Me.o_error
    Set objSalesOrderItem.o_trans_data = Me.o_trans_data
    
    Set objItem = CreateObject("converge_item.item")
    Set objItem.o_error = Me.o_error
    Set objItem.o_trans_data = Me.o_trans_data
    
    Set objSupplierItem = CreateObject("converge_purchase.supplier_item")
    Set objSupplierItem.o_error = Me.o_error
    Set objSupplierItem.o_trans_data = Me.o_trans_data
    
    Do While objRecordset.EOF = False And Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
    
        objSupplierItem.p_supplier_id = p_supplier_id
        objSupplierItem.p_item_id = objRecordset("item_id")
        objSupplierItem.GetSupplierItem
        
        objItem.p_item_id = objRecordset("item_id")
        objItem.GetItem
        
        If IsNull(objSupplierItem.p_item_id) = False Then
            If Len(objSupplierItem.p_supplier_item_nbr) > 0 Then
                objPurchaseOrderItem.p_supplier_item_nbr = objItem.p_item_nbr
            Else
                objPurchaseOrderItem.p_supplier_item_nbr = objSupplierItem.p_supplier_item_nbr
            End If
            If Len(objSupplierItem.p_supplier_item_desc) > 0 Then
                objPurchaseOrderItem.p_item_desc = objSupplierItem.p_supplier_item_desc
            Else
                objPurchaseOrderItem.p_item_desc = objItem.p_item_desc
            End If
            If IsNumeric(objPurchaseOrderItem.p_item_cost = objSupplierItem.p_supplier_item_cost) Then
               objPurchaseOrderItem.p_item_cost = objSupplierItem.p_supplier_item_cost
            Else
               objPurchaseOrderItem.p_item_cost = objItem.p_item_cost
            End If
        Else
            With objPurchaseOrderItem
                .p_supplier_item_nbr = objItem.p_item_nbr
                .p_item_desc = objItem.p_item_desc
                .p_item_cost = objItem.p_item_cost
            End With
        End If
        
        If Not IsNull(inPOItems) Then
            Dim itemId As String
            Dim poLineNbr As String
            itemId = objRecordset("item_id")
            If Not IsNull(objRecordset("drop_ship_po_line_nbr")) Then
                poLineNbr = objRecordset("drop_ship_po_line_nbr")
                
                Dim poItem As purchase_order_item
                Set poItem = inPOItems.item(poLineNbr)
                If Not IsNull(poItem) Then
                    If poItem.p_item_id = itemId Then
                        objPurchaseOrderItem.p_item_cost = poItem.p_item_cost
                    End If
                End If
            End If
        End If
        
        With objPurchaseOrderItem
            .p_po_id = p_po_id
            .p_item_id = objRecordset("item_id")
            .p_ordered_qty = objRecordset("ordered_qty")
            .p_requested_del_dt = objRecordset("requested_del_dt")
            .p_drop_ship_so_line_nbr = objRecordset("line_nbr")
            .p_inspection_req_fl = "N"
            .p_payment_terms_cd = Null
            .p_type_cd = "PI"
            .p_item_text = objRecordset("item_tx")
            Call .Add
            Call .UpdateQuantities(Me.o_error, Me.o_trans_data, Me.p_po_id, objPurchaseOrderItem.p_line_nbr)
            If objRecordset("status_cd") = "CA" Then
                .p_status_cd = "CA"
                Call .ChangeStatus
            End If
            If objRecordset("status_cd") = "CL" Then
                .p_status_cd = "CL"
                Call .ChangeStatus
            End If
            If objRecordset("status_cd") = "OH" Then
                .p_status_cd = "OH"
                Call .ChangeStatus
            End If
        End With
        
                    'for the following items generate a material requirement for
                    'a component item to be sent to the supplier, Phoenix.
        If v_blnInitialAddFl = True Then
            If objItem.p_item_id = "25792" Or _
               objItem.p_item_id = "25793" Or _
               objItem.p_item_id = "25794" Or _
               objItem.p_item_id = "25795" Or _
               objItem.p_item_id = "25796" Or _
               objItem.p_item_id = "25797" Or _
               objItem.p_item_id = "25798" Or _
               objItem.p_item_id = "25799" Then
               
                Dim objMatReq As Object
                Dim lngItemId As Long
                Dim lngOrderedQty As Long
                Dim dteReqDelDt As Date
                Dim objItems As Object
                Dim lngLeadDays As Long
                
                lngOrderedQty = objRecordset("ordered_qty")
                
                Select Case objItem.p_item_id
                    Case 25792
                        lngItemId = 25779
                    Case 25793
                        lngItemId = 25780
                    Case 25794
                        lngItemId = 25781
                    Case 25795
                        lngItemId = 25782
                    Case 25796
                        lngItemId = 25783
                    Case 25797
                        lngItemId = 25785
                    Case 25798
                        lngItemId = 25784
                    Case 25799
                        lngItemId = 25788
                End Select
                
                Dim objItem2 As Object
                Set objItem2 = CreateObject("converge_item.item")
                Set objItem2.o_error = Me.o_error
                Set objItem2.o_trans_data = Me.o_trans_data
                objItem2.p_item_id = lngItemId
                objItem2.GetItem
                
                dteReqDelDt = CDate(objRecordset("requested_del_dt"))
                lngLeadDays = CLng(objItem2.p_lead_days)
                dteReqDelDt = dteReqDelDt - lngLeadDays
                
                Set objMatReq = CreateObject("converge_prod.mat_req")
                Call objMatReq.addMatReq( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , "MR" _
                            , Empty _
                            , Empty _
                            , Empty _
                            , Empty _
                            , lngItemId _
                            , "Phoenix" _
                            , lngOrderedQty _
                            , dteReqDelDt _
                            , 0.1 _
                            , 0 _
                            , 2048 _
                            )
                            
                Set objMatReq.o_error = Me.o_error
                Set objMatReq.o_trans_data = Me.o_trans_data
                Call objMatReq.generatePlan
            End If
        End If
        
        objSalesOrderItem.p_sales_ord_id = p_drop_ship_sales_ord_id
        objSalesOrderItem.p_line_nbr = objRecordset("line_nbr")
        objSalesOrderItem.p_drop_ship_po_line_nbr = objPurchaseOrderItem.p_line_nbr
        Call objSalesOrderItem.setDropShipPoLineNbr(objPurchaseOrderItem.p_status_cd)
        objRecordset.MoveNext
    Loop
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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

' this procedure moves a po from the po & poi tables to the history tables
Public Sub MoveToHistory()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.move"
    Dim strSql As String
    Dim dblUpdatedQty As Double
    
    If o_old_purchase_order Is Nothing Then
        Set o_old_purchase_order = CreateObject("converge_purchase.purchase_order")
        Set o_old_purchase_order.o_error = Me.o_error
        Set o_old_purchase_order.o_trans_data = Me.o_trans_data
        o_old_purchase_order.p_bypass_nm_addr_fl = "N"
        o_old_purchase_order.p_bypass_desc_fl = "Y"
        o_old_purchase_order.p_po_id = p_po_id
        o_old_purchase_order.getPurchaseOrder
    End If
    
    
    strSql = "DELETE FROM Purchase_order_hist " & _
            "  WHERE " & fWhereVariable("po_id", o_old_purchase_order.p_po_id, "N", "=") & _
            "  and   " & fWhereVariable("revision_nbr", o_old_purchase_order.p_revision_nbr, "N", "=")
            
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    
    strSql = "DELETE FROM Purchase_order_item_hist " & _
            "  WHERE " & fWhereVariable("po_id", o_old_purchase_order.p_po_id, "N", "=") & _
            "  and   " & fWhereVariable("po_revision_nbr", o_old_purchase_order.p_revision_nbr, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    
    strSql = "insert into purchase_order_hist ( " & _
                        "  po_id" & _
                        ", backorder_cd" & _
                        ", buyer_assoc_id" & _
                        ", carrier_nm" & _
                        ", confirmation_nbr" & _
                        ", currency_cd" & _
                        ", drop_ship_fl" & _
                        ", drop_ship_sales_ord_id" & _
                        ", fob_terms_cd" & _
                        ", freight_cd" & _
                        ", comm_meth_cd" & _
                        ", entered_dt" & _
                        ", opened_dt" & _
                        ", payment_terms_cd" & _
                        ", revision_nbr" & _
                        ", status_cd" & _
                        ", sales_tax_pct" & _
                        ", supplier_id" & _
                        ", transportation_cd" & _
                        ", type_cd" & _
                        ", inv_loc_id"
    strSql = strSql & "  , send_to_nm" & _
                        ", send_to_street_addr" & _
                        ", send_to_city_nm" & _
                        ", send_to_state_cd" & _
                        ", send_to_zip_cd" & _
                        ", send_to_country_cd"
    strSql = strSql & "  , bill_to_nm" & _
                        ", bill_to_street_addr" & _
                        ", bill_to_city_nm" & _
                        ", bill_to_state_cd" & _
                        ", bill_to_zip_cd" & _
                        ", bill_to_country_cd"
    strSql = strSql & "  , ship_to_nm" & _
                        ", ship_to_street_addr" & _
                        ", ship_to_city_nm" & _
                        ", ship_to_state_cd" & _
                        ", ship_to_zip_cd" & _
                        ", ship_to_country_cd"
    strSql = strSql & "  , updated_assoc_id" & _
                        ", updated_ts" & _
                        ", entered_assoc_id" & _
                        ", entered_ts" & _
                        " ) "
    strSql = strSql & " values (" & _
                        " " & fInsertVariable(o_old_purchase_order.p_po_id, "N") & _
                        "," & fInsertVariable(o_old_purchase_order.p_backorder_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_buyer_assoc_id, "N") & _
                        "," & fInsertVariable(o_old_purchase_order.p_carrier_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_confirmation_nbr, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_currency_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_drop_ship_fl, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_drop_ship_sales_ord_id, "N") & _
                        "," & fInsertVariable(o_old_purchase_order.p_fob_terms_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_freight_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_comm_meth_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_entered_dt, "D") & _
                        "," & fInsertVariable(o_old_purchase_order.p_opened_dt, "D") & _
                        "," & fInsertVariable(o_old_purchase_order.p_payment_terms_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_revision_nbr, "N")
    strSql = strSql & ",  " & fInsertVariable(o_old_purchase_order.p_status_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_sales_tax_pct, "N") & _
                        "," & fInsertVariable(o_old_purchase_order.p_supplier_id, "N") & _
                        "," & fInsertVariable(o_old_purchase_order.p_transportation_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_type_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.p_inv_loc_id, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_name.p_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_address.p_Street_addr, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_address.p_City_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_address.p_state_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_address.p_zip_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_stnd_address.p_country_cd, "S")
    strSql = strSql & ",  " & fInsertVariable(o_old_purchase_order.o_bill_name.p_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_bill_address.p_Street_addr, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_bill_address.p_City_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_bill_address.p_state_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_bill_address.p_zip_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_bill_address.p_country_cd, "S")
    strSql = strSql & ",  " & fInsertVariable(o_old_purchase_order.o_ship_name.p_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_ship_address.p_Street_addr, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_ship_address.p_City_nm, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_ship_address.p_state_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_ship_address.p_zip_cd, "S") & _
                        "," & fInsertVariable(o_old_purchase_order.o_ship_address.p_country_cd, "S")
    strSql = strSql & ",  " & fInsertVariable(o_old_purchase_order.o_trans_data.p_assoc_id, "N") & _
                        "," & fInsertVariable("current_ts", "D") & _
                        "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N") & _
                        "," & fInsertVariable("current_ts", "D") & _
                        ")"
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)

    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    strSql = " insert into purchase_order_item_hist" & _
                      " ( " & _
                      "  po_id" & _
                      ", line_nbr" & _
                      ", po_revision_nbr" & _
                      ", drop_ship_so_line_nbr" & _
                      ", item_desc" & _
                      ", type_cd" & _
                      ", supplier_item_nbr" & _
                      ", item_id" & _
                      ", ordered_qty" & _
                      ", item_cost" & _
                      ", uom_cd" & _
                      ", status_cd" & _
                      ", payment_terms_cd" & _
                      ", inspection_req_fl" & _
                      ", received_qty" & _
                      ", rejected_qty" & _
                      ", invoiced_qty" & _
                      ", updated_assoc_id" & _
                      ", updated_ts" & _
                      ", entered_assoc_id" & _
                      ", entered_ts" & _
                      " ) "
    strSql = strSql & " select " & _
                      "  po_id" & _
                      ", line_nbr" & _
                      ", " & fInsertVariable(o_old_purchase_order.p_revision_nbr, "N") & _
                      ", drop_ship_so_line_nbr" & _
                      ", item_desc" & _
                      ", type_cd" & _
                      ", supplier_item_nbr" & _
                      ", item_id" & _
                      ", ordered_qty" & _
                      ", item_cost" & _
                      ", uom_cd" & _
                      ", status_cd" & _
                      ", payment_terms_cd" & _
                      ", inspection_req_fl" & _
                      ", received_qty" & _
                      ", rejected_qty" & _
                      ", invoiced_qty" & _
                      ", updated_assoc_id" & _
                      ", updated_ts" & _
                      ", entered_assoc_id" & _
                      ", entered_ts" & _
                      " from purchase_order_item " & _
                      " WHERE " & fWhereVariable("po_id", p_po_id, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, Me.o_error)
                      
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
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

Private Sub AddName(objName)

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.AddName"
    
    Set objName.o_error = Me.o_error
    Set objName.o_trans_data = Me.o_trans_data
    objName.p_obj_nm = "Purchase_order"
    objName.p_obj_id = p_po_id
    Call objName.AddName
    
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
Private Sub ChangeName(objName)

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.ChangeName"
    
    Set objName.o_error = Me.o_error
    Set objName.o_trans_data = Me.o_trans_data
    objName.p_obj_nm = "Purchase_order"
    Call objName.ChangeName
    
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

Private Sub DeleteName(objName)

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.DeleteName"
    
    Set objName.o_error = Me.o_error
    Set objName.o_trans_data = Me.o_trans_data
    
    objName.p_obj_nm = "Purchase_order"
    Call objName.DeleteName
    
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

Private Sub AddAddr(objAddress)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.AddAddr"
    
    Set objAddress.o_error = Me.o_error
    Set objAddress.o_trans_data = Me.o_trans_data
    objAddress.p_obj_id = p_po_id
    objAddress.p_obj_nm = "Purchase_order"
    Call objAddress.AddAddress
    
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

Private Sub ChangeAddr(objAddress)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.ChangeAddr"
    
    Set objAddress.o_error = Me.o_error
    Set objAddress.o_trans_data = Me.o_trans_data
    objAddress.p_obj_nm = "Purchase_order"
    Call objAddress.ChangeAddress
    
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

Private Sub DeleteAddr(objAddress)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.DeleteSAddr"
    
    Set objAddress.o_error = Me.o_error
    Set objAddress.o_trans_data = Me.o_trans_data
    objAddress.p_obj_nm = "Purchase_order"
    Call objAddress.DeleteAddress
    
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
    strRoutine_nm = "Purchase_order.cls.IssueDeleteData"

    strSql = "DELETE FROM Purchase_order " & _
            "  WHERE " & fWhereVariable("po_id", p_po_id, "N", "=")
    
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

Private Sub IssueDeleteItem()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As ADODB.Recordset
    Dim objPurchaseOrderItem As Object
    
    strRoutine_nm = "Purchase_order.cls.IssueDeleteItem"

    Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
    
    Call getPurchaseOrderItems(objRecordset)
    
    Do Until objRecordset.EOF = True Or Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F"
        Set objPurchaseOrderItem.o_error = Me.o_error
        Set objPurchaseOrderItem.o_trans_data = Me.o_trans_data
        objPurchaseOrderItem.p_po_id = p_po_id
        objPurchaseOrderItem.p_line_nbr = objRecordset("line_nbr")
        objPurchaseOrderItem.p_bypass_del_edit_fl = "Y"
        Call objPurchaseOrderItem.Delete
        objRecordset.MoveNext
    Loop
    
    Set objPurchaseOrderItem = Nothing

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
Private Sub IssueDeleteName()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "Purchase_order.cls.IssueDeleteName"
    
    strSql = "DELETE FROM name " & _
            "  WHERE " & fWhereVariable("obj_id", p_po_id, "N", "=") & _
            "    AND " & fWhereVariable("obj_nm", "Purchase_order", "S", "=")
    
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
Private Sub IssueDeleteAddr()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "Purchase_Order.cls.IssueDeleteAddr"
    
    strSql = "delete from address " & _
            " where " & fWhereVariable("obj_id", p_po_id, "N", "=") & _
            "    AND " & fWhereVariable("obj_nm", "Purchase_order", "S", "=")
    
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

Private Sub Validate(sFunction_cd)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_order.cls.Validate"
    
    If fValuePresent(p_currency_cd) = True And _
       fGetDesc("currency_cd", p_currency_cd, Me.o_error, Me.o_trans_data) = Empty Then
        Me.o_error.p_type_cd = "F"
        Me.o_error.p_err_cd = "0900"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2015
        Exit Sub
    End If
        
    If fValuePresent(p_sales_tax_pct) = True Then
        If p_sales_tax_pct >= 0 And p_sales_tax_pct <= 1 Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0900"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 3309
            Exit Sub
        End If
    Else
        If IsNull(p_sales_tax_pct) = True Then
            p_sales_tax_pct = 0
        End If
    End If
        
    If fValuePresent(p_buyer_assoc_nbr) = True Then
        p_buyer_assoc_id = fGetAssocId(p_buyer_assoc_nbr, Me.o_error, Me.o_trans_data)
        If IsNull(p_buyer_assoc_id) = True Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0900"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 3310
            Exit Sub
        End If
    Else
        If IsNull(p_buyer_assoc_nbr) = True Then
            p_buyer_assoc_id = Null
        End If
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If fValuePresent(p_supplier_nbr) = True Or fValuePresent(p_supplier_id) = True Or sFunction_cd = "add" Then
            If fValuePresent(p_supplier_nbr) = True Then
                p_supplier_id = fGetSupplierId(p_supplier_nbr, Me.o_error, Me.o_trans_data)
            Else
                p_supplier_nbr = fGetSupplierNbr(p_supplier_id, Me.o_error, Me.o_trans_data)
            End If
            If IsNull(p_supplier_id) = True Or IsNull(p_supplier_nbr) = True Then
                Me.o_error.p_type_cd = "F"
                Me.o_error.p_err_cd = "1000"
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 2100
                Exit Sub
            End If
        End If
        
    End If
    
    If sFunction_cd = "change" Then
        If o_old_purchase_order.p_status_cd = "CA" Or _
           o_old_purchase_order.p_status_cd = "ST" Or _
           o_old_purchase_order.p_status_cd = "NC" Or _
           o_old_purchase_order.p_status_cd = "CL" Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1100"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2134
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "delete" Then
        If p_status_cd = "PL" Then
        Else
            Me.o_error.p_type_cd = "F"
            Me.o_error.p_err_cd = "1200"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2135
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

Private Sub IssueInsertData()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "Purchase_order.cls.IssueInsertData"
    
    
    If fValuePresent(p_sales_tax_pct) = False Then
        p_sales_tax_pct = 0
    End If
    
    p_po_id = fGetId("po_id", Me.o_error, Me.o_trans_data)
    p_revision_nbr = 0
    p_status_cd = "PL"
    If Len(p_type_cd) > 0 Then
    Else
        p_type_cd = "PO"
    End If
    
    strSql = "insert into purchase_order ( " & _
                        "  po_id" & _
                        ", backorder_cd" & _
                        ", buyer_assoc_id" & _
                        ", carrier_nm" & _
                        ", confirmation_nbr" & _
                        ", currency_cd" & _
                        ", drop_ship_fl" & _
                        ", drop_ship_sales_ord_id" & _
                        ", fob_terms_cd" & _
                        ", freight_cd" & _
                        ", comm_meth_cd" & _
                        ", entered_dt" & _
                        ", opened_dt" & _
                        ", payment_terms_cd" & _
                        ", order_tx" & _
                        ", revision_nbr"
    strSql = strSql & ", status_cd" & _
                        ", sales_tax_pct" & _
                        ", supplier_id" & _
                        ", transportation_cd" & _
                        ", type_cd" & _
                        ", inv_loc_id" & _
                        ", conv_from_ord_id " & _
                        ", updated_assoc_id" & _
                        ", updated_ts" & _
                        ", entered_assoc_id" & _
                        ", entered_ts" & _
                        " ) " & _
                        "values (" & _
                        " " & fInsertVariable(p_po_id, "N") & _
                        "," & fInsertVariable(p_backorder_cd, "S") & _
                        "," & fInsertVariable(p_buyer_assoc_id, "N") & _
                        "," & fInsertVariable(p_carrier_nm, "S") & _
                        "," & fInsertVariable(p_confirmation_nbr, "S") & _
                        "," & fInsertVariable(p_currency_cd, "S") & _
                        "," & fInsertVariable(p_drop_ship_fl, "S") & _
                        "," & fInsertVariable(p_drop_ship_sales_ord_id, "N") & _
                        "," & fInsertVariable(p_fob_terms_cd, "S") & _
                        "," & fInsertVariable(p_freight_cd, "S") & _
                        "," & fInsertVariable(p_comm_meth_cd, "S") & _
                        "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(p_opened_dt, "D") & _
                        "," & fInsertVariable(p_payment_terms_cd, "S") & _
                        "," & fInsertVariable(p_order_text, "S") & _
                        "," & fInsertVariable(1, "N") & _
                        "," & fInsertVariable(p_status_cd, "S") & _
                        "," & fInsertVariable(p_sales_tax_pct, "N") & _
                        "," & fInsertVariable(p_supplier_id, "N") & _
                        "," & fInsertVariable(p_transportation_cd, "S") & _
                        "," & fInsertVariable(p_type_cd, "S") & _
                        "," & fInsertVariable(p_inv_loc_id, "S") & _
                        "," & fInsertVariable(p_conv_from_ord_id, "S") & _
                        "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N") & _
                        "," & fInsertVariable("current_ts", "D") & _
                        "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N") & _
                        "," & fInsertVariable("current_ts", "D") & _
                        ")"
  
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
    strRoutine_nm = "Purchase_order.cls.IssueUpdateData"
    
    If IsNull(p_sales_tax_pct) = True Then
        p_sales_tax_pct = 0
    End If
    
    
    strSql = "UPDATE Purchase_order set " & _
                      fUpdateVariable("backorder_cd", p_backorder_cd, "S", "=") & _
                "," & fUpdateVariable("buyer_assoc_id", p_buyer_assoc_id, "N", "=") & _
                "," & fUpdateVariable("carrier_nm", p_carrier_nm, "S", "=") & _
                "," & fUpdateVariable("confirmation_nbr", p_confirmation_nbr, "S", "=") & _
                "," & fUpdateVariable("currency_cd", p_currency_cd, "S", "=") & _
                "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=") & _
                "," & fUpdateVariable("fob_terms_cd", p_fob_terms_cd, "S", "=") & _
                "," & fUpdateVariable("freight_cd", p_freight_cd, "S", "=") & _
                "," & fUpdateVariable("sales_tax_pct", p_sales_tax_pct, "N", "=") & _
                "," & fUpdateVariable("comm_meth_cd", p_comm_meth_cd, "S", "=") & _
                "," & fUpdateVariable("transportation_cd", p_transportation_cd, "S", "=") & _
                "," & fUpdateVariable("order_tx", p_order_text, "S", "=") & _
                "," & fUpdateVariable("updated_assoc_id", Me.o_trans_data.p_assoc_id, "N", "=") & _
                "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
              " WHERE" & fWhereVariable("po_id", p_po_id, "N", "=")
    
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

Private Sub SelectPurchaseOrder()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
   
    strRoutine_nm = "Purchase_order.cls.SelectPurchaseOrder"
    
    If IsNull(p_po_id) = True Then
     p_po_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM Purchase_order "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("po_id", p_po_id, "N", "=")
    End If
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)

    If objRecordset.EOF = True Then
        ' Close the Record Set
        objRecordset.Close
        Set objRecordset = Nothing
        p_po_id = Null
    Else
        p_po_id = objRecordset("po_id")
        p_buyer_assoc_id = objRecordset("buyer_assoc_id")
        p_buyer_assoc_nbr = fGetAssocNbr(p_buyer_assoc_id, Me.o_error, Me.o_trans_data)
        p_confirmation_nbr = objRecordset("confirmation_nbr")
        p_carrier_nm = objRecordset("carrier_nm")
        p_currency_cd = objRecordset("currency_cd")
        p_supplier_id = objRecordset("supplier_id")
        p_drop_ship_fl = objRecordset("drop_ship_fl")
        p_comm_meth_cd = objRecordset("comm_meth_cd")
        p_backorder_cd = objRecordset("backorder_cd")
        p_fob_terms_cd = objRecordset("fob_terms_cd")
        p_freight_cd = objRecordset("freight_cd")
        p_entered_dt = objRecordset("entered_dt")
        p_opened_dt = objRecordset("opened_dt")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_revision_nbr = objRecordset("revision_nbr")
        p_status_cd = objRecordset("status_cd")
        p_drop_ship_sales_ord_id = objRecordset("drop_ship_sales_ord_id")
        p_sales_tax_pct = objRecordset("sales_tax_pct")
        p_transportation_cd = objRecordset("transportation_cd")
        p_type_cd = objRecordset("type_cd")
        p_inv_loc_id = objRecordset("inv_loc_id")
        p_order_text = objRecordset("order_tx")
        
        ' Close the Record Set
        objRecordset.Close
        Set objRecordset = Nothing
        
        ' GetDesc is Called
        If p_bypass_desc_fl = "Y" Then
        Else
            p_supplier_nbr = fGetSupplierNbr(p_supplier_id, Me.o_error, Me.o_trans_data)
            p_currency_desc = fGetDesc("currency_cd", p_currency_cd, Me.o_error, Me.o_trans_data)
            p_type_desc = fGetDesc("Purchase_order_type_cd", p_type_cd, Me.o_error, Me.o_trans_data)
            p_comm_meth_desc = fGetDesc("comm_meth_cd", p_comm_meth_cd, Me.o_error, Me.o_trans_data)
            p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, Me.o_error, Me.o_trans_data)
            p_status_desc = fGetDesc("purchase_order_status_cd", p_status_cd, Me.o_error, Me.o_trans_data)
            p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, Me.o_error, Me.o_trans_data)
            p_fob_terms_desc = fGetDesc("fob_terms_cd", p_fob_terms_cd, Me.o_error, Me.o_trans_data)
            p_freight_desc = fGetDesc("freight_cd", p_freight_cd, Me.o_error, Me.o_trans_data)
            p_backorder_desc = fGetDesc("backorder_cd", p_backorder_cd, Me.o_error, Me.o_trans_data)
            p_drop_ship_desc = fGetDesc("flag", p_drop_ship_fl, Me.o_error, Me.o_trans_data)
            Call GetTotals
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
Private Sub GetTotals()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Dim objRecordset As ADODB.Recordset
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim lc_net_amnt As Variant
    Dim lc_tax_amnt As Variant
    
    
    strRoutine_nm = "Purchase_order.cls.GetTotals"
    strSql = "SELECT ordered_qty, item_cost FROM purchase_order_item " & _
             " WHERE " & fWhereVariable("po_id", p_po_id, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)

    p_gross_amnt = 0
    p_net_amnt = 0
    p_line_item_qty = 0
    
    Do Until objRecordset.EOF = True Or Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F"
      p_line_item_qty = p_line_item_qty + objRecordset("ordered_qty")
      lc_net_amnt = objRecordset("item_cost") + objRecordset("ordered_qty")
      p_gross_amnt = p_gross_amnt + lc_net_amnt
        objRecordset.MoveNext
    Loop

    lc_tax_amnt = p_gross_amnt * p_sales_tax_pct
    p_net_amnt = p_gross_amnt + lc_tax_amnt

    objRecordset.Close
    Set objRecordset = Nothing
    
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
'Description:   Function to change the purchase order status
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub ChangeStatus()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order.cls.changeStatus"
    
    Dim dblUpdatedQty As Double, _
        objPurchaseOrderItem As Object, _
        objRecordset As ADODB.Recordset, _
        strSql As String

    If p_status_cd = "PL" Or _
       p_status_cd = "OP" Or _
       p_status_cd = "ST" Or _
       p_status_cd = "CL" Or _
       p_status_cd = "CA" Or _
       p_status_cd = "NC" Or _
       p_status_cd = "RS" Or _
       p_status_cd = "OH" Then
    Else
        Me.o_error.p_type_cd = "F"
        Me.o_error.p_err_cd = "1300"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2010
        Exit Sub
    End If
    
    Set Me.o_old_purchase_order = New purchase_order
    Set Me.o_old_purchase_order.o_trans_data = Me.o_trans_data
    Set Me.o_old_purchase_order.o_error = Me.o_error
    Me.o_old_purchase_order.p_po_id = Me.p_po_id
    Me.o_old_purchase_order.p_bypass_nm_addr_fl = "N"
    Me.o_old_purchase_order.p_bypass_desc_fl = "Y"
    Me.o_old_purchase_order.getPurchaseOrder
    
    If IsNull(Me.o_old_purchase_order.p_po_id) = True Then
        Me.o_error.p_type_cd = "F"
        Me.o_error.p_err_cd = "0800"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2136
        Exit Sub
    End If
        
    If Me.o_old_purchase_order.p_status_cd = "PL" Then
        If Me.p_status_cd = "PL" Or Me.p_status_cd = "ST" Or _
           Me.p_status_cd = "OP" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1410"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2137
            Exit Sub
        End If
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "OP" Then
        If Me.p_status_cd = "OP" Or _
           Me.p_status_cd = "ST" Or _
           Me.p_status_cd = "CL" Or _
           Me.p_status_cd = "CA" Or _
           Me.p_status_cd = "OH" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1420"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2138
            Exit Sub
        End If
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "ST" Then
        If Me.p_status_cd = "CL" Or _
           Me.p_status_cd = "CA" Or _
           Me.p_status_cd = "NC" Or _
           Me.p_status_cd = "RS" Then
            If Me.p_status_cd = "RS" Then
                If Me.p_move_to_hist_fl <> "N" Then
                    Call MoveToHistory
                End If
            End If
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1430"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2139
            Exit Sub
        End If
    End If
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "CL" Then
        If Me.p_status_cd = "ST" Or _
           Me.p_status_cd = "OP" Then
            Me.p_status_cd = "ST"
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1440"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2140
            Exit Sub
        End If
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "CA" Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1450"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2141
        Exit Sub
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "NC" Then
        If Me.p_status_cd = "RS" Or _
           Me.p_status_cd = "CA" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1460"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2142
            Exit Sub
        End If
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "RS" Then
        If Me.p_status_cd = "RS" Or _
           Me.p_status_cd = "CA" Or _
           Me.p_status_cd = "ST" Or _
           Me.p_status_cd = "NC" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1470"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2143
            Exit Sub
        End If
    End If
    
    If Me.o_old_purchase_order.p_status_cd = "OH" Then
        If Me.p_status_cd = "OP" Or _
           Me.p_status_cd = "CA" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1480"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 2144
            Exit Sub
        End If
    End If
    
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_po_id _
                , Me.p_status_cd _
                , Me.o_old_purchase_order _
                )
    
    Call SelectPurchaseOrder
    Call getPurchaseOrderItems(objRecordset)
    
    Set objPurchaseOrderItem = New purchase_order_item
    Set objPurchaseOrderItem.o_error = Me.o_error
    Set objPurchaseOrderItem.o_trans_data = Me.o_trans_data
    
    Do While objRecordset.EOF = False And Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
        If Me.p_status_cd = "CA" And objRecordset("received_qty") > 0 Then
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "1470"
                .p_routine_nm = strRoutine_nm
                .p_desc = "Can't cancel the order, one of the line items has recieved inventory, Close instead"
                .p_message_id = 0
            End With
        End If
        If objRecordset("status_cd") <> "CA" Then
            With objPurchaseOrderItem
                .p_po_id = Me.p_po_id
                .p_line_nbr = objRecordset("line_nbr")
                .p_item_id = objRecordset("item_id")
                .p_status_cd = Me.p_status_cd
                .p_received_qty = objRecordset("received_qty")
                .p_ordered_qty = objRecordset("ordered_qty")
                Call .ChangeStatus
            End With
        End If
        objRecordset.MoveNext
    Loop
    
    Set objPurchaseOrderItem = Nothing
    objRecordset.Close
    Set objRecordset = Nothing
    
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

Public Sub SetStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblPoId As Double _
            )
    
    On Error GoTo errHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order.cls.SetStatus"
    
    Dim objRecordset As ADODB.Recordset, _
        objPurchaseOrderItem As purchase_order_item, _
        strSql As String, _
        strStatusCd As String, _
        strChange_cd As String

    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_po_id = v_dblPoId
    Me.p_bypass_nm_addr_fl = "N"
    Me.p_bypass_desc_fl = "Y"
    Call Me.getPurchaseOrder
      
    If Me.p_status_cd = "PL" Or _
       Me.p_status_cd = "OP" Or _
       Me.p_status_cd = "CA" Or _
       Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "OH" Then
        Exit Sub
    End If
    
    strSql = "select * from purchase_order_item " & _
             " where " & fWhereVariable("po_id", v_dblPoId, "N", "=") & _
             " and   " & fWhereVariable("status_cd", "CP", "S", "<>") & _
             " and   " & fWhereVariable("status_cd", "CA", "S", "<>") & _
             " and   " & fWhereVariable("status_cd", "CL", "S", "<>")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = False Then
       strStatusCd = "ST"
    Else
       strStatusCd = "CP"
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Call IssueUpdateStatus( _
                      r_objError _
                    , r_objTransData _
                    , v_dblPoId _
                    , strStatusCd _
                    , Me _
                    )
                    
                    'POI call PO set status, never call POI set status
                    'from PO set status.  It could cause an infinite loop
'    Call getPurchaseOrderItems(objRecordSet)
'
'    Set objPurchaseOrderItem = New purchase_order_item
'    Set objPurchaseOrderItem.o_error = r_objError
'    Set objPurchaseOrderItem.o_trans_data = r_objTransData
'
'    Do While objRecordSet.EOF = False And Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
'        objPurchaseOrderItem.p_po_id = p_po_id
'        objPurchaseOrderItem.p_line_nbr = objRecordSet("line_nbr")
'        objPurchaseOrderItem.p_status_cd = strStatusCd
'        Call objPurchaseOrderItem.ChangeStatus
'        objRecordSet.MoveNext
'    Loop
'
'    Set objPurchaseOrderItem = Nothing
'
'    objRecordSet.Close
'    Set objRecordSet = Nothing
    
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

Public Sub GetPurchaseOrders()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_orders.cls.getPurchaseOrders"
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    Dim strPOStatus As String
        
    strPOStatus = "purchase_order_status_cd"
    
    strSql = "SELECT  " & _
                " a.opened_dt, " & _
                " a.entered_dt, " & _
                " b.supplier_nbr, " & _
                " c.code_desc as status_desc, " & _
                " d.code_desc as type_desc, " & _
                " a.po_id, " & _
                " a.drop_ship_fl, " & _
                " a.order_tx, " & _
                " a.conv_from_ord_id " & _
                " FROM purchase_order a, supplier b, code_desc c, code_desc d " & _
                " WHERE a.supplier_id = b.supplier_id " & _
                " AND (" & fWhereVariable("c.field_nm", "purchase_order_status_cd", "S", "=") & _
               " AND c.code = a.status_cd)" & _
                " AND (" & fWhereVariable("d.field_nm", "purchase_order_type_cd", "S", "=") & _
               " AND d.code = a.type_cd)"
                
    p_supplier_id = fGetSupplierId(p_supplier_nbr, Me.o_error, Me.o_trans_data)
    If fValuePresent(p_supplier_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_id", p_supplier_id, "N", "=")
    End If
    
   If fValuePresent(p_status_cd) = True Then
        If fValuePresent(p_status_cd2) = True Then
            If fValuePresent(p_status_cd3) = True Then
                If fValuePresent(p_status_cd4) = True Then
                    strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd2, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd3, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd4, "S", "=")
                    strSql = strSql & ")"
                Else
                    strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd2, "S", "=")
                    strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd3, "S", "=")
                    strSql = strSql & ")"
                End If
            Else
                strSql = strSql & "  AND (" & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
                strSql = strSql & "  OR " & fWhereVariable("a.status_cd", p_status_cd2, "S", "=")
                strSql = strSql & ")"
            End If
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
        End If
    End If
    
    If fValuePresent(p_drop_ship_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_sales_ord_id", p_drop_ship_sales_ord_id, "N", "=")
    End If
    
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
    
    If fValuePresent(p_conv_from_ord_id) = True Then
        strSql = strSql & " AND " & fWhereVariable("a.conv_from_ord_id", p_conv_from_ord_id, "N", "=")
    End If
    
    If fValuePresent(p_start_dt) = True Then
       If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("a.opened_dt", p_start_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("a.opened_dt", p_end_dt, "D", "<=") & ")"
       Else
            strSql = strSql & "  AND " & fWhereVariable("a.opened_dt", p_start_dt, "D", ">=")
       End If
    End If
    
    strSql = strSql & " ORDER BY a.po_id"
    
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
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


Private Sub GetDefaultValues(ByVal varItemID As Variant, ByRef objPurchaseOrderItem As Object)
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim objInvLoc As Object
    Dim objItem As Object
    strRoutine_nm = "purchase_orders.cls.getDefaultValues"
    
    Set objItem = CreateObject("converge_item.item")
    Set objItem.o_error = Me.o_error
    Set objItem.o_trans_data = Me.o_trans_data
    objItem.p_item_id = varItemID
    objItem.p_item_nbr = Null
    Call objItem.GetItem
    
    If IsNull(objItem.p_item_id) = True Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1500"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2090
        Exit Sub
    End If
    
    If IsNull(p_supplier_id) Then
        p_supplier_id = objItem.p_preferred_supplier_id
    End If
    
    Dim objSupplier As Object
    Set objSupplier = CreateObject("converge_purchase.supplier")
    Set objSupplier.o_error = Me.o_error
    Set objSupplier.o_trans_data = Me.o_trans_data
    
    objSupplier.p_supplier_id = p_supplier_id
    objSupplier.p_supplier_nbr = Null
    Call objSupplier.GetSupplier
    
    If IsNull(objSupplier.p_supplier_id) = True Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1610"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 2100
        Exit Sub
    End If
    
    Me.p_order_text = objSupplier.p_order_text
    
    Set objInvLoc = CreateObject("converge_inv.inventory_loc")
    Set objInvLoc.o_error = Me.o_error
    Set objInvLoc.o_trans_data = Me.o_trans_data
    Call objInvLoc.GetDefaultInventoryLoc


    Set o_stnd_name = CreateObject("converge.name")
    Set o_stnd_name = objSupplier.o_stnd_name
    Set o_stnd_address = CreateObject("converge.address")
    Set o_stnd_address = objSupplier.o_stnd_address
   
    
    If Len(objInvLoc.o_ship_name.p_nm) > 0 Then
        Set o_ship_name = CreateObject("converge.name")
      
        Set o_ship_name = objInvLoc.o_ship_name
        Set o_ship_address = CreateObject("converge.address")
        Set o_ship_address = objInvLoc.o_ship_address
        o_ship_name.p_nm_type_cd = "SHIP"
        o_ship_address.p_addr_type_cd = "SHIP"
    Else
        Set o_ship_name = CreateObject("converge.name")
        Set o_ship_name = objInvLoc.o_stnd_name
        Set o_ship_address = CreateObject("converge.address")
        Set o_ship_address = objInvLoc.o_stnd_address
        o_ship_name.p_nm_type_cd = "SHIP"
        o_ship_address.p_addr_type_cd = "SHIP"
    End If
    
    If Len(objInvLoc.o_bill_name.p_nm) > 0 Then
        Set o_bill_name = CreateObject("converge.name")
        Set o_bill_name = objInvLoc.o_bill_name
        Set o_bill_address = CreateObject("converge.address")
        Set o_bill_address = objInvLoc.o_bill_address
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_address.p_addr_type_cd = "BILL"
    Else
        Set o_bill_name = CreateObject("converge.name")
        Set o_bill_name = objInvLoc.o_stnd_name
        Set o_bill_address = CreateObject("converge.address")
        Set o_bill_address = objInvLoc.o_stnd_address
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_address.p_addr_type_cd = "BILL"
    End If
        
    Dim objSupplierItem As Object
    Set objSupplierItem = CreateObject("converge_purchase.supplier_item")
    Set objSupplierItem.o_error = Me.o_error
    Set objSupplierItem.o_trans_data = Me.o_trans_data
    objSupplierItem.p_item_id = varItemID
    
    objSupplierItem.GetSupplierItem
    
    If IsNull(objSupplierItem.p_item_id) Then
        ' Set the purchase Order Item Properties
        objPurchaseOrderItem.p_item_cost = objItem.p_item_cost
        objPurchaseOrderItem.p_item_desc = objItem.p_item_desc
        objPurchaseOrderItem.p_supplier_item_nbr = Null
    Else
        objPurchaseOrderItem.p_item_cost = objSupplierItem.p_item_cost
        objPurchaseOrderItem.p_supplier_item_nbr = objSupplierItem.p_supplier_item_nbr
        If objSupplierItem.p_item_desc = Empty Then
           objPurchaseOrderItem.p_item_desc = objItem.p_item_desc
        Else
            objPurchaseOrderItem.p_item_desc = objSupplierItem.p_item_desc
        End If
    End If
    
    p_fob_terms_cd = objSupplier.p_fob_terms_cd
    p_currency_cd = objSupplier.p_currency_cd
    p_comm_meth_cd = objSupplier.p_comm_meth_cd
    p_freight_cd = objSupplier.p_freight_cd
    p_payment_terms_cd = objSupplier.p_payment_terms_cd
        
    If Len(p_payment_terms_cd) = 0 Then
        p_payment_terms_cd = "01"
    End If
        
    objPurchaseOrderItem.p_item_id = objItem.p_item_id
    objPurchaseOrderItem.p_item_cost = objItem.p_item_cost
    objPurchaseOrderItem.p_uom_cd = objItem.p_uom_cd
    objPurchaseOrderItem.p_item_text = objItem.p_item_text
    
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

Private Sub getPurchaseOrderItems(ByRef r_objRecordset As ADODB.Recordset)

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "purchase_order.cls.getPurchaseOrderItems"
    
    strSql = "select * from purchase_order_item " & _
                " where" & _
                "   " & fWhereVariable("po_id", p_po_id, "N", "=")

    Call Me.o_trans_data.OpenRecordset(r_objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        r_objRecordset.Close
        Set r_objRecordset = Nothing
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
            , ByVal v_varStatusCd As Variant _
            , ByRef r_objOldPurchaseOrder As purchase_order _
            )

    On Error GoTo errHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order.cls.issueUpdateStatus"
    
    Dim dblUpdatedQty As Double, _
        strSql As String
    
    strSql = "update purchase_order set "
    strSql = strSql & fUpdateVariable("status_cd", v_varStatusCd, "S", "=")
    If v_varStatusCd = "OP" Then
        strSql = strSql & "," & fUpdateVariable("opened_dt", "current_ts", "D", "=")
    End If
    
    If v_varStatusCd = "ST" And IsDate(r_objOldPurchaseOrder.p_opened_dt) = False Then
        strSql = strSql & "," & fUpdateVariable("opened_dt", "current_ts", "D", "=")
    End If
    
    If r_objOldPurchaseOrder.p_status_cd = "ST" And v_varStatusCd = "RS" Then
        strSql = strSql & ", revision_nbr = revision_nbr + 1 "
    End If
    
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("po_id", v_dblPoId, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedQty, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    If dblUpdatedQty = 0 Then
        r_objError.p_type_cd = "F"
        r_objError.p_err_cd = "1800"
        r_objError.p_routine_nm = strRoutine_nm
        r_objError.p_message_id = 1005
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

Public Sub updateExistingPOs()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    Dim objPurchaseOrder As Object
    strRoutine_nm = "purchase_order.cls.updateExistingPOs"
    
    If o_old_purchase_order Is Nothing Then
        If p_po_id > 0 Then
            Set o_old_purchase_order = CreateObject("converge_purchase.purchase_order")
            Set o_old_purchase_order.o_trans_data = Me.o_trans_data
            Set o_old_purchase_order.o_error = Me.o_error
            o_old_purchase_order.p_po_id = p_po_id
            o_old_purchase_order.p_bypass_nm_addr_fl = "N"
            o_old_purchase_order.p_bypass_desc_fl = "Y"
            o_old_purchase_order.getPurchaseOrder
        Else
            Exit Sub
        End If
    End If
    
    If IsNull(o_old_purchase_order.p_po_id) = True Then
        Exit Sub
    End If
    
    Set o_old_purchase_order.o_error = Me.o_error
    Set o_old_purchase_order.o_trans_data = Me.o_trans_data

    Select Case o_old_purchase_order.p_status_cd
        Case "PL"
            o_old_purchase_order.Delete
        Case "RS"
            o_old_purchase_order.p_status_cd = "NC"
            o_old_purchase_order.ChangeStatus
        Case "OP"
            o_old_purchase_order.p_status_cd = "CA"
            o_old_purchase_order.ChangeStatus
        Case "OH"
            o_old_purchase_order.p_status_cd = "NC"
            o_old_purchase_order.ChangeStatus
        Case "ST"
            o_old_purchase_order.p_status_cd = "NC"
            o_old_purchase_order.ChangeStatus
        Case "CA"
        Case Else
            With Me.o_error
              .p_type_cd = "E"
              .p_err_cd = "2000"
              .p_nbr = Err.Number
              .p_desc = Err.Description
              .p_routine_nm = strRoutine_nm
              .p_message_id = 2010
            End With
            Exit Sub
    End Select
    
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

Public Sub getSentPurchaseOrders()
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_orders.cls.getSentPurchaseOrders"
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    Dim strPOStatus As String
        
    strPOStatus = "purchase_order_status_cd"
    
    strSql = "SELECT DISTINCT " & _
                " a.opened_dt, " & _
                " a.entered_dt, " & _
                " b.supplier_nbr, " & _
                " a.type_cd, " & _
                " c.code_desc as status_desc, " & _
                " a.po_id, " & _
                " a.order_tx, " & _
                " a.drop_ship_fl, a.drop_ship_sales_ord_id " & _
                " FROM purchase_order a, supplier b, code_desc c, shipment d " & _
               " WHERE (" & fWhereVariable("c.field_nm", "purchase_order_status_cd", "S", "=") & _
               " AND c.code = a.status_cd)" & _
                " AND a.supplier_id = b.supplier_id " & _
                " AND a.status_cd = 'ST' " & _
                " AND a.drop_ship_fl = 'Y' " & _
                " AND a.drop_ship_sales_ord_id not in (select sales_ord_id from shipment) "
        

    p_supplier_id = fGetSupplierId(p_supplier_nbr, Me.o_error, Me.o_trans_data)
    If fValuePresent(p_supplier_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_id", p_supplier_id, "N", "=")
    End If
    
    If fValuePresent(p_drop_ship_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_sales_ord_id", p_drop_ship_sales_ord_id, "N", "=")
    End If
    
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
    
    If fValuePresent(p_start_dt) = True Then
       If fValuePresent(p_end_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("a.opened_dt", p_start_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("a.opened_dt", p_end_dt, "D", "<=") & ")"
       Else
            strSql = strSql & "  AND " & fWhereVariable("a.opened_dt", p_start_dt, "D", ">=")
       End If
    End If
    
    strSql = strSql & " ORDER BY a.po_id"
    
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
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

Private Function getPOItems(ByVal v_strPO_ID) As Collection
    
    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "getPOItems"
    
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    Dim colPOItems As Collection
    
    strSql = "select * from purchase_order_item " & _
             "where po_id = " & v_strPO_ID

    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    Set colPOItems = New Collection
    Do Until objRecordset.EOF = True Or Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F"
        Dim objPurchaseOrderItem As Object
        Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
        objPurchaseOrderItem.p_po_id = p_po_id
        objPurchaseOrderItem.p_line_nbr = objRecordset("line_nbr")
        objPurchaseOrderItem.p_item_id = objRecordset("item_id")
        objPurchaseOrderItem.p_item_cost = objRecordset("item_cost")
        Call colPOItems.Add(objPurchaseOrderItem, CStr(objPurchaseOrderItem.p_line_nbr))
        objRecordset.MoveNext
    Loop
    Set getPOItems = colPOItems
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


