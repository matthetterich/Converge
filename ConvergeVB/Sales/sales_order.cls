VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_objError As Object
Private m_objTrans_data As Object
Private m_objBill_name As Object
Private m_objBill_address As Object
Private m_objShip_address As Object
Private m_objShip_name As Object
Private m_objOldSalesOrder As sales_order
Private m_varSales_ord_id As Variant
Private m_varBlanket_sales_ord_id As Variant
Private m_varBypass_nm_addr_fl As Variant
Private m_varBypass_desc_fl As Variant
Private m_varCarrier_nm As Variant
Private m_varContract_nbr As Variant
Private m_varCurrency_cd As Variant
Private m_varCurrency_desc As Variant
Private m_varCust_id As Variant
Private m_varCust_nbr As Variant
Private m_varCust_po_nbr As Variant
Private m_varConsignment_fl As Variant
Private m_varDiscount_pct As Variant
Private m_varDiscount_am As Variant
Private m_varDrop_ship_fl As Variant
Private m_varDrop_ship_supplier_id As Variant
Private m_varDrop_ship_supplier_nbr As Variant
Private m_varDrop_ship_po_id As Variant
Private m_varFob_terms_cd As Variant
Private m_varFob_terms_desc As Variant
Private m_varFreight_cd As Variant
Private m_varFreight_desc As Variant
Private m_varGross_am As Variant
Private m_varInv_loc_id As Variant
Private m_varInv_loc_nm As Variant
Private m_varNet_am As Variant
Private m_varOrdered_dt As Variant
Private m_varPayment_terms_cd As Variant
Private m_varPayment_terms_desc As Variant
Private m_varPlanned_del_dt As Variant
Private m_varprice_calc_meth_cd As Variant
Private m_varprice_calc_meth_desc As Variant
Private m_varRevision_nbr As Variant
Private m_varStatus_cd As Variant
Private m_varStatus_cd2 As Variant
Private m_varStatus_cd3 As Variant
Private m_varStatus_cd4 As Variant
Private m_varStatus_desc As Variant
Private m_varSales_quote_id As Variant
Private m_varSales_rep_assoc_id As Variant
Private m_varSales_rep_assoc_nbr As Variant
Private m_varSales_tax_am As Variant
Private m_varSales_tax_pct As Variant
Private m_varTransportation_cd As Variant
Private m_varTransportation_desc As Variant
Private m_varType_cd As Variant
Private m_varType_desc As Variant
Private m_varUpdated_assoc_id As Variant
Private m_varUpdated_ts As Variant
Private m_varEnter_assoc_id As Variant
Private m_varEntered_ts As Variant
Private mvaro_record_set As adodb.Recordset
Private mvarp_conv_from_ord_id As Variant
Private mvarp_order_text As Variant
'

Public Property Let p_order_text(ByVal vData As Variant)
    mvarp_order_text = vData
End Property
Public Property Get p_order_text() As Variant
    p_order_text = mvarp_order_text
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Set o_old_sales_order(ByVal vData As sales_order)
    Set m_objOldSalesOrder = vData
End Property
Public Property Get o_old_sales_order() As sales_order
        Set o_old_sales_order = m_objOldSalesOrder
End Property

Public Property Let p_inv_loc_id(ByVal vData As Variant)
    m_varInv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = m_varInv_loc_id
End Property

Public Property Let p_conv_from_ord_id(ByVal vData As Variant)
    mvarp_conv_from_ord_id = vData
End Property
Public Property Get p_conv_from_ord_id() As Variant
        p_conv_from_ord_id = mvarp_conv_from_ord_id
End Property

Public Property Let p_inv_loc_nm(ByVal vData As Variant)
    m_varInv_loc_nm = vData
End Property
Public Property Get p_inv_loc_nm() As Variant
        p_inv_loc_nm = m_varInv_loc_nm
End Property

Public Property Let p_bypass_nm_addr_fl(ByVal vData As Variant)
    m_varBypass_nm_addr_fl = vData
End Property
Public Property Get p_bypass_nm_addr_fl() As Variant
        p_bypass_nm_addr_fl = m_varBypass_nm_addr_fl
End Property

Public Property Let p_bypass_desc_fl(ByVal vData As Variant)
    m_varBypass_desc_fl = vData
End Property
Public Property Get p_bypass_desc_fl() As Variant
        p_bypass_desc_fl = m_varBypass_desc_fl
End Property

Public Property Let p_cust_nbr(ByVal vData As Variant)
    m_varCust_nbr = vData
End Property
Public Property Get p_cust_nbr() As Variant
    p_cust_nbr = m_varCust_nbr
End Property
Public Property Set o_ship_name(ByVal vData As Object)
    Set m_objShip_name = vData
End Property
Public Property Get o_ship_name() As Object
    Set o_ship_name = m_objShip_name
End Property
Public Property Set o_ship_address(ByVal vData As Object)
    Set m_objShip_address = vData
End Property
Public Property Get o_ship_address() As Object
    Set o_ship_address = m_objShip_address
End Property
Public Property Set o_bill_address(ByVal vData As Object)
    Set m_objBill_address = vData
End Property
Public Property Get o_bill_address() As Object
    Set o_bill_address = m_objBill_address
End Property
Public Property Set o_bill_name(ByVal vData As Object)
    Set m_objBill_name = vData
End Property
Public Property Get o_bill_name() As Object
    Set o_bill_name = m_objBill_name
End Property
Public Property Let p_blanket_sales_ord_id(ByVal vData As Variant)
    m_varBlanket_sales_ord_id = vData
End Property
Public Property Get p_blanket_sales_ord_id() As Variant
    p_blanket_sales_ord_id = m_varBlanket_sales_ord_id
End Property
Public Property Let p_transportation_desc(ByVal vData As Variant)
    m_varTransportation_desc = vData
End Property
Public Property Get p_transportation_desc() As Variant
    p_transportation_desc = m_varTransportation_desc
End Property
Public Property Let p_transportation_cd(ByVal vData As Variant)
    m_varTransportation_cd = vData
End Property
Public Property Get p_transportation_cd() As Variant
    p_transportation_cd = m_varTransportation_cd
End Property
Public Property Let p_sales_tax_am(ByVal vData As Variant)
    m_varSales_tax_am = vData
End Property
Public Property Get p_sales_tax_am() As Variant
    p_sales_tax_am = m_varSales_tax_am
End Property
Public Property Let p_sales_rep_assoc_nbr(ByVal vData As Variant)
    m_varSales_rep_assoc_nbr = vData
End Property
Public Property Get p_sales_rep_assoc_nbr() As Variant
    p_sales_rep_assoc_nbr = m_varSales_rep_assoc_nbr
End Property
Public Property Let p_price_calc_meth_desc(ByVal vData As Variant)
    m_varprice_calc_meth_desc = vData
End Property
Public Property Get p_price_calc_meth_desc() As Variant
    p_price_calc_meth_desc = m_varprice_calc_meth_desc
End Property
Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    m_varPayment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
    p_payment_terms_desc = m_varPayment_terms_desc
End Property
Public Property Let p_net_am(ByVal vData As Variant)
    m_varNet_am = vData
End Property
Public Property Get p_net_am() As Variant
    p_net_am = m_varNet_am
End Property
Public Property Let p_gross_am(ByVal vData As Variant)
    m_varGross_am = vData
End Property
Public Property Get p_gross_am() As Variant
    p_gross_am = m_varGross_am
End Property
Public Property Let p_freight_desc(ByVal vData As Variant)
    m_varFreight_desc = vData
End Property
Public Property Get p_freight_desc() As Variant
    p_freight_desc = m_varFreight_desc
End Property
Public Property Let p_fob_terms_desc(ByVal vData As Variant)
    m_varFob_terms_desc = vData
End Property
Public Property Get p_fob_terms_desc() As Variant
    p_fob_terms_desc = m_varFob_terms_desc
End Property
Public Property Let p_discount_am(ByVal vData As Variant)
    m_varDiscount_am = vData
End Property
Public Property Get p_discount_am() As Variant
    p_discount_am = m_varDiscount_am
End Property
Public Property Let p_currency_desc(ByVal vData As Variant)
    m_varCurrency_desc = vData
End Property
Public Property Get p_currency_desc() As Variant
    p_currency_desc = m_varCurrency_desc
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    m_varStatus_desc = vData
End Property
Public Property Get p_status_desc() As Variant
    p_status_desc = m_varStatus_desc
End Property
Public Property Let p_entered_ts(ByVal vData As Variant)
    m_varEntered_ts = vData
End Property
Public Property Get p_entered_ts() As Variant
    p_entered_ts = m_varEntered_ts
End Property
Public Property Let p_enter_assoc_id(ByVal vData As Variant)
    m_varEnter_assoc_id = vData
End Property
Public Property Get p_enter_assoc_id() As Variant
    p_enter_assoc_id = m_varEnter_assoc_id
End Property
Public Property Let p_updated_ts(ByVal vData As Variant)
    m_varUpdated_ts = vData
End Property
Public Property Get p_updated_ts() As Variant
    p_updated_ts = m_varUpdated_ts
End Property
Public Property Let p_updated_assoc_id(ByVal vData As Variant)
    m_varUpdated_assoc_id = vData
End Property
Public Property Get p_updated_assoc_id() As Variant
    p_updated_assoc_id = m_varUpdated_assoc_id
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
Public Property Let p_sales_tax_pct(ByVal vData As Variant)
    m_varSales_tax_pct = vData
End Property
Public Property Get p_sales_tax_pct() As Variant
    p_sales_tax_pct = m_varSales_tax_pct
End Property
Public Property Let p_sales_rep_assoc_id(ByVal vData As Variant)
    m_varSales_rep_assoc_id = vData
End Property
Public Property Get p_sales_rep_assoc_id() As Variant
    p_sales_rep_assoc_id = m_varSales_rep_assoc_id
End Property
Public Property Let p_sales_quote_id(ByVal vData As Variant)
    m_varSales_quote_id = vData
End Property
Public Property Get p_sales_quote_id() As Variant
    p_sales_quote_id = m_varSales_quote_id
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    m_varStatus_cd = vData
End Property
Public Property Get p_status_cd() As Variant
    p_status_cd = m_varStatus_cd
End Property
Public Property Let p_status_cd2(ByVal vData As Variant)
    m_varStatus_cd2 = vData
End Property
Public Property Get p_status_cd2() As Variant
    p_status_cd2 = m_varStatus_cd2
End Property
Public Property Let p_status_cd3(ByVal vData As Variant)
    m_varStatus_cd3 = vData
End Property
Public Property Get p_status_cd3() As Variant
    p_status_cd3 = m_varStatus_cd3
End Property
Public Property Let p_status_cd4(ByVal vData As Variant)
    m_varStatus_cd4 = vData
End Property
Public Property Get p_status_cd4() As Variant
    p_status_cd4 = m_varStatus_cd4
End Property

Public Property Let p_revision_nbr(ByVal vData As Variant)
    m_varRevision_nbr = vData
End Property
Public Property Get p_revision_nbr() As Variant
    p_revision_nbr = m_varRevision_nbr
End Property
Public Property Let p_price_calc_meth_cd(ByVal vData As Variant)
    m_varprice_calc_meth_cd = vData
End Property
Public Property Get p_price_calc_meth_cd() As Variant
    p_price_calc_meth_cd = m_varprice_calc_meth_cd
End Property
Public Property Let p_planned_del_dt(ByVal vData As Variant)
    m_varPlanned_del_dt = vData
End Property
Public Property Get p_planned_del_dt() As Variant
    p_planned_del_dt = m_varPlanned_del_dt
End Property
Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    m_varPayment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
    p_payment_terms_cd = m_varPayment_terms_cd
End Property
Public Property Let p_ordered_dt(ByVal vData As Variant)
    m_varOrdered_dt = vData
End Property
Public Property Get p_ordered_dt() As Variant
    p_ordered_dt = m_varOrdered_dt
End Property
Public Property Let p_freight_cd(ByVal vData As Variant)
    m_varFreight_cd = vData
End Property
Public Property Get p_freight_cd() As Variant
    p_freight_cd = m_varFreight_cd
End Property
Public Property Let p_fob_terms_cd(ByVal vData As Variant)
    m_varFob_terms_cd = vData
End Property
Public Property Get p_fob_terms_cd() As Variant
    p_fob_terms_cd = m_varFob_terms_cd
End Property
Public Property Let p_drop_ship_fl(ByVal vData As Variant)
    m_varDrop_ship_fl = vData
End Property
Public Property Get p_drop_ship_fl() As Variant
    p_drop_ship_fl = m_varDrop_ship_fl
End Property
Public Property Let p_drop_ship_supplier_id(ByVal vData As Variant)
    m_varDrop_ship_supplier_id = vData
End Property
Public Property Get p_drop_ship_supplier_id() As Variant
    p_drop_ship_supplier_id = m_varDrop_ship_supplier_id
End Property
Public Property Let p_drop_ship_supplier_nbr(ByVal vData As Variant)
    m_varDrop_ship_supplier_nbr = vData
End Property
Public Property Get p_drop_ship_supplier_nbr() As Variant
    p_drop_ship_supplier_nbr = m_varDrop_ship_supplier_nbr
End Property
Public Property Let p_drop_ship_po_id(ByVal vData As Variant)
    m_varDrop_ship_po_id = vData
End Property
Public Property Get p_drop_ship_po_id() As Variant
    p_drop_ship_po_id = m_varDrop_ship_po_id
End Property
Public Property Let p_discount_pct(ByVal vData As Variant)
    m_varDiscount_pct = vData
End Property
Public Property Get p_discount_pct() As Variant
    p_discount_pct = m_varDiscount_pct
End Property
Public Property Let p_consignment_fl(ByVal vData As Variant)
    m_varConsignment_fl = vData
End Property
Public Property Get p_consignment_fl() As Variant
    p_consignment_fl = m_varConsignment_fl
End Property
Public Property Let p_cust_po_nbr(ByVal vData As Variant)
    m_varCust_po_nbr = vData
End Property
Public Property Get p_cust_po_nbr() As Variant
    p_cust_po_nbr = m_varCust_po_nbr
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    m_varCust_id = vData
End Property
Public Property Get p_cust_id() As Variant
    p_cust_id = m_varCust_id
End Property
Public Property Let p_currency_cd(ByVal vData As Variant)
    m_varCurrency_cd = vData
End Property
Public Property Get p_currency_cd() As Variant
    p_currency_cd = m_varCurrency_cd
End Property
Public Property Let p_contract_nbr(ByVal vData As Variant)
    m_varContract_nbr = vData
End Property
Public Property Get p_contract_nbr() As Variant
    p_contract_nbr = m_varContract_nbr
End Property
Public Property Let p_carrier_nm(ByVal vData As Variant)
    m_varCarrier_nm = vData
End Property
Public Property Get p_carrier_nm() As Variant
    p_carrier_nm = m_varCarrier_nm
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

Public Sub getSalesOrder()
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
     
    On Error GoTo ErrHandler
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.GetsalesOrder"
    
    If IsNull(p_sales_ord_id) Then
        Exit Sub
    End If
    
    Call SelectSalesOrder
    
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_address = CreateObject("converge.address")
    Set o_ship_name = CreateObject("converge.name")
    Set o_ship_address = CreateObject("converge.address")
    
    If IsNull(p_sales_ord_id) = True Then
        o_bill_name.p_nm_id = Null
        o_bill_address.p_addr_id = Null
    Else
        Call CalcValues
        If p_bypass_nm_addr_fl = "Y" Then
        Else
            Set o_bill_name.o_trans_data = o_trans_data
            Set o_bill_name.o_error = o_error
            o_bill_name.p_nm_id = Null
            o_bill_name.p_obj_nm = "sales_order"
            o_bill_name.p_obj_id = p_sales_ord_id
            o_bill_name.p_nm_type_cd = "BILL"
            o_bill_name.p_def_nm_fl = "Y"
            Call o_bill_name.getDefName
                          
            Set o_bill_address.o_trans_data = o_trans_data
            Set o_bill_address.o_error = o_error
            o_bill_address.p_addr_id = Null
            o_bill_address.p_obj_nm = "sales_order"
            o_bill_address.p_obj_id = p_sales_ord_id
            o_bill_address.p_addr_type_cd = "BILL"
            o_bill_address.p_def_addr_fl = "Y"
            Call o_bill_address.getDefAddress
            
            Set o_ship_name.o_trans_data = o_trans_data
            Set o_ship_name.o_error = o_error
            o_ship_name.p_nm_id = Null
            o_ship_name.p_obj_nm = "sales_order"
            o_ship_name.p_obj_id = p_sales_ord_id
            o_ship_name.p_nm_type_cd = "SHIP"
            o_ship_name.p_def_nm_fl = "Y"
            Call o_ship_name.getDefName
                          
            Set o_ship_address.o_trans_data = o_trans_data
            Set o_ship_address.o_error = o_error
            o_ship_address.p_addr_id = Null
            o_ship_address.p_obj_nm = "sales_order"
            o_ship_address.p_obj_id = p_sales_ord_id
            o_ship_address.p_addr_type_cd = "SHIP"
            o_ship_address.p_def_addr_fl = "Y"
            Call o_ship_address.getDefAddress
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error GoTo ErrHandler
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.Add"
    
    Call validate("add")
    Call IssueInsertData
    
    Call AddName(o_bill_name)
    Call AddAddr(o_bill_address)
    Call AddName(o_ship_name)
    Call AddAddr(o_ship_address)
    
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.Change"
    
    If o_old_sales_order Is Nothing Then
        Set o_old_sales_order = CreateObject("converge_sales.sales_order")
        Set o_old_sales_order.o_error = o_error
        Set o_old_sales_order.o_trans_data = o_trans_data
        o_old_sales_order.p_sales_ord_id = p_sales_ord_id
        o_old_sales_order.p_bypass_nm_addr_fl = "N"
        o_old_sales_order.p_bypass_desc_fl = "Y"
        o_old_sales_order.p_sales_ord_id = p_sales_ord_id
        o_old_sales_order.getSalesOrder
    End If
    
    Call validate("change")
    Call IssueUpdateData
    
    If p_bypass_nm_addr_fl = "Y" Then
    Else
               
        If IsNull(o_bill_name.p_nm_id) = True Or o_bill_name.p_nm_id = Empty Then
            If IsNull(o_bill_name.p_nm) = True Or o_bill_name.p_nm = Empty Then
            Else
                Call AddName(o_bill_name)
            End If
        Else
            If IsNull(o_bill_name.p_nm) = True Then
                Call DeleteName(o_bill_name)
            Else
                Call ChangeName(o_bill_name)
            End If
        End If
        
        If IsNull(o_bill_address.p_addr_id) = True Or o_bill_address.p_addr_id = Empty Then
            If (IsNull(o_bill_address.p_Street_addr) = True Or o_bill_address.p_Street_addr = Empty) And _
               (IsNull(o_bill_address.p_City_nm) = True Or o_bill_address.p_City_nm = Empty) And _
               (IsNull(o_bill_address.p_Zip_cd) = True Or o_bill_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_bill_address)
            End If
        Else
            If IsNull(o_bill_address.p_Street_addr) = True And _
               IsNull(o_bill_address.p_City_nm) = True And _
               IsNull(o_bill_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_bill_address)
            Else
                Call ChangeAddr(o_bill_address)
            End If
        End If
        
         If IsNull(o_ship_name.p_nm_id) = True Or o_ship_name.p_nm_id = Empty Then
            If IsNull(o_ship_name.p_nm) = True Or o_ship_name.p_nm = Empty Then
            Else
                Call AddName(o_ship_name)
            End If
        Else
            If IsNull(o_ship_name.p_nm) = True Then
                Call DeleteName(o_ship_name)
            Else
                Call ChangeName(o_ship_name)
            End If
        End If
        
        If IsNull(o_ship_address.p_addr_id) = True Or o_ship_address.p_addr_id = Empty Then
            If (IsNull(o_ship_address.p_Street_addr) = True Or o_ship_address.p_Street_addr = Empty) And _
               (IsNull(o_ship_address.p_City_nm) = True Or o_ship_address.p_City_nm = Empty) And _
               (IsNull(o_ship_address.p_Zip_cd) = True Or o_ship_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_ship_address)
            End If
        Else
            If IsNull(o_ship_address.p_Street_addr) = True And _
               IsNull(o_ship_address.p_City_nm) = True And _
               IsNull(o_ship_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_ship_address)
            Else
                Call ChangeAddr(o_ship_address)
            End If
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd <> "PL" Then
        Call Me.incrementRevision(Me.o_error, Me.o_trans_data, Me.p_sales_ord_id, Me.o_old_sales_order, Me.o_trans_data.assoc_id, Now, "Changed Order")
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.Delete"
    
    If o_old_sales_order Is Nothing Then
        Set o_old_sales_order = CreateObject("converge_sales.sales_order")
        Set o_old_sales_order.o_error = o_error
        Set o_old_sales_order.o_trans_data = o_trans_data
        o_old_sales_order.p_sales_ord_id = p_sales_ord_id
        o_old_sales_order.p_bypass_nm_addr_fl = "N"
        o_old_sales_order.p_bypass_desc_fl = "Y"
        o_old_sales_order.p_sales_ord_id = p_sales_ord_id
        o_old_sales_order.getSalesOrder
    End If
    
    Call validate("delete")
    Call IssueDeleteData
    Call IssueDeleteLineItem
    Call IssueDeleteName
    Call IssueDeleteAddr

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

Public Sub AddSubSalesOrder()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim objSalesOrderItem As sales_order_item
    Dim strSql As String

    strRoutine_nm = "sales_order.cls.AddSubSalesOrder"
    
    Set o_old_sales_order = CreateObject("converge_sales.sales_order")
    Set o_old_sales_order.o_error = o_error
    Set o_old_sales_order.o_trans_data = o_trans_data
    o_old_sales_order.p_sales_ord_id = p_sales_ord_id
    o_old_sales_order.p_bypass_nm_addr_fl = "N"
    o_old_sales_order.p_bypass_desc_fl = "Y"
    o_old_sales_order.p_sales_ord_id = p_sales_ord_id
    o_old_sales_order.getSalesOrder
    
                                        ' Do not require the blanket order to be open
                                        ' Timco needs to generate sub sales orders for
                                        ' blanket orders in a planned status.  Opening
                                        ' the blanket order causes them to show up on
                                        ' reports that they should not.  So, until the
                                        ' blanket orders are excluded, PL status blanket
                                        ' orders can create sub sales orders
'    If o_old_sales_order.p_status_cd <> "OP" Then
'        With o_error
'          .p_type_cd = "E"
'          .p_err_cd = "1000"
'          .p_routine_nm = strRoutine_nm
'          .p_message_id = 4000
'        End With
'        Exit Sub
'    End If
    
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_address = CreateObject("converge.address")
    Set o_ship_name = CreateObject("converge.name")
    Set o_ship_address = CreateObject("converge.address")
    
    Set o_bill_name = o_old_sales_order.o_bill_name
    Set o_bill_address = o_old_sales_order.o_bill_address
    Set o_ship_name = o_old_sales_order.o_ship_name
    Set o_ship_address = o_old_sales_order.o_ship_address
    
    Me.p_sales_ord_id = Null
    Me.p_blanket_sales_ord_id = Me.o_old_sales_order.p_sales_ord_id
    Me.p_carrier_nm = Me.o_old_sales_order.p_carrier_nm
    Me.p_contract_nbr = Me.o_old_sales_order.p_contract_nbr
    Me.p_currency_cd = Me.o_old_sales_order.p_currency_cd
    Me.p_cust_id = Me.o_old_sales_order.p_cust_id
    Me.p_cust_nbr = Me.o_old_sales_order.p_cust_nbr
    Me.p_cust_po_nbr = Me.o_old_sales_order.p_cust_po_nbr
    Me.p_consignment_fl = Me.o_old_sales_order.p_consignment_fl
    Me.p_discount_pct = Me.o_old_sales_order.p_discount_pct
    Me.p_drop_ship_fl = Me.o_old_sales_order.p_drop_ship_fl
    Me.p_drop_ship_supplier_id = Me.o_old_sales_order.p_drop_ship_supplier_id
    Me.p_drop_ship_supplier_nbr = Me.o_old_sales_order.p_drop_ship_supplier_nbr
    Me.p_fob_terms_cd = Me.o_old_sales_order.p_fob_terms_cd
    Me.p_freight_cd = Me.o_old_sales_order.p_freight_cd
    Me.p_inv_loc_id = Me.o_old_sales_order.p_inv_loc_id
    Me.p_ordered_dt = Me.o_old_sales_order.p_ordered_dt
    Me.p_payment_terms_cd = Me.o_old_sales_order.p_payment_terms_cd
    Me.p_planned_del_dt = Me.o_old_sales_order.p_planned_del_dt
    Me.p_price_calc_meth_cd = Me.o_old_sales_order.p_price_calc_meth_cd
    Me.p_sales_quote_id = Me.o_old_sales_order.p_sales_quote_id
    Me.p_sales_rep_assoc_id = Me.o_old_sales_order.p_sales_rep_assoc_id
    Me.p_sales_rep_assoc_nbr = fGetAssocNbr(Me.p_sales_rep_assoc_id, Me.o_error, Me.o_trans_data)
    'Me.p_sales_rep_assoc_nbr = me.o_old_sales_order.p_sales_rep_assoc_nbr
    Me.p_sales_tax_pct = Me.o_old_sales_order.p_sales_tax_pct
    Me.p_transportation_cd = Me.o_old_sales_order.p_transportation_cd
    Me.p_type_cd = "SO"
    Me.p_order_text = Me.o_old_sales_order.p_order_text
    Call Me.Add
        
                                        'Also get the sales items in planned status
                                        'Timco needs to generate sub sales orders for
                                        'blanket orders in a planned status.  Leaving
                                        'them in a planned status will prevent them from
                                        'showing on any reports.  Blanket orders should be
                                        'excluded from these reports but they are not.  Matt Hetterich
'    strSql = "select * from sales_order_item " & _
'            " where" & _
'            "   " & fWhereVariable("sales_ord_id", o_old_sales_order.p_sales_ord_id, "N", "=") & _
'            " and " & fWhereVariable("status_cd", "OP", "S", "=")

    strSql = "select * from sales_order_item " & _
            " where" & _
            "   " & fWhereVariable("sales_ord_id", o_old_sales_order.p_sales_ord_id, "N", "=") & _
            " and (" & fWhereVariable("status_cd", "OP", "S", "=") & _
            " or " & fWhereVariable("status_cd", "PL", "S", "=") & ")"

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set objSalesOrderItem = New sales_order_item
    Set objSalesOrderItem.o_error = Me.o_error
    Set objSalesOrderItem.o_trans_data = Me.o_trans_data
    
    Do While objRecordset.EOF = False And Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
        objSalesOrderItem.p_sales_ord_id = p_sales_ord_id
        objSalesOrderItem.p_status_cd = "PL"
        objSalesOrderItem.p_cust_item_nbr = objRecordset("cust_item_nbr")
        objSalesOrderItem.p_item_id = objRecordset("item_id")
        objSalesOrderItem.p_item_desc = objRecordset("item_desc")
        objSalesOrderItem.p_item_price = objRecordset("item_price")
        objSalesOrderItem.p_item_type_cd = objRecordset("item_type_cd")
        objSalesOrderItem.p_ordered_qty = objRecordset("ordered_qty")
        objSalesOrderItem.p_payment_terms_cd = objRecordset("payment_terms_cd")
        objSalesOrderItem.p_requested_del_dt = objRecordset("requested_del_dt")
        objSalesOrderItem.p_planned_ship_dt = objRecordset("planned_ship_dt")
        objSalesOrderItem.p_ship_complete_var_pct = objRecordset("ship_complete_var_pct")
        objSalesOrderItem.p_supplier_id = objRecordset("supplier_id")
        objSalesOrderItem.p_type_cd = objRecordset("type_cd")
        objSalesOrderItem.p_uom_cd = objRecordset("uom_cd")
        objSalesOrderItem.p_item_tx = objRecordset("item_tx")
        Call objSalesOrderItem.Add
        objRecordset.MoveNext
    Loop
    
    Set objSalesOrderItem = Nothing
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
Public Sub AddDropShipPo()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objPurchaseOrder As Object
    
    strRoutine_nm = "converge_sales.addDropShipPo"
    
    Set o_old_sales_order = CreateObject("converge_sales.sales_order")
    Set o_old_sales_order.o_error = o_error
    Set o_old_sales_order.o_trans_data = o_trans_data
    o_old_sales_order.p_sales_ord_id = p_sales_ord_id
    o_old_sales_order.p_bypass_nm_addr_fl = "N"
    o_old_sales_order.p_bypass_desc_fl = "Y"
    o_old_sales_order.p_sales_ord_id = p_sales_ord_id
    o_old_sales_order.getSalesOrder
    
    Set objPurchaseOrder = CreateObject("converge_purchase.purchase_order")
    Set objPurchaseOrder.o_error = o_error
    Set objPurchaseOrder.o_trans_data = o_trans_data
    objPurchaseOrder.p_buyer_assoc_id = p_updated_assoc_id
    objPurchaseOrder.p_po_id = o_old_sales_order.p_drop_ship_po_id
    objPurchaseOrder.p_drop_ship_sales_ord_id = p_sales_ord_id
    objPurchaseOrder.p_order_text = o_old_sales_order.p_order_text
    objPurchaseOrder.PlanDropShipSo

    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
Public Sub setDropShipPoId()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objPurchaseOrder As Object
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "converge_sales.setDropShipPoId"
    
    strSql = "update sales_order set "
    strSql = strSql & "  " & fUpdateVariable("drop_ship_po_id", p_drop_ship_po_id, "N", "=")
    strSql = strSql & ", " & fUpdateVariable("status_cd", "PC", "S", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    If dblUpdateQty = 0 Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "1200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1010
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.changeStatus"
    
    Dim dblUpdatedQty As Double, _
        objSalesOrderItem As sales_order_item, _
        objRecordset As adodb.Recordset, _
        objRecordset2 As adodb.Recordset, _
        objPurchaseOrder As Object, _
        strSql As String
        
    If p_status_cd = "OP" Or _
       p_status_cd = "CL" Or _
       p_status_cd = "CA" Or _
       p_status_cd = "OH" Then
    Else
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
        Exit Sub
    End If
    
    If Me.o_old_sales_order Is Nothing Then
        Set Me.o_old_sales_order = New sales_order
        Set Me.o_old_sales_order.o_error = Me.o_error
        Set Me.o_old_sales_order.o_trans_data = Me.o_trans_data
        Me.o_old_sales_order.p_sales_ord_id = Me.p_sales_ord_id
        Me.o_old_sales_order.p_bypass_nm_addr_fl = "N"
        Me.o_old_sales_order.p_bypass_desc_fl = "Y"
        Me.o_old_sales_order.p_sales_ord_id = Me.p_sales_ord_id
        Call Me.o_old_sales_order.getSalesOrder
    End If
    
    If Me.o_old_sales_order.p_status_cd = "PL" Then
        If p_status_cd <> "OP" Then
            With Me.o_error
                .p_type_cd = "F"
                .p_err_cd = "1800"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 0
            End With
            Exit Sub
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd = "OP" Then
        If Me.p_status_cd = "OP" Or _
           Me.p_status_cd = "OH" Or _
           Me.p_status_cd = "CA" Or _
           Me.p_status_cd = "CL" Then
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "2000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 0
            Exit Sub
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd = "CP" Then
        If Me.p_status_cd = "OP" Or _
           Me.p_status_cd = "CL" Then
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "2000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 0
            Exit Sub
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd = "CL" Then
        If Me.p_status_cd = "CL" Or _
           Me.p_status_cd = "OP" Then
        Else
            With Me.o_error
                .p_type_cd = "F"
                .p_err_cd = "2400"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 0
            End With
            Exit Sub
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd = "OH" Then
        If p_status_cd = "OH" Or _
           p_status_cd = "OP" Or _
           p_status_cd = "CL" Or _
           p_status_cd = "CA" Then
        Else
            With Me.o_error
                .p_type_cd = "F"
                .p_err_cd = "2600"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 0
            End With
            Exit Sub
        End If
    End If
    
    If Me.o_old_sales_order.p_status_cd = "CA" Then
        If Me.p_status_cd <> "CA" Then
            With Me.o_error
                .p_type_cd = "F"
                .p_err_cd = "2800"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 0
            End With
            Exit Sub
        End If
    End If
    
    If p_status_cd = "CA" Then
        strSql = " SELECT sales_ord_id, SUM(shipped_qty) AS tot_shipped " & _
                 " From sales_order_item " & _
                 " GROUP BY sales_ord_id " & _
                 " Having " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
        
        Call Me.o_trans_data.OpenRecordset(objRecordset2, strSql, Me.o_error)
        
        objRecordset2.MoveFirst
        
        If objRecordset2("tot_shipped") > 0 Then
            Set objRecordset2 = Nothing
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "1800"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1200
            End With
            Exit Sub
        End If
    End If
    
    Call IssueUpdateStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_status_cd _
                )
    
    Call updateSOIStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.p_sales_ord_id _
                , Me.p_status_cd _
                )
                
    If Me.p_status_cd = "OP" Then
        Call setStatus( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_sales_ord_id _
                    )
    End If
    
    If Me.o_old_sales_order.p_drop_ship_fl = "Y" And Me.p_status_cd = "OP" Then
        If Me.o_old_sales_order.p_type_cd <> "BL" Then
            Call AddDropShipPo
        End If
    End If
    
    If Me.o_old_sales_order.p_drop_ship_fl = "Y" And Me.p_status_cd = "CA" Then
        Set objPurchaseOrder = CreateObject("converge_purchase.purchase_order")
        Set objPurchaseOrder.o_error = Me.o_error
        Set objPurchaseOrder.o_trans_data = Me.o_trans_data
        objPurchaseOrder.p_po_id = Me.o_old_sales_order.p_drop_ship_po_id
        objPurchaseOrder.p_drop_ship_sales_ord_id = Me.p_sales_ord_id
        objPurchaseOrder.updateExistingPOs
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

Public Sub setStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId As Variant _
            )

    On Error GoTo ErrHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.SetStatus"
    
    Dim objRecordset As adodb.Recordset, _
        strSql As String, _
        strStatusCd As String, _
        strSoiStatusCd As String, _
        strChange_cd As String, _
        objSalesOrderItem As sales_order_item
        
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_sales_ord_id = v_dblSalesOrdId
    Me.p_bypass_nm_addr_fl = "Y"
    Me.p_bypass_desc_fl = "Y"
    Call Me.getSalesOrder
    
    If Me.p_status_cd = "CA" Or _
       Me.p_status_cd = "PL" Or _
       Me.p_status_cd = "CL" Or _
       Me.p_status_cd = "OH" Then
        Exit Sub
    End If
    
    strSql = "select * from sales_order_item " & _
             " where " & fWhereVariable("sales_ord_id", Me.p_sales_ord_id, "N", "=") & _
             " and   " & fWhereVariable("status_cd", "CP", "S", "<>") & _
             " and   " & fWhereVariable("status_cd", "CL", "S", "<>") & _
             " and   " & fWhereVariable("status_cd", "CA", "S", "<>")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
                'if records are returned in the recordset, then, there
                'are open items on the order.
    If objRecordset.EOF = False Then
        strStatusCd = "OP"
    Else
        strStatusCd = "CP"
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    If Me.p_status_cd <> strStatusCd Then
        Call IssueUpdateStatus( _
                      r_objError _
                    , r_objTransData _
                    , v_dblSalesOrdId _
                    , strStatusCd _
                    )
                    
    End If
    
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "3200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub generatePlan()

    On Error GoTo ErrHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "sales_order.cls.generatePlan"

    Dim objMatReq As Object
    Dim objRecordset As adodb.Recordset
    Dim dblPlannedQty As Double
    Dim dblUnplannedQty As Double
    Set objMatReq = CreateObject("converge_prod.mat_req")
    Set objMatReq.o_error = Me.o_error
    Set objMatReq.o_trans_data = Me.o_trans_data
    objMatReq.sales_ord_id = Me.p_sales_ord_id
    Call objMatReq.GetMatReqs
    Set objRecordset = objMatReq.Recordset
    Set objMatReq = Nothing
    
    If objRecordset.RecordCount > 0 Then
        Call objRecordset.MoveFirst
    End If
    
    Do Until objRecordset.EOF = True Or _
        Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F"
        dblPlannedQty = objRecordset("back_ordered_qty") + objRecordset("ready_qty") + objRecordset("consumed_qty")
        dblUnplannedQty = objRecordset("required_qty") - dblPlannedQty
        If dblUnplannedQty > 0 And _
           objRecordset("status_cd") <> "CA" _
           And objRecordset("status_cd") <> "CL" _
           And objRecordset("status_cd") <> "OH" _
           And objRecordset("status_cd") <> "CP" Then
            Set objMatReq = CreateObject("converge_prod.mat_req")
            Set objMatReq.o_error = Me.o_error
            Set objMatReq.o_trans_data = Me.o_trans_data
            objMatReq.mat_req_id = objRecordset("mat_req_id")
            Call objMatReq.generatePlan
        End If
        Call objRecordset.MoveNext
    Loop

    Exit Sub
ErrHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub





Private Sub AddName(objName)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.AddName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "sales_order"
    objName.p_obj_id = p_sales_ord_id
    Call objName.AddName
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "3400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ChangeName(objName)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.ChangeName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "sales_order"
    Call objName.ChangeName
    
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

Private Sub DeleteName(objName)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.DeleteName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "sales_order"
    Call objName.DeleteName
    
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

Private Sub AddAddr(objAddress)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.AddAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_id = p_sales_ord_id
    objAddress.p_obj_nm = "sales_order"
    Call objAddress.AddAddress
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "4000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ChangeAddr(objAddress)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.ChangeAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_nm = "sales_order"
    Call objAddress.ChangeAddress
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "4200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub DeleteAddr(objAddress)

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.DeleteSAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_nm = "sales_order"
    Call objAddress.DeleteAddress
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "4400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteData()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order.cls.IssueDeleteData"

    strSql = "delete from sales_order "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "4600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteLineItem()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order.cls.IssueDeleteLineItem"

    strSql = "delete from sales_order_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "4800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub IssueDeleteName()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order.cls.IssueDeleteName"
    
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_sales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "sales_order", "S", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteAddr()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_Order.cls.IssueDeleteAddr"
    
    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub validate(sFunction_cd)
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim varAssoc_id As Variant
    strRoutine_nm = "sales_order.cls.Validate"
    
    If sFunction_cd = "delete" Then
        If IsNull(o_old_sales_order.p_sales_ord_id) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "5400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1206
            Exit Sub
        End If
        If o_old_sales_order.p_status_cd <> "PL" Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "5600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1207
            Exit Sub
        End If
                                        ' 9/2/2001 Blanket orders can not
                                        ' longer be deleted from the system.
                                        ' this is because we are now generating
                                        ' sub sales orders from blanket orders
                                        ' that are in a planned status.
        If o_old_sales_order.p_type_cd = "BL" Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "5600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 4002
            Exit Sub
        End If
    End If
    
    If fValuePresent(p_currency_cd) = True And _
       fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "5800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1025
        Exit Sub
    End If
    
    If fValuePresent(p_sales_tax_pct) = True And _
       (IsNumeric(p_sales_tax_pct) = False Or _
        p_sales_tax_pct < 0 Or _
        p_sales_tax_pct > 1) Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "6000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1030
        Exit Sub
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If fValuePresent(Me.p_sales_rep_assoc_nbr) = True Then
            Call ValidateAssoc_nbr(p_sales_rep_assoc_nbr, varAssoc_id, o_error, o_trans_data)
        Else
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1150"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2165
            Exit Sub
        End If
        
        p_sales_rep_assoc_id = varAssoc_id
        If fValuePresent(p_drop_ship_supplier_nbr) Then
            p_drop_ship_supplier_id = fGetSupplierId(p_drop_ship_supplier_nbr, o_error, o_trans_data)
            If IsNull(p_drop_ship_supplier_id) = True Then
                o_error.p_type_cd = "E"
                o_error.p_err_cd = "6200"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1007
                Exit Sub
            End If
        End If
    End If
    
    If sFunction_cd = "add" Then
        If IsNull(fGetInvLocNm(p_inv_loc_id, o_error, o_trans_data)) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "6400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1175
            Exit Sub
        End If
    End If
        
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

Private Sub IssueInsertData()

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order.cls.IssueInsertData"
    
    p_sales_ord_id = fGetId("sales_ord_id", o_error, o_trans_data)
    
    strSql = "insert into sales_order ( "
    strSql = strSql & "  sales_ord_id"
    strSql = strSql & ", blanket_sales_ord_id"
    strSql = strSql & ", carrier_nm"
    strSql = strSql & ", contract_nbr"
    strSql = strSql & ", currency_cd"
    strSql = strSql & ", cust_id"
    strSql = strSql & ", cust_po_nbr"
    strSql = strSql & ", consignment_fl"
    strSql = strSql & ", discount_pct"
    strSql = strSql & ", drop_ship_fl"
    strSql = strSql & ", drop_ship_supplier_id"
    strSql = strSql & ", drop_ship_po_id"
    strSql = strSql & ", fob_terms_cd"
    strSql = strSql & ", freight_cd"
    strSql = strSql & ", inv_loc_id"
    strSql = strSql & ", ordered_dt"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", price_calc_meth_cd"
    strSql = strSql & ", revision_nbr"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", sales_quote_id"
    strSql = strSql & ", sales_rep_assoc_id"
    strSql = strSql & ", sales_tax_pct"
    strSql = strSql & ", transportation_cd"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", conv_from_ord_id "
    strSql = strSql & ", order_tx "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_blanket_sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_carrier_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_contract_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_currency_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_cust_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_po_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_consignment_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_discount_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_drop_ship_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_drop_ship_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(Null, "N")
    strSql = strSql & "," & fInsertVariable(p_fob_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_freight_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_id, "S")
    strSql = strSql & "," & fInsertVariable(p_ordered_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_price_calc_meth_cd, "S")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable("PL", "S")
    strSql = strSql & "," & fInsertVariable(p_sales_quote_id, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_rep_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_tax_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_transportation_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_conv_from_ord_id, "S")
    strSql = strSql & "," & fInsertVariable(p_order_text, "S")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
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

Private Sub IssueUpdateData()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "sales_order.cls.IssueUpdateData"
    
    strSql = "update sales_order set "
    strSql = strSql & fUpdateVariable("blanket_sales_ord_id", p_blanket_sales_ord_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("carrier_nm", p_carrier_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("contract_nbr", p_contract_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("currency_cd", p_currency_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_po_nbr", p_cust_po_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("discount_pct", p_discount_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("fob_terms_cd", p_fob_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("freight_cd", p_freight_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ordered_dt", p_ordered_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("price_calc_meth_cd", p_price_calc_meth_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("drop_ship_supplier_id", p_drop_ship_supplier_id, "N", "=")
    strSql = strSql & ", revision_nbr = revision_nbr + 1 "
    strSql = strSql & "," & fUpdateVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_rep_assoc_id", p_sales_rep_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_tax_pct", p_sales_tax_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("order_tx", p_order_text, "S", "=")
    strSql = strSql & "," & fUpdateVariable("transportation_cd", p_transportation_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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

Private Sub SelectSalesOrder()

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "sales_order.cls.SelectsalesOrder"

    If IsNull(p_sales_ord_id) = True Then
       p_sales_ord_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM sales_order "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
       p_sales_ord_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_sales_ord_id = objRecordset("sales_ord_id")
        p_inv_loc_id = objRecordset("inv_loc_id")
        p_blanket_sales_ord_id = objRecordset("blanket_sales_ord_id")
        p_carrier_nm = objRecordset("carrier_nm")
        p_contract_nbr = objRecordset("contract_nbr")
        p_currency_cd = objRecordset("currency_cd")
        p_cust_id = objRecordset("cust_id")
        p_cust_po_nbr = objRecordset("cust_po_nbr")
        p_consignment_fl = objRecordset("consignment_fl")
        p_discount_pct = objRecordset("discount_pct")
        p_drop_ship_fl = objRecordset("drop_ship_fl")
        p_drop_ship_supplier_id = objRecordset("drop_ship_supplier_id")
        p_drop_ship_po_id = objRecordset("drop_ship_po_id")
        p_fob_terms_cd = objRecordset("fob_terms_cd")
        p_freight_cd = objRecordset("freight_cd")
        p_ordered_dt = objRecordset("ordered_dt")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_price_calc_meth_cd = objRecordset("price_calc_meth_cd")
        p_revision_nbr = objRecordset("revision_nbr")
        p_status_cd = objRecordset("status_cd")
        p_sales_quote_id = objRecordset("sales_quote_id")
        p_sales_rep_assoc_id = objRecordset("sales_rep_assoc_id")
        p_sales_tax_pct = objRecordset("sales_tax_pct")
        p_transportation_cd = objRecordset("transportation_cd")
        p_conv_from_ord_id = objRecordset("conv_from_ord_id")
        p_type_cd = objRecordset("type_cd")
        p_order_text = objRecordset("order_tx")
        
        objRecordset.Close
        Set objRecordset = Nothing
        
        'p_gross_am = 0
        'p_discount_am = p_gross_am * p_discount_pct
        'p_sales_tax_am = (p_gross_am - p_discount_am) * p_sales_tax_pct
        'p_net_am = p_gross_am - p_discount_am + p_sales_tax_am
    
        If p_bypass_desc_fl = "Y" Then
        Else
            p_currency_desc = fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data)
            p_cust_nbr = fGetCustNbr(p_cust_id, o_error, o_trans_data)
            p_drop_ship_supplier_nbr = fGetSupplierNbr(p_drop_ship_supplier_id, o_error, o_trans_data)
            p_fob_terms_desc = fGetDesc("fob_terms_cd", p_fob_terms_cd, o_error, o_trans_data)
            p_freight_desc = fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data)
            p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
            p_price_calc_meth_desc = fGetDesc("price_calc_cd", p_price_calc_meth_cd, o_error, o_trans_data)
            p_status_desc = fGetDesc("sales_order_status_cd", p_status_cd, o_error, o_trans_data)
            p_sales_rep_assoc_nbr = fGetAssocNbr(p_sales_rep_assoc_id, o_error, o_trans_data)
            p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data)
            p_type_desc = fGetDesc("sales_order_type_cd", p_type_cd, o_error, o_trans_data)
            p_inv_loc_nm = fGetInvLocNm(p_inv_loc_id, o_error, o_trans_data)
        End If
    End If
     
    Exit Sub
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
End Sub

Public Sub getCustSalesOrders()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.GetCustSalesOrders"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = "SELECT  " & _
                "  a.sales_ord_id, " & _
                "  a.cust_id " & _
                ", b.cust_nbr " & _
                ", a.type_cd " & _
                ", a.status_cd " & _
                ", a.drop_ship_fl " & _
                ", a.conv_from_ord_id " & _
                ", a.drop_ship_supplier_id " & _
                ", a.drop_ship_po_id, a.conv_from_ord_id " & _
                ", c.code_desc as status_desc" & _
                ", d.code_desc as type_desc " & _
                " FROM sales_order a, customer b, code_desc c, code_desc d" & _
               " WHERE (" & fWhereVariable("c.field_nm", "sales_order_status_cd", "S", "=") & _
               " AND c.code = a.status_cd)" & _
                " AND (" & fWhereVariable("d.field_nm", "sales_order_type_cd", "S", "=") & _
               " AND d.code = a.type_cd)" & _
                " AND a.cust_id = b.cust_id "
                
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
               
    p_cust_id = fGetCustId(p_cust_nbr, o_error, o_trans_data)
    If fValuePresent(p_cust_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.cust_id", p_cust_id, "S", "=")
    End If
               
    p_drop_ship_supplier_id = fGetSupplierId(p_drop_ship_supplier_nbr, o_error, o_trans_data)
    If fValuePresent(p_drop_ship_supplier_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("drop_ship_supplier_id", p_drop_ship_supplier_id, "N", "=")
    End If
               
    If fValuePresent(p_drop_ship_po_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("drop_ship_po_id", p_drop_ship_po_id, "N", "=")
    End If
               
    If fValuePresent(p_conv_from_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("conv_from_ord_id", p_conv_from_ord_id, "N", "=")
    End If
               
               
    If fValuePresent(p_blanket_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("blanket_sales_ord_id", p_blanket_sales_ord_id, "N", "=")
    End If
               
    If fValuePresent(p_cust_po_nbr) = True Then
       strSql = strSql & "  AND " & fWhereVariable("cust_po_nbr", p_cust_po_nbr, "S", "=")
    End If
               
    If fValuePresent(p_type_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("type_cd", p_type_cd, "S", "=")
    End If
               
    If fValuePresent(p_sales_quote_id) = True Then
        If IsNumeric(p_sales_quote_id) = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "7400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1008
            Exit Sub
        End If
        strSql = strSql & "  AND " & fWhereVariable("sales_quote_id", p_sales_quote_id, "S", "=")
    End If
               
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
               
    strSql = strSql & " order by sales_ord_id "
               
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
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


Private Sub IssueUpdateStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_varSalesOrdId As Variant _
            , ByVal v_varStatusCd As Variant _
            )

    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.IssueUpdateStatus"
    
    Dim strSql As String, _
        dblUpdateQty As Double
    
    strSql = "update sales_order set "
    strSql = strSql & fUpdateVariable("status_cd", v_varStatusCd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_ord_id", v_varSalesOrdId, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdateQty, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    If dblUpdateQty = 0 Then
        With r_objError
            .p_type_cd = "F"
            .p_err_cd = "8000"
            .p_routine_nm = strRoutine_nm
            .p_message_id = 1010
        End With
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

Private Sub CalcValues()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
     
    On Error GoTo ErrHandler
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "sales_order.cls.CalcValues"
    
    strSql = "SELECT item_price, ordered_qty " & _
             " FROM sales_order_item " & _
             " WHERE " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If objRecordset Is Nothing Or objRecordset.EOF = True Then
        p_gross_am = 0
        p_discount_am = 0
        p_sales_tax_am = 0
        p_net_am = 0
        Exit Sub
    End If
    
    objRecordset.MoveFirst
    p_gross_am = 0
    p_discount_am = 0
    p_sales_tax_am = 0
    p_net_am = 0
    Do Until objRecordset.EOF
        If IsNull(objRecordset("item_price")) Or IsNull(objRecordset("ordered_qty")) Or objRecordset("ordered_qty") = 0 Or objRecordset("item_price") = 0 Then
            p_gross_am = 0
        Else
            p_gross_am = p_gross_am + (objRecordset("item_price") * objRecordset("ordered_qty"))
        End If
        objRecordset.MoveNext
    Loop
    
    p_discount_am = CDbl(p_discount_pct * p_gross_am)
    p_sales_tax_am = CDbl(p_sales_tax_pct * (p_gross_am - p_discount_am))
    p_net_am = CDbl(p_gross_am - p_discount_am + p_sales_tax_am)
    
    Exit Sub
ErrHandler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8300"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getDropShipSalesOrdersReadytoPrint()
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.GetCustSalesOrders"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = "SELECT  " & _
                "  a.sales_ord_id, " & _
                "  a.cust_id " & _
                ", b.cust_nbr " & _
                ", a.type_cd " & _
                ", a.status_cd " & _
                ", a.drop_ship_fl " & _
                ", a.drop_ship_supplier_id " & _
                ", a.drop_ship_po_id " & _
                ", c.code_desc as status_desc" & _
                ", d.code_desc as type_desc " & _
                ", e.code_desc as purchase_order_status_desc " & _
              " FROM sales_order a, customer b, code_desc c, code_desc d, code_desc e, purchase_order f" & _
             " WHERE (" & fWhereVariable("c.field_nm", "sales_order_status_cd", "S", "=") & _
               " AND c.code = a.status_cd)" & _
               " AND (" & fWhereVariable("d.field_nm", "sales_order_type_cd", "S", "=") & _
               " AND d.code = a.type_cd)" & _
               " AND (" & fWhereVariable("e.field_nm", "purchase_order_status_cd", "S", "=") & _
               " AND e.code = f.status_cd)" & _
               " AND a.drop_ship_po_id = f.po_id " & _
               " AND a.cust_id = b.cust_id " & _
               " AND (f.status_cd <> 'ST' " & _
               " AND f.status_cd <> 'CL'" & _
               " AND f.status_cd <> 'OH' " & _
               " AND f.status_cd <> 'CA') "
                
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
               
               
               
    p_cust_id = fGetCustId(p_cust_nbr, o_error, o_trans_data)
    If fValuePresent(p_cust_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.cust_id", p_cust_id, "S", "=")
    End If
               
    p_drop_ship_supplier_id = fGetSupplierId(p_drop_ship_supplier_nbr, o_error, o_trans_data)
    If fValuePresent(p_drop_ship_supplier_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("drop_ship_supplier_id", p_drop_ship_supplier_id, "N", "=")
    End If
               
    If fValuePresent(p_drop_ship_po_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("drop_ship_po_id", p_drop_ship_po_id, "N", "=")
    End If
               
    If fValuePresent(p_blanket_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("blanket_sales_ord_id", p_blanket_sales_ord_id, "N", "=")
    End If
               
    If fValuePresent(p_cust_po_nbr) = True Then
       strSql = strSql & "  AND " & fWhereVariable("cust_po_nbr", p_cust_po_nbr, "S", "=")
    End If
               
               
    If fValuePresent(p_type_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("type_cd_cd", p_type_cd, "S", "=")
    End If
               
               
    If fValuePresent(p_sales_quote_id) = True Then
        If IsNumeric(p_sales_quote_id) = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "7400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1008
            Exit Sub
        End If
        strSql = strSql & "  AND " & fWhereVariable("sales_quote_id", p_sales_quote_id, "S", "=")
    End If
               
        
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
               
    strSql = strSql & " order by sales_ord_id "
               
               
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
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

Private Sub updateSOIStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId As Double _
            , ByVal v_strStatusCd As String _
            )
    
    On Error GoTo ErrHandler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.upateSOIStatus"
    
    Dim objSalesOrderItem As sales_order_item, _
        objRecordset As adodb.Recordset

    Call getSalesOrderItems( _
                  r_objError _
                , r_objTransData _
                , v_dblSalesOrdId _
                , objRecordset _
                )
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    If objRecordset.EOF = True Then
        With r_objError
            .p_type_cd = "F"
            .p_err_cd = "3005"
            .p_nbr = Err.Number
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 2185
        End With
        Exit Sub
    End If

    Set objSalesOrderItem = New sales_order_item
    Set objSalesOrderItem.o_error = r_objError
    Set objSalesOrderItem.o_trans_data = r_objTransData
    
    Do While objRecordset.EOF = False And r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F"
        objSalesOrderItem.p_sales_ord_id = v_dblSalesOrdId
        objSalesOrderItem.p_line_nbr = objRecordset("line_nbr")
        objSalesOrderItem.p_status_cd = v_strStatusCd
        Call objSalesOrderItem.changeStatus
        objRecordset.MoveNext
    Loop
    
    Set objSalesOrderItem = Nothing
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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


Private Sub getSalesOrderItems( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId _
            , ByRef r_objRecordset As adodb.Recordset _
            )

    On Error GoTo ErrHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.getSalesOrderItems"
    
    Dim strSql As String
    
    strSql = "select line_nbr, type_cd from sales_order_item " & _
                " where" & _
                "   " & fWhereVariable("sales_ord_id", v_dblSalesOrdId, "N", "=")

    Call r_objTransData.OpenRecordset(r_objRecordset, strSql, r_objError)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set r_objRecordset = Nothing
        Exit Sub
    End If

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



Public Sub incrementRevision( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblSalesOrdId As Double _
            , ByRef r_objSalesOrder As sales_order _
            , ByRef v_dblAssocId As Double _
            , ByRef v_dteEnteredTs As Date _
            , ByRef v_varCommentTx As Variant _
            )

    On Error GoTo ErrHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_order.cls.IncrementRevision"
    
    Dim strSqlTx As String _
       , lngSoRevHistId _
       , dblUpdateQty As Double
    
    If r_objSalesOrder Is Nothing Then
        With r_objSalesOrder
            .p_sales_ord_id = v_dblSalesOrdId
            .p_bypass_desc_fl = True
            .p_bypass_nm_addr_fl = True
            Call .getSalesOrder
        End With
    End If
    
    strSqlTx = _
               "update sales_order set " _
             & " revision_nbr = " & r_objSalesOrder.p_revision_nbr + 1 & " " _
             & ",updated_assoc_id = " & v_dblAssocId & " " _
             & "where sales_ord_id = " & v_dblSalesOrdId
    
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    
    lngSoRevHistId = fGetId("SoRevHistId", r_objError, r_objTransData)
    strSqlTx = _
              "insert into so_rev_hist ( " _
            & "  so_rev_hist_id " _
            & ", sales_ord_id " _
            & ", revision_nbr " _
            & ", assoc_id " _
            & ", entered_ts " _
            & ", comment_tx " _
            & " ) " _
            & "values (" _
            & " " & fInsertVariable(lngSoRevHistId, "N") _
            & "," & fInsertVariable(v_dblSalesOrdId, "N") _
            & "," & fInsertVariable(r_objSalesOrder.p_revision_nbr + 1, "N") _
            & "," & fInsertVariable(v_dblAssocId, "N") _
            & "," & fInsertVariable("current_ts", "D") _
            & "," & fInsertVariable(v_varCommentTx, "S") _
            & ")"
  
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    
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
