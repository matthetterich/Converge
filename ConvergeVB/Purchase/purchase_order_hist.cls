VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "purchase_order_hist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarp_po_id As Variant
Private mvarp_revision_nbr As Variant
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

Private mvarp_send_to_nm As Variant
Private mvarp_send_to_street_addr As Variant
Private mvarp_send_to_city_nm As Variant
Private mvarp_send_to_state_cd As Variant
Private mvarp_send_to_zip_cd As Variant
Private mvarp_send_to_country_cd As Variant

Private mvarp_bill_to_nm As Variant
Private mvarp_bill_to_street_addr As Variant
Private mvarp_bill_to_city_nm As Variant
Private mvarp_bill_to_state_cd As Variant
Private mvarp_bill_to_zip_cd As Variant
Private mvarp_bill_to_country_cd As Variant

Private mvarp_ship_to_nm As Variant
Private mvarp_ship_to_street_addr As Variant
Private mvarp_ship_to_city_nm As Variant
Private mvarp_ship_to_state_cd As Variant
Private mvarp_ship_to_zip_cd As Variant
Private mvarp_ship_to_country_cd As Variant

Private mvaro_error As Object
Private mvaro_record_set As ADODB.Recordset
Private mvaro_trans_data As Object
Private mvarp_comm_meth_cd As Variant
Private mvarp_comm_meth_desc As Variant
Private mvarp_freight_desc As Variant
Private mvarp_sales_tax_pct As Variant
Private mvarp_sales_tax_amnt As Variant
Private mvarp_drop_ship_sales_ord_id As Variant
Private mvarp_bypass_nm_addr_fl As Variant
Private mvarp_bypass_desc_fl As Variant
Private mvarp_entered_dt As Variant
Private mvarp_inv_loc_id As Variant
'

Public Property Let p_send_to_nm(ByVal vData As Variant)
    mvarp_send_to_nm = vData
End Property
Public Property Get p_send_to_nm() As Variant
    p_send_to_nm = mvarp_send_to_nm
End Property

Public Property Let p_send_to_street_addr(ByVal vData As Variant)
    mvarp_send_to_street_addr = vData
End Property
Public Property Get p_send_to_street_addr() As Variant
    p_send_to_street_addr = mvarp_send_to_street_addr
End Property

Public Property Let p_send_to_city_nm(ByVal vData As Variant)
    mvarp_send_to_city_nm = vData
End Property
Public Property Get p_send_to_city_nm() As Variant
    p_send_to_city_nm = mvarp_send_to_city_nm
End Property

Public Property Let p_send_to_state_cd(ByVal vData As Variant)
    mvarp_send_to_state_cd = vData
End Property
Public Property Get p_send_to_state_cd() As Variant
    p_send_to_state_cd = mvarp_send_to_state_cd
End Property

Public Property Let p_send_to_zip_cd(ByVal vData As Variant)
    mvarp_send_to_zip_cd = vData
End Property
Public Property Get p_send_to_zip_cd() As Variant
    p_send_to_zip_cd = mvarp_send_to_zip_cd
End Property

Public Property Let p_send_to_country_cd(ByVal vData As Variant)
    mvarp_send_to_country_cd = vData
End Property
Public Property Get p_send_to_country_cd() As Variant
    p_send_to_country_cd = mvarp_send_to_country_cd
End Property

Public Property Let p_bill_to_nm(ByVal vData As Variant)
    mvarp_bill_to_nm = vData
End Property
Public Property Get p_bill_to_nm() As Variant
    p_bill_to_nm = mvarp_bill_to_nm
End Property

Public Property Let p_bill_to_street_addr(ByVal vData As Variant)
    mvarp_bill_to_street_addr = vData
End Property
Public Property Get p_bill_to_street_addr() As Variant
    p_bill_to_street_addr = mvarp_bill_to_street_addr
End Property

Public Property Let p_bill_to_city_nm(ByVal vData As Variant)
    mvarp_bill_to_city_nm = vData
End Property
Public Property Get p_bill_to_city_nm() As Variant
    p_bill_to_city_nm = mvarp_bill_to_city_nm
End Property

Public Property Let p_bill_to_state_cd(ByVal vData As Variant)
    mvarp_bill_to_state_cd = vData
End Property
Public Property Get p_bill_to_state_cd() As Variant
    p_bill_to_state_cd = mvarp_bill_to_state_cd
End Property

Public Property Let p_bill_to_zip_cd(ByVal vData As Variant)
    mvarp_bill_to_zip_cd = vData
End Property
Public Property Get p_bill_to_zip_cd() As Variant
    p_bill_to_zip_cd = mvarp_bill_to_zip_cd
End Property

Public Property Let p_bill_to_country_cd(ByVal vData As Variant)
    mvarp_bill_to_country_cd = vData
End Property
Public Property Get p_bill_to_country_cd() As Variant
    p_bill_to_country_cd = mvarp_bill_to_country_cd
End Property

Public Property Let p_ship_to_nm(ByVal vData As Variant)
    mvarp_ship_to_nm = vData
End Property
Public Property Get p_ship_to_nm() As Variant
    p_ship_to_nm = mvarp_ship_to_nm
End Property

Public Property Let p_ship_to_street_addr(ByVal vData As Variant)
    mvarp_ship_to_street_addr = vData
End Property
Public Property Get p_ship_to_street_addr() As Variant
    p_ship_to_street_addr = mvarp_ship_to_street_addr
End Property

Public Property Let p_ship_to_city_nm(ByVal vData As Variant)
    mvarp_ship_to_city_nm = vData
End Property
Public Property Get p_ship_to_city_nm() As Variant
    p_ship_to_city_nm = mvarp_ship_to_city_nm
End Property

Public Property Let p_ship_to_state_cd(ByVal vData As Variant)
    mvarp_ship_to_state_cd = vData
End Property
Public Property Get p_ship_to_state_cd() As Variant
    p_ship_to_state_cd = mvarp_ship_to_state_cd
End Property

Public Property Let p_ship_to_zip_cd(ByVal vData As Variant)
    mvarp_ship_to_zip_cd = vData
End Property
Public Property Get p_ship_to_zip_cd() As Variant
    p_ship_to_zip_cd = mvarp_ship_to_zip_cd
End Property

Public Property Let p_ship_to_country_cd(ByVal vData As Variant)
    mvarp_ship_to_country_cd = vData
End Property
Public Property Get p_ship_to_country_cd() As Variant
    p_ship_to_country_cd = mvarp_ship_to_country_cd
End Property


Public Property Let p_move_to_hist_fl(ByVal vData As Variant)
    mvarp_move_to_hist_fl = vData
End Property
Public Property Get p_move_to_hist_fl() As Variant
        p_move_to_hist_fl = mvarp_move_to_hist_fl
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
'


Public Sub getPurchaseOrderHistory()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
     
    Dim strRoutine_nm As String
    strRoutine_nm = "purchsase_order_hist.cls.GetPurchaseOrderHistory"
    
    If p_po_id > 0 Then
    Else
        p_po_id = Null
        Exit Sub
    End If
    
    Call SelectPurchaseOrderHistory
    
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

Public Sub getPurchaseOrderHistories()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_hist.cls.getPurchaseOrderHistories"
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    
    strSql = " SELECT purchase_order_hist.po_id, purchase_order_hist.revision_nbr, supplier.supplier_nbr," & _
    " code_desc.code_desc, purchase_order_hist.entered_ts" & _
    " FROM purchase_order_hist LEFT OUTER JOIN " & _
    " code_desc ON " & _
    " purchase_order_hist.drop_ship_fl = code_desc.code LEFT OUTER Join " & _
    " supplier ON " & _
    " purchase_order_hist.supplier_id = supplier.supplier_id " & _
    " GROUP BY purchase_order_hist.revision_nbr, " & _
    " supplier.supplier_nbr, code_desc.code_desc, " & _
    " purchase_order_hist.po_id, purchase_order_hist.entered_ts "
    
    If fValuePresent(p_po_id) Then
        strSql = strSql & " HAVING " & fWhereVariable("purchase_order_hist.po_id", p_po_id, "S", "=")
    End If
       
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


Private Sub SelectPurchaseOrderHistory()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
   
    strRoutine_nm = "Purchase_order_hist.cls.SelectPurchaseOrder"
    
    If IsNull(p_po_id) = True Then
     p_po_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM Purchase_order_hist "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("po_id", p_po_id, "N", "=")
        strSql = strSql & " AND " & fWhereVariable("revision_nbr", p_revision_nbr, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If objRecordset.EOF = True Then
        ' Close the Record Set
        objRecordset.Close
        Set objRecordset = Nothing
        p_po_id = Null
    Else
        ' set properties from recordset
        p_revision_nbr = objRecordset("revision_nbr")
        p_inv_loc_id = objRecordset("revision_nbr")
        p_backorder_cd = objRecordset("backorder_cd")
        p_buyer_assoc_id = objRecordset("buyer_assoc_id")
        p_carrier_nm = objRecordset("carrier_nm")
        p_confirmation_nbr = objRecordset("confirmation_nbr")
        p_currency_cd = objRecordset("currency_cd")
        p_drop_ship_fl = objRecordset("drop_ship_fl")
        p_drop_ship_sales_ord_id = objRecordset("drop_ship_sales_ord_id")
        p_fob_terms_cd = objRecordset("fob_terms_cd")
        p_freight_cd = objRecordset("freight_cd")
        p_comm_meth_cd = objRecordset("comm_meth_cd")
        p_entered_dt = objRecordset("entered_dt")
        p_opened_dt = objRecordset("opened_dt")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_status_cd = objRecordset("status_cd")
        p_sales_tax_pct = objRecordset("sales_tax_pct")
        p_supplier_id = objRecordset("supplier_id")
        p_transportation_cd = objRecordset("transportation_cd")
        p_type_cd = objRecordset("type_cd")
        
        p_send_to_nm = objRecordset("send_to_nm")
        p_send_to_street_addr = objRecordset("send_to_street_addr")
        p_send_to_city_nm = objRecordset("send_to_city_nm")
        p_send_to_state_cd = objRecordset("send_to_state_cd")
        p_send_to_zip_cd = objRecordset("send_to_zip_cd")
        p_send_to_country_cd = objRecordset("send_to_country_cd")
        
        p_bill_to_nm = objRecordset("bill_to_nm")
        p_bill_to_street_addr = objRecordset("bill_to_street_addr")
        p_bill_to_city_nm = objRecordset("bill_to_city_nm")
        p_bill_to_state_cd = objRecordset("bill_to_state_cd")
        p_bill_to_zip_cd = objRecordset("bill_to_zip_cd")
        p_bill_to_country_cd = objRecordset("bill_to_country_cd")
                
        p_ship_to_nm = objRecordset("ship_to_nm")
        p_ship_to_street_addr = objRecordset("ship_to_street_addr")
        p_ship_to_city_nm = objRecordset("ship_to_city_nm")
        p_ship_to_state_cd = objRecordset("ship_to_state_cd")
        p_ship_to_zip_cd = objRecordset("ship_to_zip_cd")
        p_ship_to_country_cd = objRecordset("ship_to_country_cd")
        
        
        'Set Descriptions
        p_supplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
        p_currency_desc = fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data)
        p_type_desc = fGetDesc("Purchase_order_type_cd", p_type_cd, o_error, o_trans_data)
        p_comm_meth_desc = fGetDesc("comm_meth_cd", p_comm_meth_cd, o_error, o_trans_data)
        p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("purchase_order_status_cd", p_status_cd, o_error, o_trans_data)
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_fob_terms_desc = fGetDesc("fob_terms_cd", p_fob_terms_cd, o_error, o_trans_data)
        p_freight_desc = fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data)
        p_backorder_desc = fGetDesc("backorder_cd", p_backorder_cd, o_error, o_trans_data)
        p_drop_ship_desc = fGetDesc("flag", p_drop_ship_fl, o_error, o_trans_data)
        p_buyer_assoc_nbr = fGetAssocNbr(p_buyer_assoc_id, o_error, o_trans_data)
        
        
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
