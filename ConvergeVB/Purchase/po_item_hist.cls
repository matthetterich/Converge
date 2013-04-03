VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "po_item_hist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarp_po_id As Variant
Private mvarp_line_nbr As Variant
Private mvarp_po_revision_nbr As Variant
Private mvarp_item_desc As Variant
Private mvarp_inspection_req_fl As Variant
Private mvarp_item_id As Variant
Private mvarp_item_nbr As Variant
Private mvarp_item_cost As Variant
Private mvarp_drop_ship_so_line_nbr As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_requested_del_dt As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_uom_cd As Variant
Private mvarp_uom_desc As Variant
Private mvarp_supplier_item_nbr As Variant
Private mvarp_ordered_qty As Variant
Private mvarp_received_qty As Variant
Private mvarp_rejected_qty As Variant
Private mvarp_invoiced_qty As Variant
Private mvaro_record_set As ADODB.Recordset
Private mvaro_error As Object
Private mvaro_trans_data As Object
'

Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
    p_po_id = mvarp_po_id
End Property

Public Property Let p_line_nbr(ByVal vData As Variant)
    mvarp_line_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
    p_line_nbr = mvarp_line_nbr
End Property

Public Property Let p_po_revision_nbr(ByVal vData As Variant)
    mvarp_po_revision_nbr = vData
End Property
Public Property Get p_po_revision_nbr() As Variant
    p_po_revision_nbr = mvarp_po_revision_nbr
End Property

Public Property Let p_item_desc(ByVal vData As Variant)
    mvarp_item_desc = vData
End Property
Public Property Get p_item_desc() As Variant
    p_item_desc = mvarp_item_desc
End Property

Public Property Let p_inspection_req_fl(ByVal vData As Variant)
    mvarp_inspection_req_fl = vData
End Property
Public Property Get p_inspection_req_fl() As Variant
    p_inspection_req_fl = mvarp_inspection_req_fl
End Property

Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = mvarp_item_id
End Property

Public Property Let p_item_nbr(ByVal vData As Variant)
    mvarp_item_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
    p_item_nbr = mvarp_item_nbr
End Property

Public Property Let p_item_cost(ByVal vData As Variant)
    mvarp_item_cost = vData
End Property
Public Property Get p_item_cost() As Variant
    p_item_cost = mvarp_item_cost
End Property

Public Property Let p_drop_ship_so_line_nbr(ByVal vData As Variant)
    mvarp_drop_ship_so_line_nbr = vData
End Property
Public Property Get p_drop_ship_so_line_nbr() As Variant
    p_drop_ship_so_line_nbr = mvarp_drop_ship_so_line_nbr
End Property

Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    mvarp_payment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
    p_payment_terms_cd = mvarp_payment_terms_cd
End Property

Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    mvarp_payment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
    p_payment_terms_desc = mvarp_payment_terms_desc
End Property

Public Property Let p_requested_del_dt(ByVal vData As Variant)
    mvarp_requested_del_dt = vData
End Property
Public Property Get p_requested_del_dt() As Variant
    p_requested_del_dt = mvarp_requested_del_dt
End Property

Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
    p_status_cd = mvarp_status_cd
End Property

Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
    p_status_desc = mvarp_status_desc
End Property

Public Property Let p_uom_cd(ByVal vData As Variant)
    mvarp_uom_cd = vData
End Property
Public Property Get p_uom_cd() As Variant
    p_uom_cd = mvarp_uom_cd
End Property

Public Property Let p_uom_desc(ByVal vData As Variant)
    mvarp_uom_desc = vData
End Property
Public Property Get p_uom_desc() As Variant
    p_uom_desc = mvarp_uom_desc
End Property

Public Property Let p_supplier_item_nbr(ByVal vData As Variant)
    mvarp_supplier_item_nbr = vData
End Property
Public Property Get p_supplier_item_nbr() As Variant
    p_supplier_item_nbr = mvarp_supplier_item_nbr
End Property

Public Property Let p_ordered_qty(ByVal vData As Variant)
    mvarp_ordered_qty = vData
End Property
Public Property Get p_ordered_qty() As Variant
    p_ordered_qty = mvarp_ordered_qty
End Property

Public Property Let p_received_qty(ByVal vData As Variant)
    mvarp_received_qty = vData
End Property
Public Property Get p_received_qty() As Variant
    p_received_qty = mvarp_received_qty
End Property

Public Property Let p_rejected_qty(ByVal vData As Variant)
    mvarp_rejected_qty = vData
End Property
Public Property Get p_rejected_qty() As Variant
    p_rejected_qty = mvarp_rejected_qty
End Property

Public Property Let p_invoiced_qty(ByVal vData As Variant)
    mvarp_invoiced_qty = vData
End Property
Public Property Get p_invoiced_qty() As Variant
    p_invoiced_qty = mvarp_invoiced_qty
End Property

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property


Public Sub getPurchaseOrderItemHistories()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_item_hist.cls.getPurchaseOrderItemHistories"
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strSql = " SELECT purchase_order_item_hist.po_id, item.item_nbr, "
    strSql = strSql & " purchase_order_item_hist.item_desc, code_desc.code_desc,"
    strSql = strSql & " purchase_order_item_hist.supplier_item_nbr,"
    strSql = strSql & " purchase_order_item_hist.item_cost,"
    strSql = strSql & " purchase_order_item_hist.ordered_qty,"
    strSql = strSql & " purchase_order_item_hist.received_qty,"
    strSql = strSql & " purchase_order_item_hist.requested_del_dt,"
    strSql = strSql & " purchase_order_item_hist.line_nbr,"
    strSql = strSql & " purchase_order_item_hist.item_cost * purchase_order_item_hist.ordered_qty as line_am"
    strSql = strSql & " FROM purchase_order_item_hist LEFT OUTER JOIN"
    strSql = strSql & " code_desc ON"
    strSql = strSql & " purchase_order_item_hist.status_cd = code_desc.code AND"
    strSql = strSql & " code_desc.field_nm = 'purchase_order_item_status_cd' LEFT OUTER Join"
    strSql = strSql & " item ON purchase_order_item_hist.item_id = item.item_id"
    strSql = strSql & " WHERE 1 = 1"
    
    If fValuePresent(p_po_id) = True Then
        strSql = strSql & "  AND " & fWhereVariable("purchase_order_item_hist.po_id", p_po_id, "N", "=")
    End If
    
    If fValuePresent(p_po_revision_nbr) = True Then
        strSql = strSql & "  AND " & fWhereVariable("purchase_order_item_hist.po_revision_nbr", p_po_revision_nbr, "N", "=")
    End If
    
    strSql = strSql & " ORDER BY purchase_order_item_hist.requested_del_dt "
        
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
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
