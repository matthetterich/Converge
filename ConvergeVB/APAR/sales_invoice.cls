VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_currency_cd As Variant
Private mvarp_sales_inv_id As Variant
Private mvarp_standard_tx As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_type_cd As Variant
Private mvarp_type_desc As Variant
Private mvarp_last_print_dt As Variant
Private mvarp_revision_nbr As Variant
Private mvarp_gross_inv_amount As Variant
Private mvarp_discount_pct As Variant
Private mvarp_disc_amount As Variant
Private mvarp_sales_tax_pct As Variant
Private mvarp_sales_tax_amount As Variant
Private mvarp_net_inv_amount As Variant
Private mvarp_sales_ord_id As Variant
Private mvarp_ship_id As Variant
Private mvaro_record_set As adodb.Recordset
'


Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_currency_cd(ByVal vData As Variant)
    mvarp_currency_cd = vData
End Property
Public Property Get p_currency_cd() As Variant
    p_currency_cd = mvarp_currency_cd
End Property
Public Property Let p_ship_id(ByVal vData As Variant)
    mvarp_ship_id = vData
End Property
Public Property Get p_ship_id() As Variant
    p_ship_id = mvarp_ship_id
End Property
Public Property Let p_sales_ord_id(ByVal vData As Variant)
    mvarp_sales_ord_id = vData
End Property
Public Property Set p_sales_ord_id(ByVal vData As Variant)
    Set mvarp_sales_ord_id = vData
End Property
Public Property Get p_sales_ord_id() As Variant
    p_sales_ord_id = mvarp_sales_ord_id
End Property
Public Property Let p_net_inv_amount(ByVal vData As Variant)
    mvarp_net_inv_amount = vData
End Property
Public Property Get p_net_inv_amount() As Variant
    p_net_inv_amount = mvarp_net_inv_amount
End Property
Public Property Let p_sales_tax_amount(ByVal vData As Variant)
    mvarp_sales_tax_amount = vData
End Property
Public Property Set p_sales_tax_amount(ByVal vData As Variant)
    Set mvarp_sales_tax_amount = vData
End Property
Public Property Get p_sales_tax_amount() As Variant
    p_sales_tax_amount = mvarp_sales_tax_amount
End Property
Public Property Let p_sales_tax_pct(ByVal vData As Variant)
    mvarp_sales_tax_pct = vData
End Property
Public Property Get p_sales_tax_pct() As Variant
    p_sales_tax_pct = mvarp_sales_tax_pct
End Property
Public Property Let p_disc_amount(ByVal vData As Variant)
    mvarp_disc_amount = vData
End Property
Public Property Get p_disc_amount() As Variant
    p_disc_amount = mvarp_disc_amount
End Property
Public Property Let p_discount_pct(ByVal vData As Variant)
    mvarp_discount_pct = vData
End Property
Public Property Set p_discount_pct(ByVal vData As Variant)
    Set mvarp_discount_pct = vData
End Property
Public Property Get p_discount_pct() As Variant
    p_discount_pct = mvarp_discount_pct
End Property
Public Property Let p_gross_inv_amount(ByVal vData As Variant)
    mvarp_gross_inv_amount = vData
End Property
Public Property Get p_gross_inv_amount() As Variant
    p_gross_inv_amount = mvarp_gross_inv_amount
End Property
Public Property Let p_revision_nbr(ByVal vData As Variant)
    mvarp_revision_nbr = vData
End Property
Public Property Get p_revision_nbr() As Variant
    p_revision_nbr = mvarp_revision_nbr
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
Public Property Let p_type_desc(ByVal vData As Variant)
    mvarp_type_desc = vData
End Property
Public Property Get p_type_desc() As Variant
    p_type_desc = mvarp_type_desc
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    mvarp_type_cd = vData
End Property
Public Property Get p_type_cd() As Variant
    p_type_cd = mvarp_type_cd
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
Public Property Let p_standard_tx(ByVal vData As Variant)
    mvarp_standard_tx = vData
End Property
Public Property Get p_standard_tx() As Variant
    p_standard_tx = mvarp_standard_tx
End Property
Public Property Let p_sales_inv_id(ByVal vData As Variant)
    mvarp_sales_inv_id = vData
End Property
Public Property Get p_sales_inv_id() As Variant
    p_sales_inv_id = mvarp_sales_inv_id
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
Public Sub GetSalesInvoice()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice.cls.GetSalesInvoice"
    
    Call SelectSalesInvoice
    
    Exit Sub
error_handler:
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
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objSalesInvoiceItem As Object
    strRoutine_nm = "sales_invoice.cls.Add"
    
    p_sales_tax_pct = FormatNumber(p_sales_tax_pct, 2)
    p_discount_pct = FormatNumber(p_discount_pct, 2)
    
    Call ValidateData("add")
    Call IssueInsertData
    
    Set objSalesInvoiceItem = CreateObject("converge_sales.sales_invoice_item")
    Set objSalesInvoiceItem.o_error = o_error
    Set objSalesInvoiceItem.o_trans_data = o_trans_data
    objSalesInvoiceItem.p_ship_id = p_ship_id
    objSalesInvoiceItem.p_sales_inv_id = p_sales_inv_id
    
    objSalesInvoiceItem.GenerateItems
    
    Set objSalesInvoiceItem = Nothing
    
    Exit Sub
error_handler:
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
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice.cls.Change"
    
    Call ValidateData("change")
    Call IssueUpdateData
    
    Exit Sub
error_handler:
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
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice.cls.Delete"
    
    Call IssueDeleteData
    
    Exit Sub
error_handler:
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Method allows a sales invoice to be automtically
'               generated from a shipment
'  Arguments:   None
'     Return:   Nothing
'//////////////////////////////////////////////////////////////////////////////
Public Sub Generate()
    On Error GoTo error_handler
    
    Dim objShipment As Object
    Dim objSalesOrder As Object
    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice.cls.Delete"
    
    'Set objShipment = CreateObject("converge_sales.shipment")
    'Set objShipment.o_error = o_error
    'Set objShipment.o_trans_data = o_trans_data
    'objShipment.p_ship_id = p_ship_id
    'Call objShipment.GetShipment
    
    Set objSalesOrder = CreateObject("converge_sales.sales_order")
    Set objSalesOrder.o_error = o_error
    Set objSalesOrder.o_trans_data = o_trans_data
    objSalesOrder.p_sales_ord_id = p_sales_ord_id
    Call objSalesOrder.getSalesOrder
    
    p_sales_inv_id = fGetId("sales_inv_id", o_error, o_trans_data)
    p_status_cd = "OP"
    
    ' What are valid sales invoice type cd
    p_type_cd = "GE"
    
    ' objSalesOrder Values
    p_discount_pct = objSalesOrder.p_discount_pct
    p_payment_terms_cd = objSalesOrder.p_payment_terms_cd
    p_sales_tax_pct = objSalesOrder.p_sales_tax_pct
    
    ' Where do we get the standard tx
    p_standard_tx = "Generated Sales Invoice"
    
    Call Add
    
    Set objSalesOrder = Nothing
    'Set objShipment = Nothing
    
    
    Exit Sub
error_handler:
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



Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_invoice.cls.IssueDeleteData"

    strSql = "delete from sales_invoice "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    strSql = "delete from sales_invoice_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
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

Private Sub ValidateData(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice.cls.ValidateData"
    
    If fValuePresent(p_status_cd) = True And _
       fGetDesc("sales_invoice_status_cd", p_status_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1200
        Exit Sub
    End If
    
    If fValuePresent(p_currency_cd) = True And _
       fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "2000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1205
        Exit Sub
    End If
    
        
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_invoice.cls.IssueInsertData"
    
    p_sales_inv_id = fGetId("sales_inv_id", o_error, o_trans_data)
    
    strSql = "insert into sales_invoice ( "
    strSql = strSql & "  sales_inv_id"
    strSql = strSql & ", ship_id"
    strSql = strSql & ", standard_tx"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", revision_nbr"
    strSql = strSql & ", discount_pct"
    strSql = strSql & ", sales_tax_pct"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_inv_id, "N")
    strSql = strSql & "," & fInsertVariable(p_ship_id, "N")
    strSql = strSql & "," & fInsertVariable(p_standard_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & ",1"
    strSql = strSql & "," & fInsertVariable(p_discount_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_tax_pct, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_invoice.cls.IssueUpdateData"
    
    strSql = "update sales_invoice set "
    strSql = strSql & " " & fUpdateVariable("standard_tx", p_standard_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & ", revision_nbr = revision_nbr + 1 "
    strSql = strSql & "," & fUpdateVariable("discount_pct", p_discount_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_tax_pct", p_sales_tax_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
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

Private Sub SelectSalesInvoice()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "sales_invoice.cls.SelectSalesInvoice"

    If IsNull(p_sales_inv_id) = True Then
       p_sales_inv_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM sales_invoice "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    End If
          
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
       p_sales_inv_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_sales_inv_id = objRecordset("sales_inv_id")
        p_standard_tx = objRecordset("standard_tx")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_status_cd = objRecordset("status_cd")
        p_revision_nbr = objRecordset("revision_nbr")
        p_discount_pct = objRecordset("discount_pct")
        p_sales_tax_pct = objRecordset("sales_tax_pct")
        p_ship_id = objRecordset("ship_id")
        
        objRecordset.Close
        Set objRecordset = Nothing
        
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("sales_invoice_status_cd", p_status_cd, o_error, o_trans_data)
    End If
    
    Exit Sub
error_handler:
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

Public Sub getSalesInvoices()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.GetSalesInvoices"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    
   strSql = "SELECT  " & _
                " a.sales_inv_id, " & _
                " b.cust_nbr, " & _
                " c.code_desc, " & _
                " a.ship_id, " & _
                " a.type_cd " & _
                " FROM sales_invoice a, customer b, code_desc c, shipment d" & _
                " WHERE d.cust_id = b.cust_id " & _
                "   AND d.ship_id = a.ship_id " & _
                "   AND (" & fWhereVariable("c.field_nm", "sales_invoice_status_cd", "S", "=") & _
                "   AND c.code = a.status_cd)"
    
    If IsNull(p_sales_inv_id) = True Or p_sales_inv_id = 0 Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.sales_inv_id", p_sales_inv_id, "N", "=")
    End If
                
    If IsNull(p_ship_id) = True Or p_ship_id = 0 Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.ship_id", p_ship_id, "N", "=")
    End If
    
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
               
    strSql = strSql & " Order by a.sales_inv_id "
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
      .p_err_cd = "3000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub




