VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_invoice_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarp_sales_inv_id As Variant
Private mvarp_line_nbr As Variant
Private mvarp_ship_id As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_type_cd As Variant
Private mvarp_type_desc As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_item_id As Variant
Private mvarp_item_nbr As Variant
Private mvarp_cust_item_nbr As Variant
Private mvarp_item_desc As Variant
Private mvarp_uom_cd As Variant
Private mvarp_uom_desc As Variant
Private mvarp_item_price As Variant
Private mvarp_item_qty As Variant
Private mvarp_line_amount As Variant
Private mvaro_record_set As ADODB.Recordset
'

Public Property Let p_ship_id(ByVal vData As Variant)
    mvarp_ship_id = vData
End Property
Public Property Get p_ship_id() As Variant
    p_ship_id = mvarp_ship_id
End Property

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_line_amount(ByVal vData As Variant)
    mvarp_line_amount = vData
End Property
Public Property Get p_line_amount() As Variant
    p_line_amount = mvarp_line_amount
End Property
Public Property Let p_item_qty(ByVal vData As Variant)
    mvarp_item_qty = vData
End Property
Public Property Get p_item_qty() As Variant
    p_item_qty = mvarp_item_qty
End Property
Public Property Let p_item_price(ByVal vData As Variant)
    mvarp_item_price = vData
End Property
Public Property Get p_item_price() As Variant
    p_item_price = mvarp_item_price
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
Public Property Let p_item_desc(ByVal vData As Variant)
    mvarp_item_desc = vData
End Property
Public Property Get p_item_desc() As Variant
    p_item_desc = mvarp_item_desc
End Property
Public Property Let p_cust_item_nbr(ByVal vData As Variant)
    mvarp_cust_item_nbr = vData
End Property
Public Property Get p_cust_item_nbr() As Variant
    p_cust_item_nbr = mvarp_cust_item_nbr
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
Public Property Let p_line_nbr(ByVal vData As Variant)
    mvarp_line_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
    p_line_nbr = mvarp_line_nbr
End Property
Public Property Let p_sales_inv_id(ByVal vData As Variant)
    mvarp_sales_inv_id = vData
End Property
Public Property Get p_sales_inv_id() As Variant
    p_sales_inv_id = mvarp_sales_inv_id
End Property
Public Sub GetSalesInvoiceItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice_item.cls.GetSalesInvoiceItem"
  
    Call SelectSalesInvoiceItem

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
    strRoutine_nm = "sales_invoice_item.cls.Add"
    
    Dim objShipment As shipment
    Dim objSalesInvoice As sales_invoice
    
    Set objSalesInvoice = New sales_invoice
    objSalesInvoice.o_error = Me.o_error
    objSalesInvoice.o_trans_data = Me.o_trans_data
    objSalesInvoice.p_sales_inv_id = Me.p_sales_inv_id
    Call objSalesInvoice.GetSalesInvoice
    
    Call ValidateData("add")
    Call GetLineNbr
    Call fGetItemId(p_item_nbr, o_error, o_trans_data)
    Call IssueInsertData
     
    If p_item_qty > 0 Then
        Set objShipment = New shipment
        objShipment.o_error = Me.o_error
        objShipment.o_trans_data = Me.o_trans_data
        objShipment.p_ship_id = objSalesInvoice.p_ship_id
        Call objShipment.GetShipment
        
        Call addCustItem( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_item_id _
                    , objShipment.p_cust_id _
                    , Me.p_item_price _
                    , Me.p_item_desc _
                    , Me.p_cust_item_nbr _
                    , objShipment.p_planned_ship_dt _
                    , objShipment.p_planned_del_dt _
                    , objShipment.p_drop_ship_fl _
                    )
    End If
    
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
    strRoutine_nm = "sales_invoice_item.cls.Change"
    
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


Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_invoice_item.cls.IssueDeleteData"

    strSql = "delete from sales_invoice_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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


Private Sub ValidateData(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice_item.cls.ValidateData"
        
    Exit Sub
error_handler:
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


Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_invoice_item.cls.IssueInsertData"
    
    strSql = "insert into sales_invoice_item( "
    strSql = strSql & "  sales_inv_id"
    strSql = strSql & ", line_nbr"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", item_id"
    strSql = strSql & ", item_desc"
    strSql = strSql & ", item_price"
    strSql = strSql & ", cust_item_nbr"
    strSql = strSql & ", uom_cd"
    strSql = strSql & ", item_qty"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_inv_id, "N")
    strSql = strSql & "," & fInsertVariable(p_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_item_price, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_uom_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_item_qty, "N")
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
      .p_err_cd = "2200"
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
    strRoutine_nm = "sales_invoice_item.cls.IssueUpdateData"
    
    Call fGetItemId(p_item_nbr, o_error, o_trans_data)
    
    strSql = "update sales_invoice_item set "
    strSql = strSql & " " & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_id", p_item_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_desc", p_item_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_price", p_item_price, "N", "=")
    strSql = strSql & "," & fUpdateVariable("uom_cd", p_uom_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_qty", p_item_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
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

Private Sub GetLineNbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim objRecordSet As ADODB.Recordset
    Dim strRoutine_nm As String
    Dim li_line_nbr As Double
    Dim dblUpdateQty As Double
    Dim strSql As String

    strRoutine_nm = "sales_invoice_item.cls.GetLineNbr"
    
    strSql = "SELECT line_nbr FROM sales_invoice_item " & _
               " WHERE " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=") & _
               " ORDER BY line_nbr desc "
               
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    End If
    
    li_line_nbr = objRecordSet("line_nbr")
    
    If IsNull(li_line_nbr) = True Then
       p_line_nbr = 1
    Else
       p_line_nbr = li_line_nbr + 1
    End If
    
    objRecordSet.Close
    Set objRecordSet = Nothing
    
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

Private Sub SelectSalesInvoiceItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim lv_amount As Variant
    Dim objRecordSet As ADODB.Recordset
    
    strRoutine_nm = "sales_invoice_item.cls.SelectSalesInvoiceLine"

    strSql = "SELECT * "
    strSql = strSql & "FROM sales_invoice_item "
    strSql = strSql & " WHERE "
    strSql = strSql & " " & fWhereVariable("sales_inv_id", p_sales_inv_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    End If
    
    If objRecordSet.EOF = True Then
        p_line_nbr = Null
        objRecordSet.Close
        Set objRecordSet = Nothing
    Else
        p_sales_inv_id = objRecordSet("sales_inv_id")
        p_line_nbr = objRecordSet("line_nbr")
        p_payment_terms_cd = objRecordSet("payment_terms_cd")
        p_item_id = objRecordSet("item_id")
        p_cust_item_nbr = objRecordSet("cust_item_nbr")
        p_item_desc = objRecordSet("item_desc")
        p_item_price = objRecordSet("item_price")
        p_uom_cd = objRecordSet("uom_cd")
        p_item_qty = objRecordSet("item_qty")
        
        objRecordSet.Close
        Set objRecordSet = Nothing
        
        'Let p_type_cd = objRecordset("type_cd")
        'Let p_type_desc = fGetDesc("item_type_cd", objRecordset("type_cd"), o_error, o_trans_data)
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_item_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
        p_line_amount = p_item_qty * p_item_price
        p_uom_desc = fGetDesc("uom_cd", p_uom_cd, o_error, o_trans_data)
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

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice_item.cls.Delete"
    
    Call IssueDeleteData
   
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

Public Sub getSalesInvoiceItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_invoice_item.cls.getSalesInvoiceItems"
    Dim objRecordSet As ADODB.Recordset
    Dim strSql As String
   
    
   strSql = "SELECT  " & _
                " a.sales_inv_id, a.cust_item_nbr, a.uom_cd, " & _
                " a.line_nbr, " & _
                " b.item_nbr, " & _
                " b.item_desc, " & _
                " a.item_qty, " & _
                " a.item_price " & _
                " FROM sales_invoice_item a, item b " & _
                " WHERE a.item_id = b.item_id " & _
                "   AND " & fWhereVariable("a.sales_inv_id", p_sales_inv_id, "N", "=")

    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordSet
    End If
     
    Exit Sub
error_handler:
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

Public Sub GenerateItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice_item.cls.GenerateItems"
    
    Dim strSql As String, _
        objRecordSet As ADODB.Recordset, _
        objCustItem As cust_item
    
    strSql = " SELECT a.line_nbr, b.cust_item_nbr, b.item_desc, a.item_id, " & _
             "        b.item_price, a.loaded_qty, b.payment_terms_cd, " & _
             "        b.uom_cd, c.cust_id, d.planned_ship_dt, d.planned_del_dt, d.drop_ship_fl " & _
             "  FROM shipment_item a, sales_order_item b , sales_order c, shipment d" & _
             " WHERE a.sales_ord_id = b.sales_ord_id " & _
             "   and a.sales_ord_id = c.sales_ord_id " & _
             "   AND a.line_nbr = b.line_nbr " & _
             "   AND " & fWhereVariable("a.ship_id", p_ship_id, "N", "=") & _
             "   AND " & fWhereVariable("d.ship_id", p_ship_id, "N", "=") & _
             "  ORDER BY a.line_nbr "

    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Err.clear
        Exit Sub
    Else
        Do Until objRecordSet.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
            p_line_nbr = objRecordSet("line_nbr")
            p_cust_item_nbr = objRecordSet("cust_item_nbr")
            p_item_desc = objRecordSet("item_desc")
            p_item_id = objRecordSet("item_id")
            p_item_price = objRecordSet("item_price")
            p_item_qty = objRecordSet("loaded_qty")
            p_payment_terms_cd = objRecordSet("payment_terms_cd")
            p_uom_cd = objRecordSet("uom_cd")
            Call IssueInsertData
            If p_item_qty > 0 Then
                Call addCustItem( _
                              o_error _
                            , o_trans_data _
                            , objRecordSet("item_id") _
                            , objRecordSet("cust_Id") _
                            , objRecordSet("item_price") _
                            , objRecordSet("item_desc") _
                            , objRecordSet("cust_item_nbr") _
                            , objRecordSet("planned_ship_dt") _
                            , objRecordSet("planned_del_dt") _
                            , objRecordSet("drop_ship_fl") _
                            )
            End If
                                        'If the customer has no initial purchase
                                        'date for the item, record it now using todays
                                        'date.
            objRecordSet.MoveNext
        Loop
        
    End If
    
    Exit Sub
error_handler:
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


Private Sub addCustItem( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblItemId As Double _
            , ByVal v_dblCustId As Double _
            , ByVal v_dblCustItemPrice As Double _
            , ByVal v_strCustItemDesc As Variant _
            , ByVal v_strCustItemNbr As Variant _
            , ByVal v_varPlannedShipDt As Variant _
            , ByVal v_varPlannedDelDt As Variant _
            , ByVal v_varDropShipFl As Variant _
            )
    
    On Error GoTo error_handler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_invoice_item.cls.addCustItem"
    
    Dim objCustItem As Object
                                        'If the customer has no initial purchase
                                        'date for the item, record it now using todays
                                        'date.
    If v_varDropShipFl = "Y" Then
        If IsDate(v_varPlannedDelDt) = False Then
            Exit Sub
        End If
    Else
        If IsDate(v_varPlannedShipDt) = False Then
            Exit Sub
        End If
    End If
    
    Set objCustItem = New cust_item
    With objCustItem
        .itemId = v_dblItemId
        .custId = v_dblCustId
        .bypass_getcust_fl = "Y"
        .bypass_getitem_fl = "Y"
        Set .objError = r_objError
        Set .objTransData = r_objTransData
        Call .GetCustItem
        If IsNull(.itemId) Then
                                'If not found, add the cust_item record
            .custId = v_dblCustId
            .itemId = v_dblItemId
            .allowQtyDiscFl = "N"
            .custItemPrice = v_dblCustItemPrice
            .custItemDesc = v_strCustItemDesc
            .custItemNbr = v_strCustItemNbr
            If v_varDropShipFl = "Y" Then
                .initialPurchaseDt = CDate(Format(v_varPlannedDelDt, "mm/dd/yyyy"))
            Else
                .initialPurchaseDt = CDate(Format(v_varPlannedShipDt, "mm/dd/yyyy"))
            End If
            Call .Add
        Else
                                'If it was found, and the date is null
                                'update the record
            If IsNull(.initialPurchaseDt) = True Then
                If v_varDropShipFl = "Y" Then
                    .initialPurchaseDt = CDate(Format(v_varPlannedDelDt, "mm/dd/yyyy"))
                Else
                    .initialPurchaseDt = CDate(Format(v_varPlannedShipDt, "mm/dd/yyyy"))
                End If
                Call .Change
            End If
        End If
    End With
    
    Exit Sub
error_handler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "3600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
