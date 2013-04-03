VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_quote_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarp_line_amnt As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_requested_qty As Variant
Private mvarp_sales_quote_id As Variant
Private mvarp_uom_cd As Variant
Private mvarp_uom_desc As Variant
Private mvarp_line_type_desc As Variant
Private mvarp_line_type_cd As Variant
Private mvarp_line_nbr As Variant
Private mvarp_line_desc As Variant
Private mvarp_item_price As Variant
Private mvarp_item_nbr As Variant
Private mvarp_misc_item_nbr As Variant
Private mvarp_item_id As Variant
Private mvarp_cust_item_nbr As Variant
Private mvarp_bid_qty As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_item_desc As Variant
Private mvaro_record_set As ADODB.Recordset
'

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_misc_item_nbr(ByVal vData As Variant)
    mvarp_misc_item_nbr = vData
End Property
Public Property Get p_misc_item_nbr() As Variant
        p_misc_item_nbr = mvarp_misc_item_nbr
End Property

Public Property Let p_item_desc(ByVal vData As Variant)
    mvarp_item_desc = vData
End Property
Public Property Get p_item_desc() As Variant
        p_item_desc = mvarp_item_desc
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
Public Property Let p_bid_qty(ByVal vData As Variant)
    mvarp_bid_qty = vData
End Property
Public Property Get p_bid_qty() As Variant
    p_bid_qty = mvarp_bid_qty
End Property
Public Property Let p_cust_item_nbr(ByVal vData As Variant)
    mvarp_cust_item_nbr = vData
End Property
Public Property Get p_cust_item_nbr() As Variant
    p_cust_item_nbr = mvarp_cust_item_nbr
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
Public Property Let p_item_price(ByVal vData As Variant)
    mvarp_item_price = vData
End Property
Public Property Get p_item_price() As Variant
    p_item_price = mvarp_item_price
End Property
Public Property Let p_line_desc(ByVal vData As Variant)
    mvarp_line_desc = vData
End Property
Public Property Get p_line_desc() As Variant
    p_line_desc = mvarp_line_desc
End Property
Public Property Let p_line_nbr(ByVal vData As Variant)
    mvarp_line_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
    p_line_nbr = mvarp_line_nbr
End Property
Public Property Let p_line_type_cd(ByVal vData As Variant)
    mvarp_line_type_cd = vData
End Property
Public Property Get p_line_type_cd() As Variant
    p_line_type_cd = mvarp_line_type_cd
End Property
Public Property Let p_line_type_desc(ByVal vData As Variant)
        mvarp_line_type_desc = vData
End Property
Public Property Get p_line_type_desc() As Variant
        p_line_type_desc = mvarp_line_type_desc
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
Public Property Let p_sales_quote_id(ByVal vData As Variant)
        mvarp_sales_quote_id = vData
End Property
Public Property Get p_sales_quote_id() As Variant
        p_sales_quote_id = mvarp_sales_quote_id
End Property
Public Property Let p_requested_qty(ByVal vData As Variant)
    mvarp_requested_qty = vData
End Property
Public Property Get p_requested_qty() As Variant
        p_requested_qty = mvarp_requested_qty
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
Public Sub GetSalesQuoteItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote_item.cls.GetSalesQuoteItem"
  
    Call SelectSalesQuoteItem

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
    strRoutine_nm = "sales_quote_item.cls.Add"
  
    Call ValidateData("add")
    Call IssueInsertData
    
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
    strRoutine_nm = "sales_quote_item.cls.Change"
    
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
    strRoutine_nm = "sales_quote.cls.Delete"
    
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

Public Sub DeleteLineItems()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_quote_item.cls.IssueDeleteData"

    strSql = "delete from sales_quote_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    
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

Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_quote_item.cls.IssueDeleteData"

    strSql = "delete from sales_quote_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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
    Err.clear
End Sub

Private Sub ValidateData(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strRoutine_nm = "sales_quote_item.cls.ValidateData"
    
    If sFunction_cd = "add" And Len(p_line_nbr) > 0 Then
        strSql = "SELECT * FROM sales_quote_item"
        strSql = strSql & " WHERE " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
        strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
        
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
        
        If objRecordset.RecordCount > 0 Then
            DeleteLineItems
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1590"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 3426
            Exit Sub
        End If
        
    End If
    
    If fValuePresent(p_line_type_cd) = True And _
       fGetDesc("line_type_cd", p_line_type_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1225
        Exit Sub
    End If
    
    If fValuePresent(p_bid_qty) = True And _
       IsNumeric(p_bid_qty) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1230
        Exit Sub
    End If
        
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2000"
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
    strRoutine_nm = "sales_quote_item.cls.IssueInsertData"
    
    'p_item_id = fGetItemId(p_item_nbr, o_error, o_trans_data)
    If Len(p_line_nbr) > 0 Then
    Else
        Call GetLineNbr
    End If
    
    strSql = "insert into sales_quote_item( "
    strSql = strSql & "  sales_quote_id"
    strSql = strSql & ", line_nbr"
    strSql = strSql & ", bid_qty"
    strSql = strSql & ", cust_item_nbr"
    strSql = strSql & ", item_id"
    strSql = strSql & ", misc_item_nbr"
    strSql = strSql & ", requested_qty"
    strSql = strSql & ", item_desc"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", uom_cd"
    strSql = strSql & ", item_price"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_quote_id, "N")
    strSql = strSql & "," & fInsertVariable(p_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_bid_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_misc_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_requested_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_line_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_line_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_uom_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_item_price, "N")
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
      .p_err_cd = "2800"
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
    strRoutine_nm = "sales_quote_item.cls.IssueUpdateData"
    
    strSql = "update sales_quote_item set "
    strSql = strSql & " " & fUpdateVariable("bid_qty", p_bid_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("cust_item_nbr", p_cust_item_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_id", p_item_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("requested_qty", p_requested_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_desc", p_line_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_price", p_item_price, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    strSql = strSql & " and " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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

Private Sub GetLineNbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim objRecordset As ADODB.Recordset
    Dim strRoutine_nm As String
    Dim li_line_nbr As Double
    Dim dblUpdateQty As Double
    Dim strSql As String

    strRoutine_nm = "sales_quote_item.cls.GetLineNbr"
    
    strSql = "SELECT line_nbr FROM sales_quote_item " & _
               " WHERE " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=") & _
               " ORDER BY line_nbr desc "
               
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
        
    If objRecordset.EOF = True Then
       p_line_nbr = 1
       Exit Sub
    Else
        li_line_nbr = objRecordset("line_nbr")
    End If
    
    If IsNull(li_line_nbr) = True Then
       p_line_nbr = 1
    Else
       p_line_nbr = li_line_nbr + 1
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
error_handler:
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

Private Sub SelectSalesQuoteItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    strRoutine_nm = "sales_quote_item.cls.SelectSalesQuoteItem"

    If IsNull(p_line_nbr) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM sales_quote_item "
        strSql = strSql & "WHERE "
        strSql = strSql & "   " & fWhereVariable("sales_quote_id", p_sales_quote_id, "S", "=")
   
   Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM sales_quote_item "
        strSql = strSql & "WHERE "
        strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
        strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        objRecordset.Close
        Set objRecordset = Nothing
        p_line_nbr = Null
    Else
        p_line_nbr = objRecordset("line_nbr")
        p_sales_quote_id = objRecordset("sales_quote_id")
        p_bid_qty = objRecordset("bid_qty")
        p_cust_item_nbr = objRecordset("cust_item_nbr")
        p_item_id = objRecordset("item_id")
        p_requested_qty = objRecordset("requested_qty")
        p_line_desc = objRecordset("item_desc")
        p_line_type_cd = objRecordset("type_cd")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_uom_cd = objRecordset("uom_cd")
        p_item_price = objRecordset("item_price")
        p_misc_item_nbr = objRecordset("misc_item_nbr")
        
        
        objRecordset.Close
        Set objRecordset = Nothing
        
        If Len(p_item_id) > 0 Then
            p_item_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
        End If
        p_uom_desc = fGetDesc("uom_cd", p_uom_cd, o_error, o_trans_data)
        p_line_type_desc = fGetDesc("line_type_cd", p_line_type_cd, o_error, o_trans_data)
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_line_amnt = p_item_price * p_bid_qty
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


Public Sub getSalesQuoteItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Quotes.cls.getSalesQuotesItems"
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    
    
    strSql = "SELECT * " & _
                " FROM sales_quote_item a " & _
                " WHERE " & _
                " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    
    strSql = strSql & " ORDER BY a.line_nbr "
                
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        'objRecordset.Close
        'Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
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

