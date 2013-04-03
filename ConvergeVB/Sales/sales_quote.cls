VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_quote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_bill_name As Object
Private mvaro_bill_address As Object
Private mvarp_bypass_nm_addr_fl As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_sales_quote_id As Variant
Private mvarp_fax_nbr As Variant
Private mvarp_exp_dt As Variant
Private mvarp_email_addr As Variant
Private mvarp_cust_id As Variant
Private mvarp_currency_cd As Variant
Private mvarp_buyer_id As Variant
Private mvarp_issued_dt As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_rfq_nbr As Variant
Private mvarp_type_cd As Variant
Private mvarp_type_desc As Variant
Private mvarp_sales_rep_assoc_id As Variant
Private mvarp_sales_rep_assoc_nbr As Variant
Private mvarp_net_amnt As Variant
Private mvarp_sales_tax_pct As Variant
Private mvarp_sales_tax As Variant
Private mvarp_sales_cd As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_tot_amnt As Variant
Private mvarp_currency_desc As Variant
Private mvarp_requested_dt As Variant
Private mvarp_revision_nbr As Variant
Private mvarp_cust_nbr As Variant
Private mvarp_copy_tx As Variant
Private mvarp_quote_tx As Variant
Private mvarp_start_dt As Variant
Private mvarp_end_dt As Variant
Private mvaro_record_set As adodb.Recordset
'

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_end_dt(ByVal vData As Variant)
    mvarp_end_dt = vData
End Property
Public Property Get p_end_dt() As Variant
        p_end_dt = mvarp_end_dt
End Property

Public Property Let p_start_dt(ByVal vData As Variant)
    mvarp_start_dt = vData
End Property
Public Property Get p_start_dt() As Variant
        p_start_dt = mvarp_start_dt
End Property

Public Property Let p_quote_tx(ByVal vData As Variant)
    mvarp_quote_tx = vData
End Property
Public Property Get p_quote_tx() As Variant
        p_quote_tx = mvarp_quote_tx
End Property

Public Property Let p_copy_tx(ByVal vData As Variant)
    mvarp_copy_tx = vData
End Property
Public Property Get p_copy_tx() As Variant
        p_copy_tx = mvarp_copy_tx
End Property

Public Property Let p_cust_nbr(ByVal vData As Variant)
    mvarp_cust_nbr = vData
End Property
Public Property Get p_cust_nbr() As Variant
        p_cust_nbr = mvarp_cust_nbr
End Property
Public Property Let p_revision_nbr(ByVal vData As Variant)
    mvarp_revision_nbr = vData
End Property
Public Property Get p_revision_nbr() As Variant
        p_revision_nbr = mvarp_revision_nbr
End Property
Public Property Let p_requested_dt(ByVal vData As Variant)
    mvarp_requested_dt = vData
End Property
Public Property Get p_requested_dt() As Variant
        p_requested_dt = mvarp_requested_dt
End Property
Public Property Let p_currency_desc(ByVal vData As Variant)
    mvarp_currency_desc = vData
End Property
Public Property Get p_currency_desc() As Variant
        p_currency_desc = mvarp_currency_desc
End Property
Public Property Let p_tot_amnt(ByVal vData As Variant)
    mvarp_tot_amnt = vData
End Property
Public Property Get p_tot_amnt() As Variant
  p_tot_amnt = mvarp_tot_amnt
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
Public Property Let p_sales_cd(ByVal vData As Variant)
    mvarp_sales_cd = vData
End Property
Public Property Get p_sales_cd() As Variant
    p_sales_cd = mvarp_sales_cd
End Property
Public Property Let p_sales_tax(ByVal vData As Variant)
    mvarp_sales_tax = vData
End Property
Public Property Get p_sales_tax() As Variant
    p_sales_tax = mvarp_sales_tax
End Property
Public Property Let p_sales_tax_pct(ByVal vData As Variant)
    mvarp_sales_tax_pct = vData
End Property
Public Property Get p_sales_tax_pct() As Variant
    p_sales_tax_pct = mvarp_sales_tax_pct
End Property
Public Property Let p_net_amnt(ByVal vData As Variant)
    mvarp_net_amnt = vData
End Property
Public Property Get p_net_amnt() As Variant
    p_net_amnt = mvarp_net_amnt
End Property
Public Property Let p_sales_rep_assoc_nbr(ByVal vData As Variant)
    mvarp_sales_rep_assoc_nbr = vData
End Property
Public Property Get p_sales_rep_assoc_nbr() As Variant
    p_sales_rep_assoc_nbr = mvarp_sales_rep_assoc_nbr
End Property
Public Property Let p_sales_rep_assoc_id(ByVal vData As Variant)
    mvarp_sales_rep_assoc_id = vData
End Property
Public Property Get p_sales_rep_assoc_id() As Variant
    p_sales_rep_assoc_id = mvarp_sales_rep_assoc_id
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
Public Property Let p_rfq_nbr(ByVal vData As Variant)
    mvarp_rfq_nbr = vData
End Property
Public Property Get p_rfq_nbr() As Variant
    p_rfq_nbr = mvarp_rfq_nbr
End Property
Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    mvarp_payment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
    p_payment_terms_cd = mvarp_payment_terms_cd
End Property
Public Property Let p_issued_dt(ByVal vData As Variant)
    mvarp_issued_dt = vData
End Property
Public Property Get p_issued_dt() As Variant
    p_issued_dt = mvarp_issued_dt
End Property
Public Property Let p_buyer_id(ByVal vData As Variant)
    mvarp_buyer_id = vData
End Property
Public Property Get p_buyer_id() As Variant
    p_buyer_id = mvarp_buyer_id
End Property
Public Property Let p_currency_cd(ByVal vData As Variant)
    mvarp_currency_cd = vData
End Property
Public Property Get p_currency_cd() As Variant
    p_currency_cd = mvarp_currency_cd
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
    p_cust_id = mvarp_cust_id
End Property
Public Property Let p_email_addr(ByVal vData As Variant)
    mvarp_email_addr = vData
End Property
Public Property Get p_email_addr() As Variant
    p_email_addr = mvarp_email_addr
End Property
Public Property Let p_exp_dt(ByVal vData As Variant)
    mvarp_exp_dt = vData
End Property
Public Property Get p_exp_dt() As Variant
    p_exp_dt = mvarp_exp_dt
End Property
Public Property Let p_fax_nbr(ByVal vData As Variant)
    mvarp_fax_nbr = vData
End Property
Public Property Get p_fax_nbr() As Variant
    p_fax_nbr = mvarp_fax_nbr
End Property
Public Property Let p_sales_quote_id(ByVal vData As Variant)
    mvarp_sales_quote_id = vData
End Property
Public Property Get p_sales_quote_id() As Variant
    p_sales_quote_id = mvarp_sales_quote_id
End Property
Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    mvarp_payment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
    p_payment_terms_desc = mvarp_payment_terms_desc
End Property
Public Property Let p_bypass_nm_addr_fl(ByVal vData As Variant)
    mvarp_bypass_nm_addr_fl = vData
End Property
Public Property Get p_bypass_nm_addr_fl() As Variant
    p_bypass_nm_addr_fl = mvarp_bypass_nm_addr_fl
End Property
Public Property Set o_bill_address(ByVal vData As Object)
    Set mvaro_bill_address = vData
End Property
Public Property Get o_bill_address() As Object
    Set o_bill_address = mvaro_bill_address
End Property
Public Property Set o_bill_name(ByVal vData As Object)
    Set mvaro_bill_name = vData
End Property
Public Property Get o_bill_name() As Object
    Set o_bill_name = mvaro_bill_name
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
Public Sub GetSalesQuote()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.Getsales_quote"
    
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_address = CreateObject("converge.address")
    
    Call SelectSalesQuote
    
    If IsNull(p_sales_quote_id) = True Then
        o_bill_name.p_nm_id = Null
        o_bill_address.p_addr_id = Null
    Else
        Set o_bill_name.o_trans_data = o_trans_data
        Set o_bill_name.o_error = o_error
        o_bill_name.p_nm_id = Null
        o_bill_name.p_obj_nm = "sales_quote"
        o_bill_name.p_obj_id = p_sales_quote_id
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_name.p_def_nm_fl = "Y"
        Call o_bill_name.getDefName
                      
        Set o_bill_address.o_trans_data = o_trans_data
        Set o_bill_address.o_error = o_error
        o_bill_address.p_addr_id = Null
        o_bill_address.p_obj_nm = "sales_quote"
        o_bill_address.p_obj_id = p_sales_quote_id
        o_bill_address.p_addr_type_cd = "BILL"
        o_bill_address.p_def_addr_fl = "Y"
        Call o_bill_address.getDefAddress
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

Public Sub Add()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.Addsales_quote"
  
    Call ValidateData("add")
    Call IssueInsertData
    
    If o_bill_name.p_nm = Empty Then
    Else
        o_bill_name.p_obj_nm = "sales_quote"
        o_bill_name.p_obj_id = p_sales_quote_id
        Call AddName(o_bill_name, o_trans_data, o_error)
    End If
    
    If o_bill_address.p_Street_addr = Empty And o_bill_address.p_City_nm = Empty And o_bill_address.p_State_cd = Empty And o_bill_address.p_Zip_cd = Empty And o_bill_address.p_Country_cd = Empty Then
    Else
        o_bill_address.p_obj_nm = "sales_quote"
        o_bill_address.p_obj_id = p_sales_quote_id
        Call AddAddr(o_bill_address, o_trans_data, o_error)
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

Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.Change"
    
    Call ValidateData("change")
    Call IssueUpdateData
    
    If p_bypass_nm_addr_fl = "Y" Then
    Else
        If IsNull(o_bill_name.p_nm_id) = True Or o_bill_name.p_nm_id = Empty Then
            If IsNull(o_bill_name.p_nm) = True Or o_bill_name.p_nm = Empty Then
            Else
                Call AddName(o_bill_name, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_bill_name.p_nm) = True Then
                Call DeleteName(o_bill_name, o_trans_data, o_error)
            Else
                Call ChangeName(o_bill_name, o_trans_data, o_error)
            End If
        End If
        
        If IsNull(o_bill_address.p_addr_id) = True Or o_bill_address.p_addr_id = Empty Then
            If (IsNull(o_bill_address.p_Street_addr) = True Or o_bill_address.p_Street_addr = Empty) And _
               (IsNull(o_bill_address.p_City_nm) = True Or o_bill_address.p_City_nm = Empty) And _
               (IsNull(o_bill_address.p_Zip_cd) = True Or o_bill_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_bill_address, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_bill_address.p_Street_addr) = True And _
               IsNull(o_bill_address.p_City_nm) = True And _
               IsNull(o_bill_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_bill_address, o_trans_data, o_error)
            Else
                Call ChangeAddr(o_bill_address, o_trans_data, o_error)
            End If
        End If
    End If
    
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

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.Delete"
    
    Call IssueDeleteData
    o_bill_address.p_obj_id = p_sales_quote_id
    o_bill_address.p_obj_nm = "Sales_quote"
    Call DeleteAddr(o_bill_address, o_trans_data, o_error)
    o_bill_name.p_obj_id = p_sales_quote_id
    o_bill_name.p_obj_nm = "Sales_quote"
    Call DeleteName(o_bill_name, o_trans_data, o_error)
       
    
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

Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_quote.cls.IssueDeleteData"

    strSql = "delete from sales_quote "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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

Private Sub ValidateData(sFunction_cd)
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.ValidateData"
    
    If fValuePresent(p_sales_tax_pct) = True And _
       (IsNumeric(p_sales_tax_pct) = False Or _
        p_sales_tax_pct < 0 Or _
        p_sales_tax_pct > 1) Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "3800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1030
        Exit Sub
    End If
    
    Call ValidateAssoc_nbr
        
    Exit Sub
error_handler:
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

Private Sub ValidateAssoc_nbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "sales_quote.cls.ValidateAssoc_nbr"
    
    If IsNull(p_sales_rep_assoc_nbr) = True Then
        p_sales_rep_assoc_id = Null
        Exit Sub
    End If
    
    If fValuePresent(p_sales_rep_assoc_nbr) = True Then
    Else
        Exit Sub
    End If
    
    p_sales_rep_assoc_id = fGetAssocId(p_sales_rep_assoc_nbr, o_error, o_trans_data)
    
    If IsNull(p_sales_rep_assoc_id) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1185
        Exit Sub
    End If
    
    Exit Sub
error_handler:
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

Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "sales_quote.cls.IssueInsertData"
    
    p_sales_quote_id = fGetId("sales_quote_id", o_error, o_trans_data)
    
    strSql = "insert into sales_quote ( "
    strSql = strSql & "  sales_quote_id"
    strSql = strSql & ", buyer_id"
    strSql = strSql & ", copy_tx"
    strSql = strSql & ", quote_tx"
    strSql = strSql & ", currency_cd"
    strSql = strSql & ", cust_id"
    strSql = strSql & ", exp_dt"
    strSql = strSql & ", email_addr"
    strSql = strSql & ", fax_nbr"
    strSql = strSql & ", issued_dt"
    strSql = strSql & ", payment_terms_cd"
    strSql = strSql & ", rfq_nbr"
    strSql = strSql & ", requested_dt"
    strSql = strSql & ", revision_nbr"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", sales_rep_assoc_id"
    strSql = strSql & ", sales_tax_pct"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_sales_quote_id, "N")
    strSql = strSql & "," & fInsertVariable(p_buyer_id, "S")
    strSql = strSql & "," & fInsertVariable(p_copy_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_quote_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_currency_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_cust_id, "N")
    strSql = strSql & "," & fInsertVariable(p_exp_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_email_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_fax_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_issued_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_rfq_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_requested_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_revision_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_sales_rep_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_tax_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
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
      .p_err_cd = "4600"
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
    strRoutine_nm = "sales_quote.cls.IssueUpdateData"
    
    strSql = "update sales_quote set "
    strSql = strSql & " " & fUpdateVariable("buyer_id", p_buyer_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("copy_tx", p_copy_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("quote_tx", p_quote_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("currency_cd", p_currency_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("exp_dt", p_exp_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("email_addr", p_email_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("fax_nbr", p_fax_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("issued_dt", p_issued_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("rfq_nbr", p_rfq_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("requested_dt", p_requested_dt, "S", "=")
    strSql = strSql & ", revision_nbr = revision_nbr + 1 "
    strSql = strSql & "," & fUpdateVariable("type_cd", p_type_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("sales_rep_assoc_id", p_sales_rep_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_tax_pct", p_sales_tax_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("sales_quote_id", p_sales_quote_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
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

Private Sub SelectSalesQuote()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "sales_quote.cls.SelectSalesQuote"

    strSql = "SELECT *      "
    strSql = strSql & "FROM sales_quote "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("sales_quote_id", p_sales_quote_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_sales_quote_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_sales_quote_id = objRecordset("sales_quote_id")
        p_buyer_id = objRecordset("buyer_id")
        p_currency_cd = objRecordset("currency_cd")
        p_exp_dt = objRecordset("exp_dt")
        p_email_addr = objRecordset("email_addr")
        p_fax_nbr = objRecordset("fax_nbr")
        p_issued_dt = objRecordset("issued_dt")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_rfq_nbr = objRecordset("rfq_nbr")
        p_requested_dt = objRecordset("requested_dt")
        p_revision_nbr = objRecordset("revision_nbr")
        p_type_cd = objRecordset("type_cd")
        p_sales_rep_assoc_id = objRecordset("sales_rep_assoc_id")
        p_sales_tax_pct = objRecordset("sales_tax_pct")
        p_status_cd = objRecordset("status_cd")
        p_cust_id = objRecordset("cust_id")
        p_quote_tx = objRecordset("quote_tx")
        p_copy_tx = objRecordset("copy_tx")
      
        objRecordset.Close
        Set objRecordset = Nothing
        
        p_currency_desc = fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data)
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_type_desc = fGetDesc("sales_quote_type_cd", p_type_cd, o_error, o_trans_data)
        p_sales_rep_assoc_nbr = fGetAssocNbr(p_sales_rep_assoc_id, o_error, o_trans_data)
        p_status_desc = fGetDesc("sales_quote_status_cd", p_status_cd, o_error, o_trans_data)
        p_cust_nbr = fGetCustNbr(p_cust_id, o_error, o_trans_data)
        ' NEED TO ADD NET AMOUNT, TAX AND TOTAL AMOUNT
        
    End If
    
    Exit Sub
error_handler:
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


Public Sub getCustSalesQuotes()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.GetCustSalesQuotes"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    
   strSql = "SELECT sales_quote.sales_quote_id, customer.cust_nbr, " & _
        " code_desc.code_desc, sales_quote.type_cd," & _
        " associate.assoc_nbr" & _
    " FROM sales_quote LEFT OUTER JOIN" & _
        " associate ON" & _
        " sales_quote.sales_rep_assoc_id = associate.assoc_id LEFT OUTER" & _
        " Join" & _
        " customer ON" & _
        " sales_quote.cust_id = customer.cust_id LEFT OUTER JOIN" & _
        " code_desc ON" & _
        " sales_quote.status_cd = code_desc.code" & _
    " WHERE (code_desc.field_nm = 'sales_quote_status_cd')"
    
    
    If IsNull(p_cust_id) = True Or p_cust_id = 0 Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("sales_quote.cust_id", p_cust_id, "N", "=")
    End If
    
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("sales_quote.status_cd", p_status_cd, "S", "=")
    End If
    
    If IsNull(p_start_dt) = True Or p_start_dt = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("sales_quote.entered_ts", p_start_dt, "D", ">=")
    End If
    
    If IsNull(p_end_dt) = True Or p_start_dt = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("sales_quote.entered_ts", p_end_dt, "D", "<")
    End If
    
    If IsNull(p_sales_rep_assoc_nbr) = True Or p_sales_rep_assoc_nbr = "" Then
    Else
        strSql = strSql & " AND " & fWhereVariable("associate.assoc_nbr", p_sales_rep_assoc_nbr, "S", "=")
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
      .p_err_cd = "5200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


