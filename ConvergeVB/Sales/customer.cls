VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_stnd_name As Object
Private mvaro_bill_name As Object
Private mvaro_ship_name As Object
Private mvaro_stnd_address As Object
Private mvaro_bill_address As Object
Private mvaro_ship_address As Object
Private mvaro_cust_group As Object
Private mvaro_old_customer As Object
Private mvarp_bypass_nm_addr_fl As Variant
Private mvarp_carrier_nm As Variant
Private mvarp_class_cd As Variant
Private mvarp_class_desc As Variant
Private mvarp_credit_limit_am As Variant
Private mvarp_consignment_fl As Variant
Private mvarp_cons_inv_loc_id As Variant
Private mvarp_currency_cd As Variant
Private mvarp_currency_desc As Variant
Private mvarp_cust_grp_id As Variant
Private mvarp_cust_id As Variant
Private mvarp_cust_nbr As Variant
Private mvarp_email_addr As Variant
Private mvarp_fax_nbr As Variant
Private mvarp_fob_terms_cd As Variant
Private mvarp_fob_terms_desc As Variant
Private mvarp_freight_cd As Variant
Private mvarp_freight_desc As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_phone_nbr As Variant
Private mvarp_registration_dt As Variant
Private mvarp_sales_order_limit_am As Variant
Private mvarp_sales_rep_assoc_id As Variant
Private mvarp_sales_rep_assoc_nbr As Variant
Private mvarp_cust_rep_assoc_id As Variant
Private mvarp_cust_rep_assoc_nbr As Variant
Private mvarp_sales_tax_pct As Variant
Private mvarp_sales_terr_cd As Variant
Private mvarp_sales_terr_desc As Variant
Private mvarp_ship_complete_var_pct As Variant
Private mvarp_sic_cd As Variant
Private mvarp_sic_desc As Variant
Private mvarp_source_cd As Variant
Private mvarp_source_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_stnd_tx As Variant
Private mvarp_stnd_tx_on_invoice_fl As Variant
Private mvarp_stnd_tx_on_pack_slip_fl As Variant
Private mvarp_tax_exempt_fl As Variant
Private mvarp_tax_id As Variant
Private mvarp_transportation_cd As Variant
Private mvarp_transportation_desc As Variant
Private mvarp_timco_supp_cd As Variant
Private mvarp_web_site_addr As Variant
Private mvarp_ship_acct_nbr As Variant
Private mvaro_record_set As adodb.Recordset
'


Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_ship_complete_var_pct(ByVal vData As Variant)
    mvarp_ship_complete_var_pct = vData
End Property
Public Property Get p_ship_complete_var_pct() As Variant
        p_ship_complete_var_pct = mvarp_ship_complete_var_pct
End Property
Public Property Let p_consignment_fl(ByVal vData As Variant)
    mvarp_consignment_fl = vData
End Property
Public Property Get p_consignment_fl() As Variant
        p_consignment_fl = mvarp_consignment_fl
End Property
Public Property Let p_carrier_nm(ByVal vData As Variant)
    mvarp_carrier_nm = vData
End Property
Public Property Get p_carrier_nm() As Variant
        p_carrier_nm = mvarp_carrier_nm
End Property
Public Property Let p_cons_inv_loc_id(ByVal vData As Variant)
    mvarp_cons_inv_loc_id = vData
End Property
Public Property Get p_cons_inv_loc_id() As Variant
        p_cons_inv_loc_id = mvarp_cons_inv_loc_id
End Property
Public Property Let p_transportation_desc(ByVal vData As Variant)
    mvarp_transportation_desc = vData
End Property
Public Property Get p_transportation_desc() As Variant
        p_transportation_desc = mvarp_transportation_desc
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property
Public Property Let p_source_desc(ByVal vData As Variant)
    mvarp_source_desc = vData
End Property
Public Property Get p_source_desc() As Variant
        p_source_desc = mvarp_source_desc
End Property
Public Property Let p_sic_desc(ByVal vData As Variant)
    mvarp_sic_desc = vData
End Property
Public Property Get p_sic_desc() As Variant
        p_sic_desc = mvarp_sic_desc
End Property
Public Property Let p_sales_terr_desc(ByVal vData As Variant)
    mvarp_sales_terr_desc = vData
End Property
Public Property Get p_sales_terr_desc() As Variant
        p_sales_terr_desc = mvarp_sales_terr_desc
End Property
Public Property Let p_payment_terms_desc(ByVal vData As Variant)
    mvarp_payment_terms_desc = vData
End Property
Public Property Get p_payment_terms_desc() As Variant
        p_payment_terms_desc = mvarp_payment_terms_desc
End Property
Public Property Let p_freight_desc(ByVal vData As Variant)
    mvarp_freight_desc = vData
End Property
Public Property Get p_freight_desc() As Variant
        p_freight_desc = mvarp_freight_desc
End Property
Public Property Let p_fob_terms_desc(ByVal vData As Variant)
    mvarp_fob_terms_desc = vData
End Property
Public Property Set p_fob_terms_desc(ByVal vData As Variant)
    Set mvarp_fob_terms_desc = vData
End Property
Public Property Get p_fob_terms_desc() As Variant
        p_fob_terms_desc = mvarp_fob_terms_desc
End Property
Public Property Let p_currency_desc(ByVal vData As Variant)
    mvarp_currency_desc = vData
End Property
Public Property Get p_currency_desc() As Variant
        p_currency_desc = mvarp_currency_desc
End Property
Public Property Let p_class_desc(ByVal vData As Variant)
    mvarp_class_desc = vData
End Property
Public Property Get p_class_desc() As Variant
        p_class_desc = mvarp_class_desc
End Property
Public Property Let p_sales_rep_assoc_nbr(ByVal vData As Variant)
    mvarp_sales_rep_assoc_nbr = vData
End Property
Public Property Get p_sales_rep_assoc_nbr() As Variant
        p_sales_rep_assoc_nbr = mvarp_sales_rep_assoc_nbr
End Property
'JRP4
Public Property Let p_cust_rep_assoc_nbr(ByVal vData As Variant)
    mvarp_cust_rep_assoc_nbr = vData
End Property
Public Property Get p_cust_rep_assoc_nbr() As Variant
        p_cust_rep_assoc_nbr = mvarp_cust_rep_assoc_nbr
End Property

Public Property Let p_tax_exempt_fl(ByVal vData As Variant)
    mvarp_tax_exempt_fl = vData
End Property
Public Property Get p_tax_exempt_fl() As Variant
        p_tax_exempt_fl = mvarp_tax_exempt_fl
End Property
Public Property Let p_sales_tax_pct(ByVal vData As Variant)
    mvarp_sales_tax_pct = vData
End Property
Public Property Get p_sales_tax_pct() As Variant
        p_sales_tax_pct = mvarp_sales_tax_pct
End Property
Public Property Let p_stnd_tx_on_pack_slip_fl(ByVal vData As Variant)
    mvarp_stnd_tx_on_pack_slip_fl = vData
End Property
Public Property Get p_stnd_tx_on_pack_slip_fl() As Variant
        p_stnd_tx_on_pack_slip_fl = mvarp_stnd_tx_on_pack_slip_fl
End Property
Public Property Let p_stnd_tx_on_invoice_fl(ByVal vData As Variant)
    mvarp_stnd_tx_on_invoice_fl = vData
End Property
Public Property Get p_stnd_tx_on_invoice_fl() As Variant
        p_stnd_tx_on_invoice_fl = mvarp_stnd_tx_on_invoice_fl
End Property
Public Property Let p_stnd_tx(ByVal vData As Variant)
    mvarp_stnd_tx = vData
End Property
Public Property Get p_stnd_tx() As Variant
        p_stnd_tx = mvarp_stnd_tx
End Property
Public Property Let p_payment_terms_cd(ByVal vData As Variant)
    mvarp_payment_terms_cd = vData
End Property
Public Property Get p_payment_terms_cd() As Variant
        p_payment_terms_cd = mvarp_payment_terms_cd
End Property
Public Property Let p_source_cd(ByVal vData As Variant)
    mvarp_source_cd = vData
End Property
Public Property Get p_source_cd() As Variant
        p_source_cd = mvarp_source_cd
End Property
Public Property Let p_transportation_cd(ByVal vData As Variant)
    mvarp_transportation_cd = vData
End Property
Public Property Set p_transportation_cd(ByVal vData As Variant)
    Set mvarp_transportation_cd = vData
End Property
Public Property Get p_transportation_cd() As Variant
        p_transportation_cd = mvarp_transportation_cd
End Property
Public Property Let p_tax_id(ByVal vData As Variant)
    mvarp_tax_id = vData
End Property
Public Property Get p_tax_id() As Variant
        p_tax_id = mvarp_tax_id
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_sic_cd(ByVal vData As Variant)
    mvarp_sic_cd = vData
End Property
Public Property Get p_sic_cd() As Variant
        p_sic_cd = mvarp_sic_cd
End Property
Public Property Let p_sales_terr_cd(ByVal vData As Variant)
    mvarp_sales_terr_cd = vData
End Property
Public Property Get p_sales_terr_cd() As Variant
        p_sales_terr_cd = mvarp_sales_terr_cd
End Property
Public Property Let p_sales_rep_assoc_id(ByVal vData As Variant)
    mvarp_sales_rep_assoc_id = vData
End Property
Public Property Get p_sales_rep_assoc_id() As Variant
        p_sales_rep_assoc_id = mvarp_sales_rep_assoc_id
End Property
'JRP4
Public Property Let p_cust_rep_assoc_id(ByVal vData As Variant)
    mvarp_cust_rep_assoc_id = vData
End Property
Public Property Get p_cust_rep_assoc_id() As Variant
        p_cust_rep_assoc_id = mvarp_cust_rep_assoc_id
End Property

Public Property Let p_sales_order_limit_am(ByVal vData As Variant)
    mvarp_sales_order_limit_am = vData
End Property
Public Property Get p_sales_order_limit_am() As Variant
        p_sales_order_limit_am = mvarp_sales_order_limit_am
End Property
Public Property Let p_registration_dt(ByVal vData As Variant)
    mvarp_registration_dt = vData
End Property
Public Property Get p_registration_dt() As Variant
        p_registration_dt = mvarp_registration_dt
End Property
Public Property Let p_phone_nbr(ByVal vData As Variant)
    mvarp_phone_nbr = vData
End Property
Public Property Get p_phone_nbr() As Variant
        p_phone_nbr = mvarp_phone_nbr
End Property
Public Property Let p_freight_cd(ByVal vData As Variant)
    mvarp_freight_cd = vData
End Property
Public Property Get p_freight_cd() As Variant
        p_freight_cd = mvarp_freight_cd
End Property
Public Property Let p_fob_terms_cd(ByVal vData As Variant)
    mvarp_fob_terms_cd = vData
End Property
Public Property Get p_fob_terms_cd() As Variant
        p_fob_terms_cd = mvarp_fob_terms_cd
End Property
Public Property Let p_fax_nbr(ByVal vData As Variant)
    mvarp_fax_nbr = vData
End Property
Public Property Get p_fax_nbr() As Variant
        p_fax_nbr = mvarp_fax_nbr
End Property
Public Property Let p_email_addr(ByVal vData As Variant)
    mvarp_email_addr = vData
End Property
Public Property Get p_email_addr() As Variant
        p_email_addr = mvarp_email_addr
End Property
Public Property Let p_cust_grp_id(ByVal vData As Variant)
    mvarp_cust_grp_id = vData
End Property
Public Property Get p_cust_grp_id() As Variant
        p_cust_grp_id = mvarp_cust_grp_id
End Property
Public Property Let p_currency_cd(ByVal vData As Variant)
    mvarp_currency_cd = vData
End Property
Public Property Get p_currency_cd() As Variant
        p_currency_cd = mvarp_currency_cd
End Property
Public Property Let p_credit_limit_am(ByVal vData As Variant)
    mvarp_credit_limit_am = vData
End Property
Public Property Get p_credit_limit_am() As Variant
        p_credit_limit_am = mvarp_credit_limit_am
End Property
Public Property Let p_class_cd(ByVal vData As Variant)
    mvarp_class_cd = vData
End Property
Public Property Get p_class_cd() As Variant
        p_class_cd = mvarp_class_cd
End Property
Public Property Let p_cust_nbr(ByVal vData As Variant)
    mvarp_cust_nbr = vData
End Property
Public Property Get p_cust_nbr() As Variant
        p_cust_nbr = mvarp_cust_nbr
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
        p_cust_id = mvarp_cust_id
End Property
   
Public Property Let p_bypass_nm_addr_fl(ByVal vData As Variant)
    mvarp_bypass_nm_addr_fl = vData
End Property
Public Property Get p_bypass_nm_addr_fl() As Variant
        p_bypass_nm_addr_fl = mvarp_bypass_nm_addr_fl
End Property
Public Property Set o_cust_group(ByVal vData As Object)
    Set mvaro_cust_group = vData
End Property
Public Property Get o_cust_group() As Object
    Set o_cust_group = mvaro_cust_group
End Property
Public Property Set o_old_customer(ByVal vData As Object)
    Set mvaro_old_customer = vData
End Property
Public Property Get o_old_customer() As Object
    Set o_old_customer = mvaro_old_customer
End Property
Public Property Set o_ship_address(ByVal vData As Variant)
    Set mvaro_ship_address = vData
End Property
Public Property Get o_ship_address() As Variant
    Set o_ship_address = mvaro_ship_address
End Property
Public Property Set o_bill_address(ByVal vData As Object)
    Set mvaro_bill_address = vData
End Property
Public Property Get o_bill_address() As Object
    Set o_bill_address = mvaro_bill_address
End Property
Public Property Set o_stnd_address(ByVal vData As Object)
    Set mvaro_stnd_address = vData
End Property
Public Property Get o_stnd_address() As Object
    Set o_stnd_address = mvaro_stnd_address
End Property
Public Property Set o_ship_name(ByVal vData As Variant)
    Set mvaro_ship_name = vData
End Property
Public Property Get o_ship_name() As Variant
    Set o_ship_name = mvaro_ship_name
End Property
Public Property Set o_bill_name(ByVal vData As Object)
    Set mvaro_bill_name = vData
End Property
Public Property Get o_bill_name() As Object
    Set o_bill_name = mvaro_bill_name
End Property
Public Property Set o_stnd_name(ByVal vData As Object)
    Set mvaro_stnd_name = vData
End Property
Public Property Get o_stnd_name() As Object
    Set o_stnd_name = mvaro_stnd_name
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
Public Property Let p_timco_supp_cd(ByVal vData As Variant)
    mvarp_timco_supp_cd = vData
End Property
Public Property Get p_timco_supp_cd() As Variant
        p_timco_supp_cd = mvarp_timco_supp_cd
End Property

Public Property Let p_ship_acct_nbr(ByVal vData As Variant)
    mvarp_ship_acct_nbr = vData
End Property
Public Property Get p_ship_acct_nbr() As Variant
        p_ship_acct_nbr = mvarp_ship_acct_nbr
End Property

Public Property Let p_web_site_addr(ByVal vData As Variant)
    mvarp_web_site_addr = vData
End Property
Public Property Get p_web_site_addr() As Variant
        p_web_site_addr = mvarp_web_site_addr
End Property




Public Sub GetCustomer()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.GetCustomer"
    Dim objRecordset As adodb.Recordset
    
    Set o_stnd_name = CreateObject("converge.name")
    Set o_bill_name = CreateObject("converge.name")
    Set o_ship_name = CreateObject("converge.name")
    Set o_stnd_address = CreateObject("converge.address")
    Set o_bill_address = CreateObject("converge.address")
    Set o_ship_address = CreateObject("converge.address")
    Set o_cust_group = CreateObject("converge_sales.cust_group")
    
    Call SelectCustomer
    
    If IsNull(p_cust_id) = True Then
        o_stnd_name.p_nm_id = Null
        o_stnd_address.p_addr_id = Null
        o_bill_name.p_nm_id = Null
        o_bill_address.p_addr_id = Null
        o_ship_name.p_nm_id = Null
        o_ship_address.p_addr_id = Null
    Else
        Set o_stnd_name.o_trans_data = o_trans_data
        Set o_stnd_name.o_error = o_error
        o_stnd_name.p_nm_id = Null
        o_stnd_name.p_obj_nm = "customer"
        o_stnd_name.p_obj_id = p_cust_id
        o_stnd_name.p_nm_type_cd = "STND"
        o_stnd_name.p_def_nm_fl = "Y"
        Call o_stnd_name.getDefName
        
        Set o_bill_name.o_trans_data = o_trans_data
        Set o_bill_name.o_error = o_error
        o_bill_name.p_nm_id = Null
        o_bill_name.p_obj_nm = "customer"
        o_bill_name.p_obj_id = p_cust_id
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_name.p_def_nm_fl = "Y"
        Call o_bill_name.getDefName
        
        Set o_ship_name.o_trans_data = o_trans_data
        Set o_ship_name.o_error = o_error
        o_ship_name.p_nm_id = Null
        o_ship_name.p_obj_nm = "customer"
        o_ship_name.p_obj_id = p_cust_id
        o_ship_name.p_nm_type_cd = "SHIP"
        o_ship_name.p_def_nm_fl = "Y"
        Call o_ship_name.getDefName
        
        Set o_stnd_address.o_trans_data = o_trans_data
        Set o_stnd_address.o_error = o_error
        o_stnd_address.p_addr_id = Null
        o_stnd_address.p_obj_nm = "customer"
        o_stnd_address.p_obj_id = p_cust_id
        o_stnd_address.p_addr_type_cd = "STND"
        o_stnd_address.p_def_addr_fl = "Y"
        Call o_stnd_address.getDefAddress
        
        Set o_bill_address.o_trans_data = o_trans_data
        Set o_bill_address.o_error = o_error
        o_bill_address.p_addr_id = Null
        o_bill_address.p_obj_nm = "customer"
        o_bill_address.p_obj_id = p_cust_id
        o_bill_address.p_addr_type_cd = "BILL"
        o_bill_address.p_def_addr_fl = "Y"
        Call o_bill_address.getDefAddress
        
        Set o_ship_address.o_trans_data = o_trans_data
        Set o_ship_address.o_error = o_error
        o_ship_address.p_addr_id = Null
        o_ship_address.p_obj_nm = "customer"
        o_ship_address.p_obj_id = p_cust_id
        o_ship_address.p_addr_type_cd = "SHIP"
        o_ship_address.p_def_addr_fl = "Y"
        Call o_ship_address.getDefAddress
        
        If IsNull(p_cust_grp_id) = True Then
            o_cust_group.p_cust_grp_id = Null
        Else
            Set o_cust_group.o_trans_data = o_trans_data
            Set o_cust_group.o_error = o_error
            o_cust_group.p_cust_grp_id = p_cust_grp_id
            Call o_cust_group.GetCustGrp
        End If
    End If
    
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

Public Sub AddCustomer()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.AddCustomer"
    
    Call validate("add")
    Call AddCustData
    Call AddCustName(o_stnd_name)
    Call AddCustAddr(o_stnd_address)
    
    If o_bill_name.p_nm = Empty Then
    Else
        Call AddCustName(o_bill_name)
    End If
    
    If o_bill_address.p_Street_addr = Empty And o_bill_address.p_City_nm = Empty And o_bill_address.p_State_cd = Empty And o_bill_address.p_Zip_cd = Empty And o_bill_address.p_Country_cd = Empty Then
    Else
        Call AddCustAddr(o_bill_address)
    End If
    
    If o_ship_name.p_nm = Empty Then
    Else
        Call AddCustName(o_ship_name)
    End If
    
    If o_ship_address.p_Street_addr = Empty And o_ship_address.p_City_nm = Empty And o_ship_address.p_State_cd = Empty And o_ship_address.p_Zip_cd = Empty And o_ship_address.p_Country_cd = Empty Then
    Else
        Call AddCustAddr(o_ship_address)
    End If
    
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_cust_id)
    
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

Public Sub ChangeCustomer()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.ChangeCustomer"
    
    Call validate("change")
    Call IssueUpdateCustData
    
    If p_bypass_nm_addr_fl = "Y" Then
    Else
        Call ChangeCustName(o_stnd_name)
        Call ChangeCustAddr(o_stnd_address)
        
        If IsNull(o_bill_name.p_nm_id) = True Or o_bill_name.p_nm_id = Empty Then
            If IsNull(o_bill_name.p_nm) = True Or o_bill_name.p_nm = Empty Then
            Else
                Call AddCustName(o_bill_name)
            End If
        Else
            If IsNull(o_bill_name.p_nm) = True Then
                Call DeleteCustName(o_bill_name)
            Else
                Call ChangeCustName(o_bill_name)
            End If
        End If
        
        If IsNull(o_bill_address.p_addr_id) = True Or o_bill_address.p_addr_id = Empty Then
            If (IsNull(o_bill_address.p_Street_addr) = True Or o_bill_address.p_Street_addr = Empty) And _
               (IsNull(o_bill_address.p_City_nm) = True Or o_bill_address.p_City_nm = Empty) And _
               (IsNull(o_bill_address.p_Zip_cd) = True Or o_bill_address.p_Zip_cd = Empty) Then
            Else
                Call AddCustAddr(o_bill_address)
            End If
        Else
            If IsNull(o_bill_address.p_Street_addr) = True And _
               IsNull(o_bill_address.p_City_nm) = True And _
               IsNull(o_bill_address.p_Zip_cd) = True Then
                Call DeleteCustAddr(o_bill_address)
            Else
                Call ChangeCustAddr(o_bill_address)
            End If
        End If
        
        If IsNull(o_ship_name.p_nm_id) = True Or o_ship_name.p_nm_id = Empty Then
            If IsNull(o_ship_name.p_nm) = True Or o_ship_name.p_nm = Empty Then
            Else
                Call AddCustName(o_ship_name)
            End If
        Else
            If IsNull(o_ship_name.p_nm) = True Then
                Call DeleteCustName(o_ship_name)
            Else
                Call ChangeCustName(o_ship_name)
            End If
        End If
        
        If IsNull(o_ship_address.p_addr_id) = True Or o_ship_address.p_addr_id = Empty Then
            If (IsNull(o_ship_address.p_Street_addr) = True Or o_ship_address.p_Street_addr = Empty) And _
               (IsNull(o_ship_address.p_City_nm) = True Or o_ship_address.p_City_nm = Empty) And _
               (IsNull(o_ship_address.p_Zip_cd) = True Or o_ship_address.p_Zip_cd = Empty) Then
            Else
                Call AddCustAddr(o_ship_address)
            End If
        Else
            If IsNull(o_ship_address.p_Street_addr) = True And _
               IsNull(o_ship_address.p_City_nm) = True And _
               IsNull(o_ship_address.p_Zip_cd) = True Then
                Call DeleteCustAddr(o_ship_address)
            Else
                Call ChangeCustAddr(o_ship_address)
            End If
        End If
    End If
    
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_cust_id)
    
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

Public Sub DeleteCustomer()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.DeleteCustomer"
    
    Set o_old_customer = CreateObject("converge_sales.customer")
    Set o_old_customer.o_error = o_error
    Set o_old_customer.o_trans_data = o_trans_data
    o_old_customer.p_cust_id = p_cust_id
    Call o_old_customer.GetCustomer
    
    Call validate("delete")
    Call IssueDeleteCustData
    Call IssueDeleteCustName
    Call IssueDeleteCustAddr
    Call IssueDeleteContacts
    
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

Private Sub AddCustName(objName)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.AddCustName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "customer"
    objName.p_obj_id = p_cust_id
    Call objName.AddName
    
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
Private Sub ChangeCustName(objName)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.ChangeCustName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "customer"
    Call objName.ChangeName
    
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

Private Sub DeleteCustName(objName)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.DeleteCustName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_nm = "customer"
    Call objName.DeleteName
    
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

Private Sub AddCustAddr(objAddress)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.AddCustAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_id = p_cust_id
    objAddress.p_obj_nm = "customer"
    Call objAddress.AddAddress
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ChangeCustAddr(objAddress)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.ChangeCustAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_nm = "customer"
    Call objAddress.ChangeAddress
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub DeleteCustAddr(objAddress)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.DeleteCustAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_nm = "customer"
    Call objAddress.DeleteAddress
    
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

Private Sub AddCustData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.AddCustData"
    
    Me.p_cust_id = fGetId("cust_id", o_error, o_trans_data)
                    'all customers are added in the credit hold status.
                    'they must then be taken off of credit hold
    Me.p_status_cd = " CH"
    Call IssueInsert
    
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


Private Sub IssueDeleteCustData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueDeleteCustData"

    strSql = "delete from customer "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    
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
Private Sub IssueDeleteCustName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueDeleteCustName"
    
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_cust_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "customer", "S", "=")
    
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

Private Sub IssueDeleteCustAddr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueDeleteCustAddr"
    
    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_cust_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "customer", "S", "=")
    
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

Private Sub IssueDeleteContacts()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueDeleteContacts"
    
    strSql = "delete from contact "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    
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


Private Sub validate(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    Dim varSales_rep_assoc_id As Variant
    Dim varCust_rep_assoc_id As Variant
    strRoutine_nm = "customer.cls.validate"
    
    If sFunction_cd = "delete" Then
                
        strSql = "SELECT cust_id      "
        strSql = strSql & "FROM sales_order "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            objRecordset.Close
            Set objRecordset = Nothing
            Exit Sub
        End If
        
        If objRecordset.EOF = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "3200"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 3300
            Exit Sub
        End If
        objRecordset.Close
        Set objRecordset = Nothing
        
        'If o_old_customer.p_status_cd = "AC" Then
        '    o_error.p_type_cd = "E"
        '    o_error.p_err_cd = "3400"
        '    o_error.p_routine_nm = strRoutine_nm
        '    o_error.p_message_id = 1082
        '    Exit Sub
        'End If
        
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If fValuePresent(p_cust_nbr) = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "3600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1075
            Exit Sub
        End If
        Call ValidateCust_nbr
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
    End If
    
    If fValuePresent(p_class_cd) = True And _
       fGetDesc("class_cd", p_class_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "3800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1090
        Exit Sub
    End If
    
    If fValuePresent(p_credit_limit_am) = True And _
       IsNumeric(p_credit_limit_am) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1095
        Exit Sub
    End If
    
    If fValuePresent(p_currency_cd) = True And _
       fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1100
        Exit Sub
    End If
    
    If fValuePresent(p_fob_terms_cd) = True And _
       fGetDesc("FOB_terms_cd", p_fob_terms_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1105
        Exit Sub
    End If
    
    If fValuePresent(p_freight_cd) = True And _
       fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1110
        Exit Sub
    End If
    
    If fValuePresent(p_payment_terms_cd) = True And _
       fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "4800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1115
        Exit Sub
    End If
    
    If fValuePresent(p_sales_order_limit_am) = True And _
       IsNumeric(p_sales_order_limit_am) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "5000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1120
        Exit Sub
    End If
    
    If fValuePresent(p_sales_tax_pct) = True And _
       (IsNumeric(p_sales_tax_pct) = False Or _
       p_sales_tax_pct < 0 Or _
       p_sales_tax_pct > 1) Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "5200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1030
        Exit Sub
    End If

    If fValuePresent(p_sales_terr_cd) = True And _
       fGetDesc("sales_terr_cd", p_sales_terr_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "5400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1125
        Exit Sub
    End If
    
    If fValuePresent(p_ship_complete_var_pct) = True And _
       (IsNumeric(p_ship_complete_var_pct) = False Or _
       p_ship_complete_var_pct < 0 Or _
       p_ship_complete_var_pct > 1) Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "5600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1130
        Exit Sub
    End If

    If fValuePresent(p_sic_cd) = True And _
       IsNumeric(p_sic_cd) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "5800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1135
        Exit Sub
    End If
    
    If fValuePresent(p_source_cd) = True And _
       fGetDesc("cust_source_cd", p_source_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "6000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1140
        Exit Sub
    End If
    
    If sFunction_cd = "delete" Or sFunction_cd = "add" Or (sFunction_cd = "change" And p_status_cd = Empty) Then
    Else
        If IsNull(p_status_cd) = True Or _
           fGetDesc("cust_status_cd", p_status_cd, o_error, o_trans_data) = Empty Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "6200"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1145
            Exit Sub
        End If
    End If
        
    If sFunction_cd = "delete" Or (sFunction_cd = "change" And p_stnd_tx_on_invoice_fl = Empty) Then
    Else
        If IsNull(p_stnd_tx_on_invoice_fl) = True Or _
           Not (p_stnd_tx_on_invoice_fl = "Y" Or p_stnd_tx_on_invoice_fl = "N") Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "6400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1150
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "delete" Or (sFunction_cd = "change" And p_stnd_tx_on_pack_slip_fl = Empty) Then
    Else
        If IsNull(p_stnd_tx_on_pack_slip_fl) = True Or _
           Not (p_stnd_tx_on_pack_slip_fl = "Y" Or p_stnd_tx_on_pack_slip_fl = "N") Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "6600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1155
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "delete" Or (sFunction_cd = "change" And p_tax_exempt_fl = Empty) Then
    Else
        If IsNull(p_tax_exempt_fl) = True Or _
           Not (p_tax_exempt_fl = "Y" Or p_tax_exempt_fl = "N") Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "6800"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1160
            Exit Sub
        End If
    End If
    
    If fValuePresent(p_transportation_cd) = True And _
       fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "7000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1165
        Exit Sub
    End If
    
    If sFunction_cd = "delete" Or (sFunction_cd = "change" And p_consignment_fl = Empty) Then
    Else
        If IsNull(p_consignment_fl) = True Or _
           Not (p_consignment_fl = "Y" Or p_consignment_fl = "N") Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "7200"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1170
            Exit Sub
        End If
    End If
    
    If fValuePresent(p_cons_inv_loc_id) = True Then
        Dim objLocation As Object
        Set objLocation = CreateObject("converge_inv.inventory_loc")
        Set objLocation.o_error = o_error
        Set objLocation.o_trans_data = o_trans_data
        objLocation.p_inv_loc_id = p_cons_inv_loc_id
        Call objLocation.getInventoryLoc
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
           Exit Sub
        End If
    
        If objLocation.p_inv_loc_id = Null Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "7400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1175
            Exit Sub
        End If
    Else
        If p_consignment_fl = "Y" Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "7600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1180
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If fValuePresent(p_sales_rep_assoc_nbr) Then
            Call ValidateAssoc_nbr(p_sales_rep_assoc_nbr, varSales_rep_assoc_id, o_error, o_trans_data)
            If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "7801"
                o_error.p_routine_nm = strRoutine_nm & p_sales_rep_assoc_nbr
                o_error.p_message_id = 1185
                Exit Sub
            Else
                p_sales_rep_assoc_id = varSales_rep_assoc_id
            End If
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "7802"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1185
            Exit Sub
        End If
        
        ' JRP4
        If fValuePresent(p_cust_rep_assoc_nbr) Then
            Call ValidateAssoc_nbr(p_cust_rep_assoc_nbr, varCust_rep_assoc_id, o_error, o_trans_data)
            If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "8800"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1187
                Exit Sub
            Else
                p_cust_rep_assoc_id = varCust_rep_assoc_id
            End If
        End If
    
    End If
        
    If sFunction_cd = "change" Or sFunction_cd = "add" Then
        If Len(p_stnd_tx) > 2999 Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "8000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1186
            Exit Sub
        End If
        Call ValidateCust_grp_nbr(sFunction_cd)
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ValidateCust_nbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.ValidateCust_nbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If IsNull(p_cust_nbr) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "8400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1075
        Exit Sub
    End If
    
    If InStr(1, p_cust_nbr, "&", vbBinaryCompare) > 0 Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "8410"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1076
        o_error.p_desc = "The '&' character is not valid in customer number"
        Exit Sub
    End If
    
    If IsNull(p_cust_id) = True Then
        strSql = "SELECT cust_id      "
        strSql = strSql & "FROM customer "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_nbr", p_cust_nbr, "S", "=")
    Else
        strSql = "SELECT cust_id      "
        strSql = strSql & "FROM customer "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_nbr", p_cust_nbr, "S", "=")
        strSql = strSql & "and" & fWhereVariable("cust_id", p_cust_id, "N", "<>")
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
    Else
        objRecordset.Close
        Set objRecordset = Nothing
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "8600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1190
        Exit Sub
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "8800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ValidateCust_grp_nbr(ByVal v_sFunction_cd As String)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "customer.cls.ValidateCust_Grp_nbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If IsNull(o_cust_group.p_cust_grp_nbr) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "9005"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = "10001"
        o_error.p_desc = "Customer Group is required"
        Exit Sub
    End If
    
    If (v_sFunction_cd = "add") Then
        If fValuePresent(o_cust_group.p_cust_grp_nbr) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "9005"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = "10001"
            o_error.p_desc = "Customer Group is required"
            Exit Sub
        End If
    End If
    
    If (v_sFunction_cd = "change") Then
        If fValuePresent(o_cust_group.p_cust_grp_nbr) = False Then
            p_cust_grp_id = Empty
            Exit Sub
        End If
    End If
    
    Set o_cust_group.o_trans_data = o_trans_data
    Set o_cust_group.o_error = o_error
    o_cust_group.p_cust_grp_id = Null
    Call o_cust_group.GetCustGrp
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    If IsNull(o_cust_group.p_cust_grp_id) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "9000"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1061
        Exit Sub
    End If
    
    p_cust_grp_id = o_cust_group.p_cust_grp_id
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "9100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueInsert"
    
    strSql = "insert into customer ( "
    strSql = strSql & "  carrier_nm "
    strSql = strSql & ", class_cd "
    strSql = strSql & ", consignment_fl "
    strSql = strSql & ", cons_inv_loc_id "
    strSql = strSql & ", credit_limit_am "
    strSql = strSql & ", currency_cd "
    strSql = strSql & ", cust_grp_id "
    strSql = strSql & ", cust_id "
    strSql = strSql & ", cust_nbr "
    strSql = strSql & ", email_addr "
    strSql = strSql & ", fax_nbr "
    strSql = strSql & ", FOB_terms_cd "
    strSql = strSql & ", freight_cd "
    strSql = strSql & ", payment_terms_cd "
    strSql = strSql & ", phone_nbr "
    strSql = strSql & ", registration_dt "
    strSql = strSql & ", sales_order_limit_am "
    strSql = strSql & ", sales_rep_assoc_id "
    ' JRP4
    strSql = strSql & ", cust_rep_assoc_id "
    strSql = strSql & ", sales_tax_pct "
    strSql = strSql & ", sales_terr_cd "
    strSql = strSql & ", ship_complete_var_pct "
    strSql = strSql & ", sic_cd "
    strSql = strSql & ", source_cd "
    strSql = strSql & ", status_cd "
    strSql = strSql & ", stnd_tx "
    strSql = strSql & ", stnd_tx_on_invoice_fl "
    strSql = strSql & ", stnd_tx_on_pack_slip_fl "
    strSql = strSql & ", tax_exempt_fl "
    strSql = strSql & ", tax_id "
    strSql = strSql & ", timco_supp_cd "
    strSql = strSql & ", ship_acct_nbr "
    strSql = strSql & ", web_site_addr "
    strSql = strSql & ", transportation_cd "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_carrier_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_class_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_consignment_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_cons_inv_loc_id, "S")
    strSql = strSql & "," & fInsertVariable(p_credit_limit_am, "N")
    strSql = strSql & "," & fInsertVariable(p_currency_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_cust_grp_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_email_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_fax_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_fob_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_freight_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_payment_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_phone_nbr, "S")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(p_sales_order_limit_am, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_rep_assoc_id, "N")
    
    ' JRP4
    strSql = strSql & "," & fInsertVariable(p_cust_rep_assoc_id, "N")
    
    strSql = strSql & "," & fInsertVariable(p_sales_tax_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_sales_terr_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_ship_complete_var_pct, "N")
    strSql = strSql & "," & fInsertVariable(p_sic_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_source_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_stnd_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_stnd_tx_on_invoice_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_stnd_tx_on_pack_slip_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_tax_exempt_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_tax_id, "S")
    strSql = strSql & "," & fInsertVariable(p_timco_supp_cd, "S")
    strSql = strSql & "," & fInsertVariable(Me.p_ship_acct_nbr, "S")
    strSql = strSql & "," & fInsertVariable(Me.p_web_site_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_transportation_cd, "S")
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
      .p_err_cd = "9200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateCustData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "customer.cls.IssueUpdateCustData"
    
    strSql = "update customer set "
    strSql = strSql & " " & fUpdateVariable("cust_nbr", p_cust_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("carrier_nm", p_carrier_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("class_cd", p_class_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("consignment_fl", p_consignment_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cons_inv_loc_id", p_cons_inv_loc_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("credit_limit_am", p_credit_limit_am, "N", "=")
    strSql = strSql & "," & fUpdateVariable("currency_cd", p_currency_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_grp_id", p_cust_grp_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("email_addr", p_email_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("fax_nbr", p_fax_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("FOB_terms_cd", p_fob_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("freight_cd", p_freight_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("phone_nbr", p_phone_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("sales_order_limit_am", p_sales_order_limit_am, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_rep_assoc_id", p_sales_rep_assoc_id, "N", "=")
    
    ' JRP4
    strSql = strSql & "," & fUpdateVariable("cust_rep_assoc_id", p_cust_rep_assoc_id, "N", "=")
    
    strSql = strSql & "," & fUpdateVariable("sales_tax_pct", p_sales_tax_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sales_terr_cd", p_sales_terr_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ship_complete_var_pct", p_ship_complete_var_pct, "N", "=")
    strSql = strSql & "," & fUpdateVariable("sic_cd", p_sic_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("source_cd", p_source_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("stnd_tx", p_stnd_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("stnd_tx_on_invoice_fl", p_stnd_tx_on_invoice_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("stnd_tx_on_pack_slip_fl", p_stnd_tx_on_pack_slip_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("tax_exempt_fl", p_tax_exempt_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("tax_id", p_tax_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("timco_supp_cd", p_timco_supp_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ship_acct_nbr", Me.p_ship_acct_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("web_site_addr", Me.p_web_site_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("transportation_cd", p_transportation_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "9300"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub SelectCustomer()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "customer.cls.SelectCustomer"
    
    If IsNull(p_cust_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM customer "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_nbr", p_cust_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM customer "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_cust_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_carrier_nm = objRecordset("carrier_nm")
        p_class_cd = objRecordset("class_cd")
        p_consignment_fl = objRecordset("consignment_fl")
        p_cons_inv_loc_id = objRecordset("cons_inv_loc_id")
        If IsNull(objRecordset("credit_limit_am")) = True Then
            p_credit_limit_am = Null
        Else
            p_credit_limit_am = CDbl(objRecordset("credit_limit_am"))
        End If
        p_currency_cd = objRecordset("currency_cd")
        p_cust_grp_id = objRecordset("cust_grp_id")
        p_cust_id = objRecordset("Cust_id")
        p_cust_nbr = objRecordset("Cust_nbr")
        p_email_addr = objRecordset("email_addr")
        p_fax_nbr = objRecordset("fax_nbr")
        p_fob_terms_cd = objRecordset("FOB_terms_cd")
        p_freight_cd = objRecordset("freight_cd")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_phone_nbr = objRecordset("phone_nbr")
        p_registration_dt = objRecordset("registration_dt")
        p_sales_order_limit_am = objRecordset("sales_order_limit_am")
        p_sales_rep_assoc_id = objRecordset("sales_rep_assoc_id")
        ' JRP4
        p_cust_rep_assoc_id = objRecordset("cust_rep_assoc_id")
        If IsNull(objRecordset("sales_tax_pct")) = True Then
            p_sales_tax_pct = Null
        Else
            p_sales_tax_pct = CDbl(objRecordset("sales_tax_pct"))
        End If
        p_sales_terr_cd = objRecordset("sales_terr_cd")
        p_sic_cd = objRecordset("sic_cd")
        p_sic_desc = fGetDesc("sic_cd", objRecordset("sic_cd"), o_error, o_trans_data)
        p_source_cd = objRecordset("source_cd")
        p_status_cd = objRecordset("status_cd")
        p_stnd_tx = objRecordset("stnd_tx")
        p_stnd_tx_on_invoice_fl = objRecordset("stnd_tx_on_invoice_fl")
        p_stnd_tx_on_pack_slip_fl = objRecordset("stnd_tx_on_pack_slip_fl")
        p_tax_exempt_fl = objRecordset("tax_exempt_fl")
        p_tax_id = objRecordset("tax_id")
        p_timco_supp_cd = objRecordset("timco_supp_cd")
        Me.p_web_site_addr = objRecordset("web_site_addr")
        Me.p_ship_acct_nbr = objRecordset("ship_acct_nbr")
        p_transportation_cd = objRecordset("transportation_cd")
        If IsNull(objRecordset("ship_complete_var_pct")) Then
            p_ship_complete_var_pct = Null
        Else
            p_ship_complete_var_pct = CDbl(objRecordset("ship_complete_var_pct"))
        End If
    
        objRecordset.Close
        Set objRecordset = Nothing
        
        p_class_desc = fGetDesc("class_cd", p_class_cd, o_error, o_trans_data)
        p_currency_desc = fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data)
        p_fob_terms_desc = fGetDesc("FOB_terms_cd", p_fob_terms_cd, o_error, o_trans_data)
        p_freight_desc = fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data)
        p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data)
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_sales_rep_assoc_nbr = fGetAssocNbr(p_sales_rep_assoc_id, o_error, o_trans_data)
        ' JRP4
        p_cust_rep_assoc_nbr = fGetAssocNbr(p_cust_rep_assoc_id, o_error, o_trans_data)
        
        p_sales_terr_desc = fGetDesc("sales_terr_cd", p_sales_terr_cd, o_error, o_trans_data)
        p_source_desc = fGetDesc("cust_source_cd", p_source_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("cust_status_cd", p_status_cd, o_error, o_trans_data)
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "9400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub GetCustomers()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim strCustNbr As String
    Dim lngCount As Long
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "Customer.cls.getCustomers"
        
    strSql = "SELECT  " & _
                " a.cust_id, " & _
                " a.cust_nbr, " & _
                " b.nm, " & _
                " c.code_desc, " & _
                " d.city_nm, " & _
                " a.source_cd " & _
           "  FROM customer a, name b, code_desc c, address d" & _
            " WHERE (" & fWhereVariable("c.field_nm", "cust_status_cd", "S", "=") & _
              " AND c.code = a.status_cd)" & _
              " AND (" & fWhereVariable("b.obj_nm", "customer", "S", "=") & _
              " AND  b.nm_type_cd = 'STND'" & _
              " AND  a.cust_id = b.obj_id)" & _
              " AND (" & fWhereVariable("d.obj_nm", "customer", "S", "=") & _
              " AND  a.cust_id = d.obj_id " & _
              " AND  d.addr_type_cd = 'STND' )"
                
    If IsNull(p_cust_nbr) = True Or p_cust_nbr = "" Then
    Else
        lngCount = InStr(1, p_cust_nbr, "*", vbBinaryCompare)
        If lngCount > 0 Then
            strCustNbr = Left(p_cust_nbr, lngCount - 1)
            strSql = strSql & " AND a.cust_nbr LIKE '" & strCustNbr & "%'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.cust_nbr", p_cust_nbr, "S", "=")
        End If
    End If
    
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
    
    strSql = strSql & " ORDER BY a.cust_nbr "
    
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
      .p_err_cd = "9500"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub GetCustGrpCustomers()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "Customer.cls.GetCustGrpCustomers"
        
    strSql = "SELECT  " & _
                " a.cust_id, " & _
                " a.cust_nbr, " & _
                " b.nm, " & _
                " c.code_desc, " & _
                " d.city_nm, " & _
                " a.source_cd " & _
            "  FROM customer a, name b, code_desc c, address d" & _
            " WHERE (" & fWhereVariable("c.field_nm", "cust_status_cd", "S", "=") & _
              " AND c.code = a.status_cd)" & _
              " AND (" & fWhereVariable("b.obj_nm", "customer", "S", "=") & _
              " AND  b.nm_type_cd = 'STND'" & _
              " AND  a.cust_id = b.obj_id)" & _
              " AND (" & fWhereVariable("d.obj_nm", "customer", "S", "=") & _
              " AND  a.cust_id = d.obj_id " & _
              " AND  d.addr_type_cd = 'STND' )" & _
              " AND " & fWhereVariable("a.cust_grp_id", p_cust_grp_id, "N", "=")

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    strSql = strSql & " ORDER BY a.cust_nbr "


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
      .p_err_cd = "9600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub InsertQBTrans( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByVal v_varCustId As Variant _
        )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = "Converge_Sales" & "." & "customer" & "." & "InsertQBTrans"
    
    Dim strSqlTx As String
    Dim dblUpdateQty As Double
    
    strSqlTx = "Insert Into quickbooks_trans ( "
    strSqlTx = strSqlTx & " trans_type_cd "
    strSqlTx = strSqlTx & ", cust_id "
    strSqlTx = strSqlTx & ", insert_ts "
    strSqlTx = strSqlTx & " ) "
    strSqlTx = strSqlTx & "values ("
    strSqlTx = strSqlTx & "  1 "
    strSqlTx = strSqlTx & "," & v_varCustId & " "
    strSqlTx = strSqlTx & ", GetDate()  "
    strSqlTx = strSqlTx & ") "
   
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Err.clear
End Sub


