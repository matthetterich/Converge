VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_supplier_id As Variant
Private mvarp_supplier_nbr As Variant
Private mvarp_comm_meth_cd As Variant
Private mvarp_comm_meth_desc As Variant
Private mvarp_currency_cd As Variant
Private mvarp_currency_desc As Variant
Private mvarp_dun_brad_nbr As Variant
Private mvarp_email_addr As Variant
Private mvarp_fax_nbr As Variant
Private mvarp_fob_terms_cd As Variant
Private mvarp_fob_terms_desc As Variant
Private mvarp_freight_cd As Variant
Private mvarp_freight_desc As Variant
Private mvarp_reg_dt As Variant
Private mvarp_sic_cd As Variant
Private mvarp_sic_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_tax_id As Variant
Private mvarp_tax_exempt_fl As Variant
Private mvarp_transportation_cd As Variant
Private mvarp_transportation_desc As Variant
Private mvarp_type_cd As Variant
Private mvarp_type_desc As Variant
Private mvarp_phone_nbr As Variant
Private mvarp_order_text As Variant
Private mvarp_bypass_nm_addr_fl As Variant
Private mvarp_registration_dt As Variant
Private mvarp_payment_terms_cd As Variant
Private mvarp_payment_terms_desc As Variant
Private mvarp_web_site_addr As Variant
Private mvaro_stnd_name As Object
Private mvaro_stnd_address As Object
Private mvaro_oth_name As Object
Private mvaro_oth_address As Object
Private mvaro_record_set As ADODB.Recordset
'

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

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Set o_oth_address(ByVal vData As Object)
    Set mvaro_oth_address = vData
End Property
Public Property Get o_oth_address() As Object
    Set o_oth_address = mvaro_oth_address
End Property

Public Property Set o_oth_name(ByVal vData As Object)
    Set mvaro_oth_name = vData
End Property
Public Property Get o_oth_name() As Object
    Set o_oth_name = mvaro_oth_name
End Property
Public Property Set o_stnd_address(ByVal vData As Object)
    Set mvaro_stnd_address = vData
End Property
Public Property Get o_stnd_address() As Object
    Set o_stnd_address = mvaro_stnd_address
End Property
Public Property Set o_stnd_name(ByVal vData As Object)
    Set mvaro_stnd_name = vData
End Property
Public Property Get o_stnd_name() As Object
    Set o_stnd_name = mvaro_stnd_name
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
Public Property Let p_bypass_nm_addr_fl(ByVal vData As Variant)
    mvarp_bypass_nm_addr_fl = vData
End Property
Public Property Get p_bypass_nm_addr_fl() As Variant
        p_bypass_nm_addr_fl = mvarp_bypass_nm_addr_fl
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
Public Property Let p_tax_exempt_fl(ByVal vData As Variant)
    mvarp_tax_exempt_fl = vData
End Property
Public Property Get p_tax_exempt_fl() As Variant
        p_tax_exempt_fl = mvarp_tax_exempt_fl
End Property
Public Property Let p_tax_id(ByVal vData As Variant)
    mvarp_tax_id = vData
End Property
Public Property Get p_tax_id() As Variant
        p_tax_id = mvarp_tax_id
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
Public Property Let p_sic_desc(ByVal vData As Variant)
    mvarp_sic_desc = vData
End Property
Public Property Get p_sic_desc() As Variant
        p_sic_desc = mvarp_sic_desc
End Property
Public Property Let p_sic_cd(ByVal vData As Variant)
    mvarp_sic_cd = vData
End Property
Public Property Get p_sic_cd() As Variant
        p_sic_cd = mvarp_sic_cd
End Property
Public Property Let p_reg_dt(ByVal vData As Variant)
    mvarp_reg_dt = vData
End Property
Public Property Get p_reg_dt() As Variant
        p_reg_dt = mvarp_reg_dt
End Property
Public Property Let p_freight_desc(ByVal vData As Variant)
    mvarp_freight_desc = vData
End Property
Public Property Get p_freight_desc() As Variant
        p_freight_desc = mvarp_freight_desc
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

Public Property Let p_web_site_addr(ByVal vData As Variant)
    mvarp_web_site_addr = vData
End Property
Public Property Get p_web_site_addr() As Variant
        p_web_site_addr = mvarp_web_site_addr
End Property


Public Property Let p_dun_brad_nbr(ByVal vData As Variant)
    mvarp_dun_brad_nbr = vData
End Property
Public Property Get p_dun_brad_nbr() As Variant
        p_dun_brad_nbr = mvarp_dun_brad_nbr
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
Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
End Property
Public Property Let p_order_text(ByVal vData As Variant)
    mvarp_order_text = vData
End Property
Public Property Get p_order_text() As Variant
        p_order_text = mvarp_order_text
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
Public Sub GetSupplier()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.GetSupplier"
    
    Set o_stnd_name = CreateObject("converge.name")
    Set o_stnd_address = CreateObject("converge.address")
    Set o_oth_name = CreateObject("converge.name")
    Set o_oth_address = CreateObject("converge.address")
    
    Call SelectSupplier
    
    If IsNull(p_supplier_id) = True Then
      o_stnd_name.p_nm_id = Null
      o_stnd_address.p_addr_id = Null
      o_oth_name.p_nm_id = Null
      o_oth_address.p_addr_id = Null
    Else
        Set o_stnd_name.o_trans_data = o_trans_data
        Set o_stnd_name.o_error = o_error
        o_stnd_name.p_nm_id = Null
        o_stnd_name.p_obj_nm = "supplier"
        o_stnd_name.p_obj_id = p_supplier_id
        o_stnd_name.p_nm_type_cd = "STND"
        o_stnd_name.p_def_nm_fl = "Y"
        Call o_stnd_name.getDefName
        
        Set o_oth_name.o_trans_data = o_trans_data
        Set o_oth_name.o_error = o_error
        o_oth_name.p_nm_id = Null
        o_oth_name.p_obj_nm = "supplier"
        o_oth_name.p_obj_id = p_supplier_id
        o_oth_name.p_nm_type_cd = "OTH"
        o_oth_name.p_def_nm_fl = "Y"
        Call o_oth_name.getDefName
      
        Set o_stnd_address.o_trans_data = o_trans_data
        Set o_stnd_address.o_error = o_error
        o_stnd_address.p_addr_id = Null
        o_stnd_address.p_obj_nm = "supplier"
        o_stnd_address.p_obj_id = p_supplier_id
        o_stnd_address.p_addr_type_cd = "STND"
        o_stnd_address.p_def_addr_fl = "Y"
        Call o_stnd_address.getDefAddress
        
        Set o_oth_address.o_trans_data = o_trans_data
        Set o_oth_address.o_error = o_error
        o_oth_address.p_addr_id = Null
        o_oth_address.p_obj_nm = "supplier"
        o_oth_address.p_obj_id = p_supplier_id
        o_oth_address.p_addr_type_cd = "OTH"
        o_oth_address.p_def_addr_fl = "Y"
        Call o_oth_address.getDefAddress
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


Public Sub AddSupplier()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.AddSupplier"
    
    Call ValidateSupplierData("add")
    Call AddSupplierData
    Call AddSupplierName(o_stnd_name)
    Call AddSupplierAddr(o_stnd_address)
    
    If o_oth_name.p_nm = Empty Then
    Else
        Call AddSupplierName(o_oth_name)
    End If
    
    If o_oth_address.p_Street_addr = Empty And o_oth_address.p_City_nm = Empty And o_oth_address.p_state_cd = Empty And o_oth_address.p_zip_cd = Empty And o_oth_address.p_country_cd = Empty Then
    Else
        Call AddSupplierAddr(o_oth_address)
    End If
    
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_supplier_id)
    
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

Public Sub ChangeSupplier()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.ChangeSupplier"
    
    Call ValidateSupplierData("change")
    Call IssueUpdateSupplierData
    
    If p_bypass_nm_addr_fl = "Y" Then
    Else
        Call ChangeSupplierName(o_stnd_name)
        Call ChangeSupplierAddr(o_stnd_address)
        
        If IsNull(o_oth_name.p_nm_id) = True Or o_oth_name.p_nm_id = Empty Then
            If IsNull(o_oth_name.p_nm) = True Or o_oth_name.p_nm = Empty Then
            Else
                Call AddSupplierName(o_oth_name)
            End If
        Else
            If IsNull(o_oth_name.p_nm) = True Then
                Call DeleteSupplierName(o_oth_name)
            Else
                Call ChangeSupplierName(o_oth_name)
            End If
        End If
        
        If IsNull(o_oth_address.p_addr_id) = True Or o_oth_address.p_addr_id = Empty Then
            If (IsNull(o_oth_address.p_Street_addr) = True Or o_oth_address.p_Street_addr = Empty) And _
               (IsNull(o_oth_address.p_City_nm) = True Or o_oth_address.p_City_nm = Empty) And _
               (IsNull(o_oth_address.p_zip_cd) = True Or o_oth_address.p_zip_cd = Empty) Then
            Else
                Call AddSupplierAddr(o_oth_address)
            End If
        Else
            If IsNull(o_oth_address.p_Street_addr) = True And _
               IsNull(o_oth_address.p_City_nm) = True And _
               IsNull(o_oth_address.p_zip_cd) = True Then
                Call DeleteSupplierAddr(o_oth_address)
            Else
                Call ChangeSupplierAddr(o_oth_address)
            End If
        End If
        
    End If
    
    strRoutine_nm = "supplier.cls.ChangeSupplier4"
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_supplier_id)
    
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

Public Sub DeleteSupplier()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.DeleteSupplier"
    
    Call IssueDeleteSupplierData
    Call IssueDeleteSupplierName
    Call IssueDeleteSupplierAddr
    Call IssueDeleteContacts
    
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


Private Sub AddSupplierName(objName)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.AddSupplierName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_id = p_supplier_id
    objName.p_obj_nm = "supplier"
    Call objName.AddName
    
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

Private Sub ChangeSupplierName(objName)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.ChangeSupplierName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_id = p_supplier_id
    objName.p_obj_nm = "supplier"
    Call objName.ChangeName
    
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


Private Sub DeleteSupplierName(objName)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.DeleteSupplierName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    objName.p_obj_id = p_supplier_id
    objName.p_obj_nm = "supplier"
    Call objName.DeleteName
    
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


Private Sub AddSupplierAddr(objAddress)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.AddSupplierAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_id = p_supplier_id
    objAddress.p_obj_nm = "supplier"
    Call objAddress.AddAddress
    
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


Private Sub ChangeSupplierAddr(objAddress)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.ChangeSupplierAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_id = p_supplier_id
    objAddress.p_obj_nm = "supplier"
    Call objAddress.ChangeAddress
    
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


Private Sub DeleteSupplierAddr(objAddress)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.DeleteSupplierAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    objAddress.p_obj_id = p_supplier_id
    objAddress.p_obj_nm = "supplier"
    Call objAddress.DeleteAddress
    
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


Private Sub AddSupplierData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.AddSupplierData"
    
    p_supplier_id = fGetId("supplier_id", o_error, o_trans_data)
    Call IssueInsertSupplierData
    
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


Private Sub IssueDeleteSupplierData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier.cls.IssueDeleteSupplierData"
    
    strSql = "delete from supplier "
    strSql = strSql & "where " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
        
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
    
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

Private Sub IssueDeleteSupplierName()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier.cls.IssueDeleteSupplierName"
      
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_supplier_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "supplier", "S", "=")
      
     Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)

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

Private Sub IssueDeleteSupplierAddr()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double

    strRoutine_nm = "supplier.cls.IssueDeleteSupplierAddr"
    
    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_supplier_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "supplier", "S", "=")
      
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
    
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

Private Sub IssueDeleteContacts()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
     
    strRoutine_nm = "supplier.cls.IssueDeleteContacts"
      
    strSql = "delete from contact "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
      
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
   
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


Private Sub ValidateSupplierData(sFunction_cd)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "supplier.cls.ValidateSupplierData"

    If sFunction_cd = "add" Then
        If fValuePresent(p_supplier_nbr) = False Then
          o_error.p_type_cd = "F"
          o_error.p_err_cd = "1100"
          o_error.p_routine_nm = strRoutine_nm
          o_error.p_message_id = 2105
        End If
    End If
    
    If fValuePresent(p_comm_meth_cd) = True And _
       fGetDesc("comm_meth_cd", p_comm_meth_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2055
        Exit Sub
    End If
    
    If fValuePresent(p_currency_cd) = True And _
       fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1300"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2015
        Exit Sub
    End If
    
    If fValuePresent(p_fob_terms_cd) = True And _
       fGetDesc("FOB_terms_cd", p_fob_terms_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1105
        Exit Sub
    End If
    
    If fValuePresent(p_freight_cd) = True And _
       fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1500"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1110
        Exit Sub
    End If
    
    If sFunction_cd = "change" And p_status_cd = Empty Then
    Else
        If IsNull(p_status_cd) = True Or _
           fGetDesc("status_cd", p_status_cd, o_error, o_trans_data) = Empty Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1700"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2010
            Exit Sub
        End If
    End If
        
    If fValuePresent(p_transportation_cd) = True And _
       fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2070
        Exit Sub
    End If
    
    If sFunction_cd = "change" And p_type_cd = Empty Then
    Else
        If IsNull(p_type_cd) = True Or _
           fGetDesc("supplier_type_cd", p_type_cd, o_error, o_trans_data) = Empty Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1850"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2149
            Exit Sub
        End If
    End If
    
    Call ValidateSupplier_nbr
    
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


Private Sub ValidateSupplier_nbr()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim objRecordset As ADODB.Recordset
    Dim objRecordSet1 As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "supplier.cls.ValidateSupplier_nbr"
    
    If IsNull(p_supplier_nbr) = True Then
      o_error.p_type_cd = "F"
      o_error.p_err_cd = "1900"
      o_error.p_routine_nm = strRoutine_nm
      o_error.p_message_id = 2080
        Exit Sub
    End If
    
    If InStr(1, p_supplier_nbr, "&", vbBinaryCompare) > 0 Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1905"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1077
        o_error.p_desc = "The '&' character is not valid in customer number"
        Exit Sub
    End If
    
    If IsNull(p_supplier_id) = True Or p_supplier_id = Empty Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM supplier "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("supplier_nbr", p_supplier_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM supplier "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("supplier_nbr", p_supplier_nbr, "S", "=")
        strSql = strSql & "and" & fWhereVariable("supplier_id", p_supplier_id, "N", "<>")
    End If
    
   Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    
    If objRecordset.EOF = True Then
    Else
      o_error.p_type_cd = "F"
      o_error.p_err_cd = "2100"
      o_error.p_routine_nm = strRoutine_nm
      o_error.p_message_id = 2100
    End If
    
    strSql = "select count(*) as intCount from customer Where cust_nbr = '" & p_supplier_nbr & "'"
    Call o_trans_data.OpenRecordset(objRecordSet1, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet1.Close
        Set objRecordSet1 = Nothing
        Exit Sub
    End If
    
    If objRecordSet1.Fields("intCount") > 0 Then
         With o_error
          .p_type_cd = "F"
          .p_err_cd = "0400"
          .p_nbr = 400
          .p_desc = " The Supplier Number must not be the same as a customer number"
          .p_routine_nm = strRoutine_nm
          .p_message_id = 3623
        End With
    End If
    
    
    objRecordSet1.Close
    Set objRecordSet1 = Nothing
    objRecordset.Close
    Set objRecordset = Nothing
    
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

Private Sub IssueInsertSupplierData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier.cls.IssueInsertSupplier"
      
    strSql = "insert into supplier ( "
    strSql = strSql & "  supplier_id "
    strSql = strSql & ", supplier_nbr "
    strSql = strSql & ", comm_meth_cd "
    strSql = strSql & ", currency_cd "
    strSql = strSql & ", dun_brad_nbr "
    strSql = strSql & ", email_addr "
    strSql = strSql & ", web_site_addr "
    strSql = strSql & ", fax_nbr "
    strSql = strSql & ", fob_terms_cd "
    strSql = strSql & ", freight_cd "
    strSql = strSql & ", phone_nbr "
    strSql = strSql & ", registration_dt "
    strSql = strSql & ", sic_cd "
    strSql = strSql & ", status_cd "
    strSql = strSql & ", tax_id "
    strSql = strSql & ", tax_exempt_fl "
    strSql = strSql & ", transportation_cd "
    strSql = strSql & ", type_cd "
    strSql = strSql & ", order_tx "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(p_supplier_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_comm_meth_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_currency_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_dun_brad_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_email_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_web_site_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_fax_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_fob_terms_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_freight_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_phone_nbr, "S")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(p_sic_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_tax_id, "S")
    strSql = strSql & "," & fInsertVariable(p_tax_exempt_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_transportation_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_order_text, "S")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
    
    
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


Private Sub IssueUpdateSupplierData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "supplier.cls.IssueUpdateSupplierData"
      
    strSql = "UPDATE supplier set "
    strSql = strSql & " " & fUpdateVariable("supplier_nbr", p_supplier_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("comm_meth_cd", p_comm_meth_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("currency_cd", p_currency_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("dun_brad_nbr", p_dun_brad_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("email_addr", p_email_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("web_site_addr", p_web_site_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("fax_nbr", p_fax_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("FOB_terms_cd", p_fob_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("freight_cd", p_freight_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("payment_terms_cd", p_payment_terms_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("phone_nbr", p_phone_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("SIC_cd", p_sic_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("tax_id", p_tax_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("tax_exempt_fl", p_tax_exempt_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("transportation_cd", p_transportation_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("type_cd", p_type_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("order_tx", p_order_text, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
      
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
 
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


Private Sub SelectSupplier()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    strRoutine_nm = "supplier.cls.SelectSupplier"


    
    If IsNull(p_supplier_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM supplier "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("supplier_nbr", p_supplier_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM supplier "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If objRecordset.EOF = True Then
        p_supplier_id = Null
    Else
        p_supplier_id = objRecordset("Supplier_id")
        p_supplier_nbr = objRecordset("Supplier_nbr")
        p_comm_meth_cd = objRecordset("comm_meth_cd")
        p_currency_cd = objRecordset("currency_cd")
        p_dun_brad_nbr = objRecordset("dun_brad_nbr")
        p_email_addr = objRecordset("email_addr")
        p_web_site_addr = objRecordset("web_site_addr")
        p_fax_nbr = objRecordset("fax_nbr")
        p_fob_terms_cd = objRecordset("FOB_terms_cd")
        p_freight_cd = objRecordset("freight_cd")
        p_payment_terms_cd = objRecordset("payment_terms_cd")
        p_phone_nbr = objRecordset("phone_nbr")
        p_registration_dt = objRecordset("registration_dt")
        p_sic_cd = objRecordset("SIC_cd")
        p_status_cd = objRecordset("status_cd")
        p_tax_exempt_fl = objRecordset("tax_exempt_fl")
        p_tax_id = objRecordset("tax_id")
        p_transportation_cd = objRecordset("transportation_cd")
        p_type_cd = objRecordset("type_cd")
        p_order_text = objRecordset("order_tx")
          
        objRecordset.Close
        Set objRecordset = Nothing
          
        p_payment_terms_desc = fGetDesc("payment_terms_cd", p_payment_terms_cd, o_error, o_trans_data)
        p_type_desc = fGetDesc("supplier_type_cd", p_type_cd, o_error, o_trans_data)
        p_comm_meth_desc = fGetDesc("comm_meth_cd", p_comm_meth_cd, o_error, o_trans_data)
        p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("status_cd", p_status_cd, o_error, o_trans_data)
        p_freight_desc = fGetDesc("freight_cd", p_freight_cd, o_error, o_trans_data)
        p_fob_terms_desc = fGetDesc("FOB_terms_cd", p_fob_terms_cd, o_error, o_trans_data)
        p_currency_desc = fGetDesc("currency_cd", p_currency_cd, o_error, o_trans_data)
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


Public Sub GetSuppliers()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim lngCount As Long
    Dim strSupplierNbr As String
    Dim strPOStatus As String
    Dim objRecordset As ADODB.Recordset
    strRoutine_nm = "supplier.cls.getSuppliers"
        
    strPOStatus = "status_cd"
    
    strSql = "SELECT  " & _
                " a.supplier_id, " & _
                " a.supplier_nbr, " & _
                " b.nm, " & _
                " c.code_desc, " & _
                " a.order_tx, " & _
                " a.type_cd " & _
            "  FROM supplier a, name b, code_desc c" & _
            " WHERE (" & fWhereVariable("c.field_nm", strPOStatus, "S", "=") & _
              " AND c.code = a.status_cd)" & _
              " AND (" & fWhereVariable("b.obj_nm", "supplier", "S", "=") & _
              " AND  a.supplier_id = b.obj_id " & _
              " AND  b.nm_type_cd = 'STND') "
              
    If IsNull(p_supplier_nbr) = True Or p_supplier_nbr = "" Then
    Else
        lngCount = InStr(1, p_supplier_nbr, "*", vbBinaryCompare)
        If lngCount > 0 Then
            strSupplierNbr = Left(p_supplier_nbr, lngCount - 1)
            strSql = strSql & " AND a.supplier_nbr LIKE '" & strSupplierNbr & "%'"
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.supplier_nbr", p_supplier_nbr, "S", "=")
        End If
    End If
    
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
        
    strSql = strSql & " ORDER BY a.supplier_nbr"
    
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


Public Sub InsertQBTrans( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByVal v_varSupplierId As Variant _
        )

    On Error GoTo errHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = "Converge_purchase" & "." & "supplier" & "." & "InsertQBTrans"
    
    Dim strSqlTx As String
    Dim dblUpdateQty As Double
    
    strSqlTx = "Insert Into quickbooks_trans ( "
    strSqlTx = strSqlTx & " trans_type_cd "
    strSqlTx = strSqlTx & ", supplier_id "
    strSqlTx = strSqlTx & ", insert_ts "
    strSqlTx = strSqlTx & " ) "
    strSqlTx = strSqlTx & "values ("
    strSqlTx = strSqlTx & "  2 "
    strSqlTx = strSqlTx & "," & v_varSupplierId & " "
    strSqlTx = strSqlTx & ", GetDate()  "
    strSqlTx = strSqlTx & ") "
   
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


