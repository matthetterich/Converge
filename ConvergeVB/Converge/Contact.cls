VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_objError As Object
Private m_objTrans_data As Object
Private m_objCustomer As Object
Private m_objSupplier As Object
Private m_varContact_id As Variant
Private m_varEmail_addr As Variant
Private m_varFax_nbr As Variant
Private m_varNm As Variant
Private m_varPhone_nbr As Variant
Private m_varTitle As Variant
Private m_varCust_id As Variant
Private m_varSupplier_id As Variant
Private m_objAddress As Object
Private m_varBypass_getcust_fl As Variant
Private m_varBypass_getsupplier_fl As Variant
Private mvaro_record_set As adodb.Recordset

' Address Object Property for Contacts
Public Property Set o_address(ByVal vData As Object)
    Set m_objAddress = vData
End Property
Public Property Get o_address() As Object
    Set o_address = m_objAddress
End Property

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property


Public Property Set Supplier(ByVal vData As Object)
    Set m_objSupplier = vData
End Property

Public Property Get Supplier() As Object
    Set Supplier = m_objSupplier
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property

Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property

Public Property Set Trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property

Public Property Get Trans_data() As Object
    Set Trans_data = m_objTrans_data
End Property

Public Property Set Customer(ByVal vData As Object)
    Set m_objCustomer = vData
End Property

Public Property Get Customer() As Object
    Set Customer = m_objCustomer
End Property

Public Property Set error(ByVal vData As Object)
    Set m_objError = vData
End Property

Public Property Get error() As Object
    Set error = m_objError
End Property

Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property

Public Property Get o_error() As Object
    Set o_error = m_objError
End Property

Public Property Let bypass_getsupplier_fl(ByVal vData As Variant)
    m_varBypass_getsupplier_fl = vData
End Property
Public Property Get bypass_getsupplier_fl() As Variant
    bypass_getsupplier_fl = m_varBypass_getsupplier_fl
End Property

Public Property Let bypass_getcust_fl(ByVal vData As Variant)
    m_varBypass_getcust_fl = vData
End Property
Public Property Get bypass_getcust_fl() As Variant
    bypass_getcust_fl = m_varBypass_getcust_fl
End Property

Public Property Let supplier_id(ByVal vData As Variant)
    m_varSupplier_id = vData
End Property

Public Property Get supplier_id() As Variant
    supplier_id = m_varSupplier_id
End Property

Public Property Let cust_id(ByVal vData As Variant)
    m_varCust_id = vData
End Property

Public Property Get cust_id() As Variant
    cust_id = m_varCust_id
End Property

Public Property Let title(ByVal vData As Variant)
    m_varTitle = vData
End Property

Public Property Get title() As Variant
    title = m_varTitle
End Property

Public Property Let phone_nbr(ByVal vData As Variant)
    m_varPhone_nbr = vData
End Property

Public Property Get phone_nbr() As Variant
    phone_nbr = m_varPhone_nbr
End Property

Public Property Let nm(ByVal vData As Variant)
    m_varNm = vData
End Property

Public Property Get nm() As Variant
    nm = m_varNm
End Property

Public Property Let fax_nbr(ByVal vData As Variant)
    m_varFax_nbr = vData
End Property

Public Property Get fax_nbr() As Variant
    fax_nbr = m_varFax_nbr
End Property

Public Property Let email_addr(ByVal vData As Variant)
    m_varEmail_addr = vData
End Property
Public Property Get email_addr() As Variant
    email_addr = m_varEmail_addr
End Property

Public Property Let contact_id(ByVal vData As Variant)
    m_varContact_id = vData
End Property

Public Property Get contact_id() As Variant
    contact_id = m_varContact_id
End Property

Public Sub GetContact()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "contact.cls.GetContact"
    Set o_address = CreateObject("converge.address")

    
'    Set Customer = CreateObject("converge_sales.customer")
'    Set Supplier = CreateObject("converge_purchase.supplier")

    strSql = "SELECT * FROM contact "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("contact_id", contact_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
        
    If error.p_type_cd = "E" Or error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        contact_id = Null
        o_address.p_addr_id = Null
    Else
        contact_id = objRecordset("contact_id")
        email_addr = objRecordset("email_addr")
        fax_nbr = objRecordset("fax_nbr")
        nm = objRecordset("nm")
        phone_nbr = objRecordset("phone_nbr")
        title = objRecordset("title")
        
        cust_id = objRecordset("cust_id")
'        If bypass_getcust_fl = "Y" Or IsNull(cust_id) = True Then
'        Else
'            Set Customer.o_error = o_error
'            Set Customer.o_trans_data = o_trans_data
'            Customer.p_cust_id = cust_id
'            Call Customer.GetCustomer
'        End If
        
        supplier_id = objRecordset("supplier_id")
'        If bypass_getsupplier_fl = "Y" Or IsNull(supplier_id) = True Then
'        Else
'            Set Supplier = CreateObject("converge_purchase.supplier")
'            Set Supplier.o_error = o_error
'            Set Supplier.o_trans_data = o_trans_data
'            Supplier.p_supplier_id = supplier_id
'            Call Supplier.GetSupplier
'        End If

         Set o_address.o_trans_data = o_trans_data
         Set o_address.o_error = o_error
         o_address.p_addr_id = Null
         o_address.p_obj_nm = "contact"
         o_address.p_obj_id = contact_id
         o_address.p_addr_type_cd = "STND"
         o_address.p_def_addr_fl = "Y"
         Call o_address.getDefAddress

    End If
    
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


    
Public Sub AddContact()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.AddContact"
    
    Call Validate("add")
    
    Call IssueInsert
    
    Set o_address.o_error = o_error
    Set o_address.o_trans_data = o_trans_data
    o_address.p_obj_id = contact_id
    o_address.p_obj_nm = "contact"
    o_address.p_addr_type_cd = "STND"
    
    Call o_address.AddAddress
    
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


Public Sub ChangeContact()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.ChangeContact"
    
    Call Validate("change")
    Call IssueUpdate
    
    If Len(o_address.p_addr_id) = 0 Then
        If Len(o_address.p_Street_addr) > 0 Then
            Set o_address.o_error = o_error
            Set o_address.o_trans_data = o_trans_data
            o_address.p_obj_id = contact_id
            o_address.p_obj_nm = "contact"
            o_address.p_addr_type_cd = "STND"
            
            Call o_address.AddAddress
        End If
    Else
        Set o_address.o_error = o_error
        Set o_address.o_trans_data = o_trans_data
        o_address.p_obj_id = contact_id
        o_address.p_obj_nm = "contact"
        o_address.p_addr_type_cd = "STND"
        Call o_address.ChangeAddress
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


Public Sub DeleteContact()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.DeleteContact"
    
    Call IssueDelete
    Set o_address.o_error = o_error
    Set o_address.o_trans_data = o_trans_data
    o_address.p_obj_id = contact_id
    o_address.p_obj_nm = "contact"
    o_address.p_addr_type_cd = "STND"
    Call o_address.DeleteAddress
    
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

Private Sub Validate(varFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    strRoutine_nm = "contact.cls.validate"
    
    If varFunction_cd = "add" And fValuePresent(nm) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
        Exit Sub
        
        If Len(supplier_id) > 0 Then
            If IsNull(fGetSupplierNbr(supplier_id, o_error, o_trans_data)) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "0800"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 0
                Exit Sub
            End If
        Else
            If IsNull(fGetCustNbr(cust_id, o_error, o_trans_data)) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "0800"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 0
                Exit Sub
            End If
        End If
            
    End If
    
    If varFunction_cd = "change" And IsNull(nm) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0810"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
        Exit Sub
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

Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.IssueInsert"
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    contact_id = fGetId("contact_id", o_error, o_trans_data)
    
    strSql = "insert into contact ( "
    strSql = strSql & "  contact_id "
    strSql = strSql & ", email_addr "
    strSql = strSql & ", fax_nbr "
    strSql = strSql & ", nm "
    strSql = strSql & ", phone_nbr "
    strSql = strSql & ", title "
    strSql = strSql & ", cust_id "
    strSql = strSql & ", supplier_id "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(contact_id, "N")
    strSql = strSql & "," & fInsertVariable(email_addr, "S")
    strSql = strSql & "," & fInsertVariable(fax_nbr, "S")
    strSql = strSql & "," & fInsertVariable(nm, "S")
    strSql = strSql & "," & fInsertVariable(phone_nbr, "S")
    strSql = strSql & "," & fInsertVariable(title, "S")
    strSql = strSql & "," & fInsertVariable(cust_id, "N")
    strSql = strSql & "," & fInsertVariable(supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(Trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    
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
    Err.Clear
End Sub


Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.IssueUpdate"
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strSql = "update contact set "
    strSql = strSql & " " & fUpdateVariable("email_addr", email_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("fax_nbr", fax_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("nm", nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("phone_nbr", phone_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("title", title, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("contact_id", contact_id, "N", "=")
    
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
    Err.Clear
End Sub


Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "contact.cls.IssueDelete"
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strSql = "delete from contact "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("contact_id", contact_id, "N", "=")
    
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
    Err.Clear
End Sub


Public Sub getContacts()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Contact.cls.GetContacts"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    If Len(supplier_id) = 0 Then
        strSql = "select * from contact "
        strSql = strSql & "where "
        strSql = strSql & " " & fWhereVariable("cust_id", cust_id, "N", "=")
    Else
        strSql = "select * from contact "
        strSql = strSql & "where "
        strSql = strSql & " " & fWhereVariable("supplier_id", supplier_id, "N", "=")
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



