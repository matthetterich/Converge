VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_loc"
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
Private mvaro_stnd_name As Object
Private mvaro_stnd_address As Object
Private mvaro_ship_name As Object
Private mvaro_ship_address As Object
Private mvarp_inv_loc_id As Variant
Private mvarp_internal_loc_id As Variant
Private mvarp_consignment_fl As Variant
Private mvarp_def_inv_loc_fl As Variant
Private mvarp_inv_loc_nm As Variant
Private mvarp_inv_loc_desc As Variant
Private mvaro_record_set As adodb.Recordset
'

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

Public Property Set o_ship_address(ByVal vData As Object)
    Set mvaro_ship_address = vData
End Property
Public Property Get o_ship_address() As Object
    Set o_ship_address = mvaro_ship_address
End Property
Public Property Set o_ship_name(ByVal vData As Object)
    Set mvaro_ship_name = vData
End Property
Public Property Get o_ship_name() As Object
    Set o_ship_name = mvaro_ship_name
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property
Public Property Let p_def_inv_loc_fl(ByVal vData As Variant)
    mvarp_def_inv_loc_fl = vData
End Property
Public Property Get p_def_inv_loc_fl() As Variant
    p_def_inv_loc_fl = mvarp_def_inv_loc_fl
End Property
Public Property Let p_consignment_fl(ByVal vData As Variant)
    mvarp_consignment_fl = vData
End Property
Public Property Get p_consignment_fl() As Variant
    p_consignment_fl = mvarp_consignment_fl
End Property
Public Property Let p_inv_loc_desc(ByVal vData As Variant)
    mvarp_inv_loc_desc = vData
End Property
Public Property Get p_inv_loc_desc() As Variant
    p_inv_loc_desc = mvarp_inv_loc_desc
End Property
Public Property Let p_inv_loc_nm(ByVal vData As Variant)
    mvarp_inv_loc_nm = vData
End Property
Public Property Get p_inv_loc_nm() As Variant
    p_inv_loc_nm = mvarp_inv_loc_nm
End Property
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
    p_inv_loc_id = mvarp_inv_loc_id
End Property

Public Property Let p_internal_loc_id(ByVal vData As Variant)
    mvarp_internal_loc_id = vData
End Property
Public Property Get p_internal_loc_id() As Variant
    p_internal_loc_id = mvarp_internal_loc_id
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

Public Sub GetInventoryLoc()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.GetInventoryLoc"
    Dim objRecordset As adodb.Recordset
  
  
    Call SelectInventoryLoc(p_inv_loc_id, objRecordset, o_error, o_trans_data)
    Call BuildObjectProperties(objRecordset)
    Set o_ship_name = CreateObject("converge.name")
    Set o_ship_address = CreateObject("converge.address")
    Set o_stnd_name = CreateObject("converge.name")
    Set o_stnd_address = CreateObject("converge.address")
    
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_address = CreateObject("converge.address")
    
    
    If IsNull(p_internal_loc_id) = True Then
        ' STND Name and Address
        o_stnd_name.p_nm_id = Null
        o_stnd_name.p_nm = Null
        o_stnd_address.p_addr_id = Null
        o_stnd_address.p_city_nm = Null
        o_stnd_address.p_state_nm = Null
        o_stnd_address.p_country_nm = Null
        o_stnd_address.p_street_addr = Null
        
        ' Ship Name and Address
        o_ship_name.p_nm_id = Null
        o_ship_name.p_nm = Null
        o_ship_address.p_addr_id = Null
        o_ship_address.p_city_nm = Null
        o_ship_address.p_state_nm = Null
        o_ship_address.p_country_nm = Null
        o_ship_address.p_street_addr = Null
        
        ' Bill Name and Address
        o_bill_name.p_nm_id = Null
        o_bill_name.p_nm = Null
        o_bill_address.p_addr_id = Null
        o_bill_address.p_city_nm = Null
        o_bill_address.p_state_nm = Null
        o_bill_address.p_country_nm = Null
        o_bill_address.p_street_addr = Null
         
         
    Else
        Set o_stnd_name.o_trans_data = o_trans_data
        Set o_stnd_name.o_error = o_error
        o_stnd_name.p_nm_id = Null
        o_stnd_name.p_obj_nm = "inventory_loc"
        o_stnd_name.p_obj_id = p_internal_loc_id
        o_stnd_name.p_nm_type_cd = "STND"
        o_stnd_name.p_def_nm_fl = "Y"
        Call o_stnd_name.getDefName
                      
        Set o_stnd_address.o_trans_data = o_trans_data
        Set o_stnd_address.o_error = o_error
        o_stnd_address.p_addr_id = Null
        o_stnd_address.p_obj_nm = "inventory_loc"
        o_stnd_address.p_obj_id = p_internal_loc_id
        o_stnd_address.p_addr_type_cd = "STND"
        o_stnd_address.p_def_addr_fl = "Y"
        Call o_stnd_address.getDefAddress
    
        Set o_ship_name.o_trans_data = o_trans_data
        Set o_ship_name.o_error = o_error
        o_ship_name.p_nm_id = Null
        o_ship_name.p_obj_nm = "inventory_loc"
        o_ship_name.p_obj_id = p_internal_loc_id
        o_ship_name.p_nm_type_cd = "SHIP"
        o_ship_name.p_def_nm_fl = "Y"
        Call o_ship_name.getDefName
                      
        Set o_ship_address.o_trans_data = o_trans_data
        Set o_ship_address.o_error = o_error
        o_ship_address.p_addr_id = Null
        o_ship_address.p_obj_nm = "inventory_loc"
        o_ship_address.p_obj_id = p_internal_loc_id
        o_ship_address.p_addr_type_cd = "SHIP"
        o_ship_address.p_def_addr_fl = "Y"
        Call o_ship_address.getDefAddress
        
        Set o_bill_name.o_trans_data = o_trans_data
        Set o_bill_name.o_error = o_error
        o_bill_name.p_nm_id = Null
        o_bill_name.p_obj_nm = "inventory_loc"
        o_bill_name.p_obj_id = p_internal_loc_id
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_name.p_def_nm_fl = "Y"
        Call o_bill_name.getDefName
                      
        Set o_bill_address.o_trans_data = o_trans_data
        Set o_bill_address.o_error = o_error
        o_bill_address.p_addr_id = Null
        o_bill_address.p_obj_nm = "inventory_loc"
        o_bill_address.p_obj_id = p_internal_loc_id
        o_bill_address.p_addr_type_cd = "BILL"
        o_bill_address.p_def_addr_fl = "Y"
        Call o_bill_address.getDefAddress
        
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


Public Sub GetDefaultInventoryLoc()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.GetDefaultInventoryLoc"
    Dim objRecordset As adodb.Recordset
    Set o_stnd_name = CreateObject("converge.name")
    Set o_stnd_address = CreateObject("converge.address")
    Set o_ship_name = CreateObject("converge.name")
    Set o_ship_address = CreateObject("converge.address")
    Set o_bill_name = CreateObject("converge.name")
    Set o_bill_address = CreateObject("converge.address")
    
    Call SelectDefaultInventoryLoc(objRecordset)
    Call BuildObjectProperties(objRecordset)
    
    If IsNull(p_inv_loc_id) = True Then
        o_stnd_name.p_nm_id = Null
        o_stnd_address.p_addr_id = Null
    Else
        Set o_stnd_name.o_trans_data = o_trans_data
        Set o_stnd_name.o_error = o_error
        o_stnd_name.p_nm_id = Null
        o_stnd_name.p_obj_nm = "inventory_loc"
        o_stnd_name.p_obj_id = p_internal_loc_id
        o_stnd_name.p_nm_type_cd = "STND"
        o_stnd_name.p_def_nm_fl = "Y"
        Call o_stnd_name.getDefName
                      
        Set o_stnd_address.o_trans_data = o_trans_data
        Set o_stnd_address.o_error = o_error
        o_stnd_address.p_addr_id = Null
        o_stnd_address.p_obj_nm = "inventory_loc"
        o_stnd_address.p_obj_id = p_internal_loc_id
        o_stnd_address.p_addr_type_cd = "STND"
        o_stnd_address.p_def_addr_fl = "Y"
        Call o_stnd_address.getDefAddress
    
        ' Ship to Name and Address
        Set o_ship_name.o_trans_data = o_trans_data
        Set o_ship_name.o_error = o_error
        o_ship_name.p_nm_id = Null
        o_ship_name.p_obj_nm = "inventory_loc"
        o_ship_name.p_obj_id = p_internal_loc_id
        o_ship_name.p_nm_type_cd = "SHIP"
        o_ship_name.p_def_nm_fl = "Y"
        Call o_ship_name.getDefName
                      
        Set o_ship_address.o_trans_data = o_trans_data
        Set o_ship_address.o_error = o_error
        o_ship_address.p_addr_id = Null
        o_ship_address.p_obj_nm = "inventory_loc"
        o_ship_address.p_obj_id = p_internal_loc_id
        o_ship_address.p_addr_type_cd = "SHIP"
        o_ship_address.p_def_addr_fl = "Y"
        Call o_ship_address.getDefAddress
        
        ' Bill to Name and Address
        Set o_bill_name.o_trans_data = o_trans_data
        Set o_bill_name.o_error = o_error
        o_bill_name.p_nm_id = Null
        o_bill_name.p_obj_nm = "inventory_loc"
        o_bill_name.p_obj_id = p_internal_loc_id
        o_bill_name.p_nm_type_cd = "BILL"
        o_bill_name.p_def_nm_fl = "Y"
        Call o_bill_name.getDefName
                      
        Set o_bill_address.o_trans_data = o_trans_data
        Set o_bill_address.o_error = o_error
        o_bill_address.p_addr_id = Null
        o_bill_address.p_obj_nm = "inventory_loc"
        o_bill_address.p_obj_id = p_internal_loc_id
        o_bill_address.p_addr_type_cd = "BILL"
        o_bill_address.p_def_addr_fl = "Y"
        Call o_bill_address.getDefAddress
    
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

Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.Add"
    
    Call Validate("add")
    Call IssueInsertData
    
    If o_stnd_name.p_nm = Empty Then
    Else
        o_stnd_name.p_obj_nm = "inventory_loc"
        o_stnd_name.p_obj_id = p_internal_loc_id
        
        Call AddName(o_stnd_name, o_trans_data, o_error)
    End If
    
    If o_stnd_address.p_street_addr = Empty And o_stnd_address.p_city_nm = Empty And o_stnd_address.p_State_cd = Empty And o_stnd_address.p_Zip_cd = Empty And o_stnd_address.p_Country_cd = Empty Then
    Else
        o_stnd_address.p_obj_nm = "inventory_loc"
        o_stnd_address.p_obj_id = p_internal_loc_id
        
        Call AddAddr(o_stnd_address, o_trans_data, o_error)
    End If
 
    
    If o_ship_name.p_nm = Empty Then
    Else
        o_ship_name.p_obj_nm = "inventory_loc"
        o_ship_name.p_obj_id = p_internal_loc_id
        
        Call AddName(o_ship_name, o_trans_data, o_error)
    End If
    
    If o_ship_address.p_street_addr = Empty And o_ship_address.p_city_nm = Empty And o_ship_address.p_State_cd = Empty And o_ship_address.p_Zip_cd = Empty And o_ship_address.p_Country_cd = Empty Then
    Else
        o_ship_address.p_obj_nm = "inventory_loc"
        o_ship_address.p_obj_id = p_internal_loc_id
        
        Call AddAddr(o_ship_address, o_trans_data, o_error)
    End If
    
    If o_bill_name.p_nm = Empty Then
    Else
        o_bill_name.p_obj_nm = "inventory_loc"
        o_bill_name.p_obj_id = p_internal_loc_id
        
        Call AddName(o_bill_name, o_trans_data, o_error)
    End If
    
    
    If o_bill_address.p_street_addr = Empty And o_bill_address.p_city_nm = Empty And o_bill_address.p_State_cd = Empty And o_bill_address.p_Zip_cd = Empty And o_bill_address.p_Country_cd = Empty Then
    Else
        o_bill_address.p_obj_nm = "inventory_loc"
        o_bill_address.p_obj_id = p_internal_loc_id
        
        Call AddAddr(o_bill_address, o_trans_data, o_error)
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

Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.Change"
    
    Call Validate("change")
    Call IssueUpdateData
    o_stnd_name.p_obj_nm = "inventory_loc"
    o_stnd_name.p_obj_id = p_internal_loc_id
    o_stnd_address.p_obj_nm = "inventory_loc"
    o_stnd_address.p_obj_id = p_internal_loc_id
       
    o_ship_name.p_obj_nm = "inventory_loc"
    o_ship_name.p_obj_id = p_internal_loc_id
    o_ship_address.p_obj_nm = "inventory_loc"
    o_ship_address.p_obj_id = p_internal_loc_id
    
    o_bill_name.p_obj_nm = "inventory_loc"
    o_bill_name.p_obj_id = p_internal_loc_id
    o_bill_address.p_obj_nm = "inventory_loc"
    o_bill_address.p_obj_id = p_internal_loc_id
       
    ' Standard Name and Address
        If IsNull(o_stnd_name.p_nm_id) = True Or o_stnd_name.p_nm_id = Empty Then
            If IsNull(o_stnd_name.p_nm) = True Or o_stnd_name.p_nm = Empty Then
            Else
                
                Call AddName(o_stnd_name, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_stnd_name.p_nm) = True Then
                Call DeleteName(o_stnd_name, o_trans_data, o_error)
            Else
                Call ChangeName(o_stnd_name, o_trans_data, o_error)
            End If
        End If
        
        If IsNull(o_stnd_address.p_addr_id) = True Or o_stnd_address.p_addr_id = Empty Then
            If (IsNull(o_stnd_address.p_street_addr) = True Or o_stnd_address.p_street_addr = Empty) And _
               (IsNull(o_stnd_address.p_city_nm) = True Or o_stnd_address.p_city_nm = Empty) And _
               (IsNull(o_stnd_address.p_Zip_cd) = True Or o_stnd_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_stnd_address, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_stnd_address.p_street_addr) = True And _
               IsNull(o_stnd_address.p_city_nm) = True And _
               IsNull(o_stnd_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_stnd_address, o_trans_data, o_error)
            Else
                Call ChangeAddr(o_stnd_address, o_trans_data, o_error)
            End If
        End If
    
    ' Bill to Name and Address
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
            If (IsNull(o_bill_address.p_street_addr) = True Or o_bill_address.p_street_addr = Empty) And _
               (IsNull(o_bill_address.p_city_nm) = True Or o_bill_address.p_city_nm = Empty) And _
               (IsNull(o_bill_address.p_Zip_cd) = True Or o_bill_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_bill_address, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_bill_address.p_street_addr) = True And _
               IsNull(o_bill_address.p_city_nm) = True And _
               IsNull(o_bill_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_bill_address, o_trans_data, o_error)
            Else
                Call ChangeAddr(o_bill_address, o_trans_data, o_error)
            End If
        End If
    
    ' Ship to Name and Address
           
        If IsNull(o_ship_name.p_nm_id) = True Or o_ship_name.p_nm_id = Empty Then
            If IsNull(o_ship_name.p_nm) = True Or o_ship_name.p_nm = Empty Then
            Else
                
                Call AddName(o_ship_name, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_ship_name.p_nm) = True Then
                Call DeleteName(o_ship_name, o_trans_data, o_error)
            Else
                Call ChangeName(o_ship_name, o_trans_data, o_error)
            End If
        End If
        
        If IsNull(o_ship_address.p_addr_id) = True Or o_ship_address.p_addr_id = Empty Then
            If (IsNull(o_ship_address.p_street_addr) = True Or o_ship_address.p_street_addr = Empty) And _
               (IsNull(o_ship_address.p_city_nm) = True Or o_ship_address.p_city_nm = Empty) And _
               (IsNull(o_ship_address.p_Zip_cd) = True Or o_ship_address.p_Zip_cd = Empty) Then
            Else
                Call AddAddr(o_ship_address, o_trans_data, o_error)
            End If
        Else
            If IsNull(o_ship_address.p_street_addr) = True And _
               IsNull(o_ship_address.p_city_nm) = True And _
               IsNull(o_ship_address.p_Zip_cd) = True Then
                Call DeleteAddr(o_ship_address, o_trans_data, o_error)
            Else
                Call ChangeAddr(o_ship_address, o_trans_data, o_error)
            End If
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

Public Sub Delete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.Delete"
    o_stnd_name.p_obj_nm = "inventory_loc"
    o_stnd_name.p_obj_id = p_internal_loc_id
    
    o_stnd_address.p_obj_nm = "inventory_loc"
    o_stnd_address.p_obj_id = p_internal_loc_id
    
    o_bill_name.p_obj_nm = "inventory_loc"
    o_bill_name.p_obj_id = p_internal_loc_id
    
    o_bill_address.p_obj_nm = "inventory_loc"
    o_stnd_address.p_obj_id = p_internal_loc_id
    
    o_ship_name.p_obj_nm = "inventory_loc"
    o_ship_name.p_obj_id = p_internal_loc_id
    
    o_ship_address.p_obj_nm = "inventory_loc"
    o_ship_address.p_obj_id = p_internal_loc_id
    
    
    Call IssueDeleteData
    Call DeleteName(o_stnd_name, o_trans_data, o_error)
    Call DeleteAddr(o_stnd_address, o_trans_data, o_error)
    Call DeleteName(o_ship_name, o_trans_data, o_error)
    Call DeleteAddr(o_ship_address, o_trans_data, o_error)
    Call DeleteName(o_bill_name, o_trans_data, o_error)
    Call DeleteAddr(o_bill_address, o_trans_data, o_error)

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


Private Sub IssueDeleteData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "inventory_loc.cls.IssueDeleteData"

    strSql = "delete from inventory_loc "
    strSql = strSql & "where " & fWhereVariable("inv_loc_id", p_inv_loc_id, "S", "=")
    
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

Private Sub Validate(sFunction_cd)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.Validate"

    If sFunction_cd = "add" Then
        If p_consignment_fl = "Y" Or p_consignment_fl = "N" Then
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2212
        End If
        If fValuePresent(p_inv_loc_id) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1150"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1175
        End If
        Call ValidateID
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

Private Sub ValidateID()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_loc.cls.ValidateID"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    If IsNull(p_inv_loc_nm) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1900"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2213
        Exit Sub
    End If
    
    strSql = "SELECT * " & _
             "from inventory_loc " & _
            "WHERE " & fWhereVariable("inv_loc_id", p_inv_loc_id, "S", "=")
    
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
        o_error.p_message_id = 1175
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

Private Sub IssueInsertData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "inventory_loc.cls.IssueInsert"
    
    If p_def_inv_loc_fl = "Y" Then
        strSql = "update inventory_loc set "
        strSql = strSql & " " & fUpdateVariable("def_inv_loc_fl", "N", "S", "=")
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
            Exit Sub
        End If
    End If
    
    p_internal_loc_id = fGetId("internal_loc_id", o_error, o_trans_data)
    
    strSql = "insert into inventory_loc ( "
    strSql = strSql & "  inv_loc_id "
    strSql = strSql & ", inv_loc_nm " & _
                      ", internal_loc_id "
    strSql = strSql & ", inv_loc_desc "
    strSql = strSql & ", consignment_fl "
    strSql = strSql & ", def_inv_loc_fl "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_inv_loc_id, "S")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_nm, "S") & _
                      "," & fInsertVariable(p_internal_loc_id, "N")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_consignment_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_def_inv_loc_fl, "S")
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
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub IssueUpdateData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "inventory_loc.cls.IssueUpdateData"
    
    If p_def_inv_loc_fl = "Y" Then
        strSql = "update inventory_loc set "
        strSql = strSql & " " & fUpdateVariable("def_inv_loc_fl", "N", "S", "=")
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
            Exit Sub
        End If
    End If
    
    strSql = "update inventory_loc set "
    strSql = strSql & " " & fUpdateVariable("inv_loc_desc", p_inv_loc_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("inv_loc_nm", p_inv_loc_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("def_inv_loc_fl", p_def_inv_loc_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("inv_loc_id", p_inv_loc_id, "S", "=")
    
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


Private Sub SelectInventoryLoc(ByVal varInv_loc_id As Variant, ByRef objRecordset As Object, ByRef o_error As Object, ByRef o_trans_data As Object)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "inventory_loc.cls.Selectinventory_loc"
    
    strSql = "SELECT * " & _
             "from inventory_loc " & _
            "WHERE " & fWhereVariable("inv_loc_id", p_inv_loc_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
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



Private Sub SelectDefaultInventoryLoc(ByRef objRecordset As Object)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "inventory_loc.cls.SelectDefaultInventory_loc"

    strSql = "SELECT * " & _
             "from inventory_loc " & _
            "WHERE " & fWhereVariable("def_inv_loc_fl", "Y", "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
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


Private Sub BuildObjectProperties(ByRef objRecordset As Object)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "inventory_loc.cls.BuildObjectProperties"

    If objRecordset.EOF = True Then
        p_inv_loc_id = Null
        p_inv_loc_nm = Empty
        p_inv_loc_desc = Empty
        p_consignment_fl = Empty
    Else
        p_inv_loc_id = objRecordset("inv_loc_id")
        p_internal_loc_id = objRecordset("internal_loc_id")
        p_inv_loc_nm = objRecordset("inv_loc_nm")
        p_inv_loc_desc = objRecordset("inv_loc_desc")
        p_consignment_fl = objRecordset("consignment_fl")
        p_def_inv_loc_fl = objRecordset("def_inv_loc_fl")
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


Public Sub GetInventoryLocs()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strRoutine_nm = "inventory_locs.cls.getInventoryLocs"
    strSql = "select * from inventory_loc "

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



