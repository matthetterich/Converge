VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_Street_addr As Variant
Private mvarp_cust_id As Variant
Private mvarp_City_nm As Variant
Private mvarp_State_cd As Variant
Private mvarp_state_nm As Variant
Private mvarp_Zip_cd As Variant
Private mvarp_Country_cd As Variant
Private mvarp_country_nm As Variant
Private mvarp_addr_id As Variant
Private mvarp_addr_type_cd As Variant
Private mvarp_def_addr_fl As Variant
Private mvaro_error As Object
Private mvarp_obj_nm As Variant
Private mvaro_trans_data As Object
Private mvarp_obj_id As Variant
Private mvaro_record_set As adodb.Recordset
'

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
        p_cust_id = mvarp_cust_id
End Property

Public Property Let p_state_nm(ByVal vData As Variant)
    mvarp_state_nm = vData
End Property
Public Property Get p_state_nm() As Variant
        p_state_nm = mvarp_state_nm
End Property
Public Property Let p_country_nm(ByVal vData As Variant)
    mvarp_country_nm = vData
End Property
Public Property Get p_country_nm() As Variant
        p_country_nm = mvarp_country_nm
End Property
Public Property Let p_obj_id(ByVal vData As Variant)
    mvarp_obj_id = vData
End Property
Public Property Get p_obj_id() As Variant
        p_obj_id = mvarp_obj_id
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
Public Property Let p_def_addr_fl(ByVal vData As Variant)
    mvarp_def_addr_fl = vData
End Property
Public Property Get p_def_addr_fl() As Variant
        p_def_addr_fl = mvarp_def_addr_fl
End Property
Public Property Let p_addr_type_cd(ByVal vData As Variant)
    mvarp_addr_type_cd = vData
End Property
Public Property Get p_addr_type_cd() As Variant
        p_addr_type_cd = mvarp_addr_type_cd
End Property
Public Property Let p_addr_id(ByVal vData As Variant)
    mvarp_addr_id = vData
End Property
Public Property Get p_addr_id() As Variant
        p_addr_id = mvarp_addr_id
End Property
Public Property Let p_Country_cd(ByVal vData As Variant)
    mvarp_Country_cd = vData
End Property
Public Property Get p_Country_cd() As Variant
        p_Country_cd = mvarp_Country_cd
End Property
Public Property Let p_Zip_cd(ByVal vData As Variant)
    mvarp_Zip_cd = vData
End Property
Public Property Get p_Zip_cd() As Variant
        p_Zip_cd = mvarp_Zip_cd
End Property
Public Property Let p_State_cd(ByVal vData As Variant)
    mvarp_State_cd = vData
End Property
Public Property Get p_State_cd() As Variant
        p_State_cd = mvarp_State_cd
End Property
Public Property Let p_City_nm(ByVal vData As Variant)
    mvarp_City_nm = vData
End Property
Public Property Get p_City_nm() As Variant
        p_City_nm = mvarp_City_nm
End Property
Public Property Let p_Street_addr(ByVal vData As Variant)
    mvarp_Street_addr = vData
End Property
Public Property Get p_Street_addr() As Variant
        p_Street_addr = mvarp_Street_addr
End Property
Public Property Let p_obj_nm(ByVal vData As Variant)
    mvarp_obj_nm = vData
End Property
Public Property Get p_obj_nm() As Variant
        p_obj_nm = mvarp_obj_nm
End Property
Public Sub getDefAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "address.cls.GetDefAddress"
    
    Dim strSql
    Dim objRecordset As adodb.Recordset

    strSql = "SELECT * "
    strSql = strSql & "FROM Address "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("obj_id", p_obj_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
    strSql = strSql & "and" & fWhereVariable("addr_type_cd", p_addr_type_cd, "S", "=")
    strSql = strSql & "and" & fWhereVariable("def_addr_fl", "Y", "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    If objRecordset.EOF = True Then
        p_addr_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_addr_id = objRecordset("addr_id")
        p_addr_type_cd = objRecordset("addr_type_cd")
        p_def_addr_fl = objRecordset("def_addr_fl")
        p_obj_nm = objRecordset("obj_nm")
        p_obj_id = objRecordset("obj_id")
        p_Street_addr = objRecordset("street_addr")
        p_City_nm = objRecordset("city_nm")
        p_State_cd = objRecordset("state_cd")
        p_Zip_cd = objRecordset("zip_cd")
        p_Country_cd = objRecordset("country_cd")
            
        objRecordset.Close
        Set objRecordset = Nothing
        
        p_country_nm = fGetDesc("country_cd", p_Country_cd, o_error, o_trans_data)
        p_state_nm = fGetDesc("state_cd", p_State_cd, o_error, o_trans_data)
    
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

    
Public Sub AddAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "address.cls.AddAddress"
    
    Call ValidateAddress
    Call IssueInsertAddress
    
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

Public Sub ChangeAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "address.cls.ChangeAddress"
    
    Call ValidateAddress
    Call IssueUpdateAddress
    
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


Public Sub DeleteAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "address.cls.DeleteAddress"
    
    Call IssueDeleteAddress
    
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
' Will delete all name associated to an obj id
Public Sub DeleteAll()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblupdatedQty As Double
    strRoutine_nm = "Address.cls.DeleteAddr"
    
    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_obj_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
   
    Call o_trans_data.ExecuteDatabase(strSql, dblupdatedQty, o_error)
  
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

Private Sub ValidateAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "address.cls.ValidateAddress"
    
    If p_City_nm = Empty Or IsNull(p_City_nm) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1500"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
    End If
    
    If p_State_cd = Empty Or _
       IsNull(p_State_cd) = True Or _
       fGetDesc("state_cd", p_State_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
    End If
    
    If IsNull(p_Zip_cd) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1700"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
    End If
    
    If p_Country_cd = Empty Or _
       IsNull(p_Country_cd) = True Or _
       fGetDesc("country_cd", p_Country_cd, o_error, o_trans_data) = Empty Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 0
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


Private Sub IssueInsertAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "address.cls.IssueInsertAddress"
    
    p_addr_id = fGetId("addr_id", o_error, o_trans_data)
   
    strSql = "insert into address ( "
    strSql = strSql & " addr_id"
    strSql = strSql & ",addr_type_cd"
    strSql = strSql & ",def_addr_fl"
    strSql = strSql & ",street_addr"
    strSql = strSql & ",city_nm"
    strSql = strSql & ",state_cd"
    strSql = strSql & ",zip_cd"
    strSql = strSql & ",country_cd"
    strSql = strSql & ",obj_nm"
    strSql = strSql & ",obj_id"
    strSql = strSql & ",updated_assoc_id"
    strSql = strSql & ",updated_ts"
    strSql = strSql & ",entered_assoc_id"
    strSql = strSql & ",entered_ts"
    strSql = strSql & ")"
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_addr_id, "N")
    strSql = strSql & "," & fInsertVariable(p_addr_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_def_addr_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_Street_addr, "S")
    strSql = strSql & "," & fInsertVariable(p_City_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_State_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_Zip_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_Country_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_obj_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_obj_id, "N")
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



Private Sub IssueUpdateAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "address.cls.IssueUpdateAddress"

    strSql = "update address set "
    strSql = strSql & " " & fUpdateVariable("street_addr", p_Street_addr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("city_nm", p_City_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("state_cd", p_State_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("zip_cd", p_Zip_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("country_cd", p_Country_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("def_addr_fl", p_def_addr_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & "where"
    strSql = strSql & "   " & fWhereVariable("addr_id", p_addr_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
    
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


Private Sub IssueDeleteAddress()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "address.cls.IssueDeleteAddress"

    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & "   " & fWhereVariable("addr_id", p_addr_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
    
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



