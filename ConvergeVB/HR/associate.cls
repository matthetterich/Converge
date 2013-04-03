VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "associate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_name As Object
Private mvaro_address As Object
Private mvarp_new_password As Variant
Private mvarp_current_password As Variant
Private mvarp_confirmation As Variant

Private mvarp_assoc_id As Variant
Private mvarp_assoc_nbr As Variant
Private mvarp_birth_dt As Variant
Private mvarp_start_dt As Variant
Private mvarp_status_cd As Variant
Private mvarp_type_cd As Variant
Private mvarp_system_pw As Variant
Private mvarp_security_prof_id As Variant
Private mvarp_security_prof_nbr As Variant
Private mvarp_status_desc As Variant
Private mvarp_left_portal_1_cd As Variant
Private mvarp_left_portal_1_desc As Variant
Private mvarp_left_portal_2_cd As Variant
Private mvarp_left_portal_2_desc As Variant
Private mvarp_left_portal_3_cd As Variant
Private mvarp_left_portal_3_desc As Variant
Private mvarp_right_portal_1_cd As Variant
Private mvarp_right_portal_1_desc As Variant
Private mvarp_right_portal_2_cd As Variant
Private mvarp_right_portal_2_desc As Variant
Private mvarp_right_portal_3_cd As Variant
Private mvarp_right_portal_3_desc As Variant
Private mvarp_type_desc As Variant
Private mvaro_record_set As ADODB.Recordset
'

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

' Codes for the left side of the portal screens
Public Property Let p_left_portal_1_cd(ByVal vData As Variant)
    mvarp_left_portal_1_cd = vData
End Property
Public Property Get p_left_portal_1_cd() As Variant
        p_left_portal_1_cd = mvarp_left_portal_1_cd
End Property
Public Property Let p_left_portal_2_cd(ByVal vData As Variant)
    mvarp_left_portal_2_cd = vData
End Property
Public Property Get p_left_portal_2_cd() As Variant
        p_left_portal_2_cd = mvarp_left_portal_2_cd
End Property
Public Property Let p_left_portal_3_cd(ByVal vData As Variant)
    mvarp_left_portal_3_cd = vData
End Property
Public Property Get p_left_portal_3_cd() As Variant
        p_left_portal_3_cd = mvarp_left_portal_3_cd
End Property

Public Property Let p_left_portal_1_desc(ByVal vData As Variant)
    mvarp_left_portal_1_desc = vData
End Property
Public Property Get p_left_portal_1_desc() As Variant
        p_left_portal_1_desc = mvarp_left_portal_1_desc
End Property
Public Property Let p_left_portal_2_desc(ByVal vData As Variant)
    mvarp_left_portal_2_desc = vData
End Property
Public Property Get p_left_portal_2_desc() As Variant
        p_left_portal_2_desc = mvarp_left_portal_2_desc
End Property
Public Property Let p_left_portal_3_desc(ByVal vData As Variant)
    mvarp_left_portal_3_desc = vData
End Property
Public Property Get p_left_portal_3_desc() As Variant
        p_left_portal_3_desc = mvarp_left_portal_3_desc
End Property


' Codes for the left side of the portal screens
Public Property Let p_right_portal_1_cd(ByVal vData As Variant)
    mvarp_right_portal_1_cd = vData
End Property
Public Property Get p_right_portal_1_cd() As Variant
        p_right_portal_1_cd = mvarp_right_portal_1_cd
End Property
Public Property Let p_right_portal_2_cd(ByVal vData As Variant)
    mvarp_right_portal_2_cd = vData
End Property
Public Property Get p_right_portal_2_cd() As Variant
        p_right_portal_2_cd = mvarp_right_portal_2_cd
End Property
Public Property Let p_right_portal_3_cd(ByVal vData As Variant)
    mvarp_right_portal_3_cd = vData
End Property
Public Property Get p_right_portal_3_cd() As Variant
        p_right_portal_3_cd = mvarp_right_portal_3_cd
End Property

Public Property Let p_right_portal_1_desc(ByVal vData As Variant)
    mvarp_right_portal_1_desc = vData
End Property
Public Property Get p_right_portal_1_desc() As Variant
        p_right_portal_1_desc = mvarp_right_portal_1_desc
End Property
Public Property Let p_right_portal_2_desc(ByVal vData As Variant)
    mvarp_right_portal_2_desc = vData
End Property
Public Property Get p_right_portal_2_desc() As Variant
        p_right_portal_2_desc = mvarp_right_portal_2_desc
End Property
Public Property Let p_right_portal_3_desc(ByVal vData As Variant)
    mvarp_right_portal_3_desc = vData
End Property
Public Property Get p_right_portal_3_desc() As Variant
        p_right_portal_3_desc = mvarp_right_portal_3_desc
End Property


Public Property Let p_type_desc(ByVal vData As Variant)
    mvarp_type_desc = vData
End Property
Public Property Get p_type_desc() As Variant
        p_type_desc = mvarp_type_desc
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property
Public Property Let p_security_prof_nbr(ByVal vData As Variant)
    mvarp_security_prof_nbr = vData
End Property
Public Property Get p_security_prof_nbr() As Variant
        p_security_prof_nbr = mvarp_security_prof_nbr
End Property
Public Property Let p_security_prof_id(ByVal vData As Variant)
    mvarp_security_prof_id = vData
End Property
Public Property Get p_security_prof_id() As Variant
        p_security_prof_id = mvarp_security_prof_id
End Property
Public Property Let p_system_pw(ByVal vData As Variant)
    mvarp_system_pw = vData
End Property
Public Property Get p_system_pw() As Variant
        p_system_pw = mvarp_system_pw
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    mvarp_type_cd = vData
End Property
Public Property Get p_type_cd() As Variant
        p_type_cd = mvarp_type_cd
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_start_dt(ByVal vData As Variant)
    mvarp_start_dt = vData
End Property
Public Property Get p_start_dt() As Variant
        p_start_dt = mvarp_start_dt
End Property
Public Property Let p_birth_dt(ByVal vData As Variant)
    mvarp_birth_dt = vData
End Property
Public Property Get p_birth_dt() As Variant
        p_birth_dt = mvarp_birth_dt
End Property
Public Property Let p_assoc_nbr(ByVal vData As Variant)
    mvarp_assoc_nbr = vData
End Property
Public Property Get p_assoc_nbr() As Variant
        p_assoc_nbr = mvarp_assoc_nbr
End Property
Public Property Let p_assoc_id(ByVal vData As Variant)
    mvarp_assoc_id = vData
End Property
Public Property Get p_assoc_id() As Variant
        p_assoc_id = mvarp_assoc_id
End Property

Public Property Let p_current_password(ByVal vData As Variant)
    mvarp_current_password = vData
End Property
Public Property Get p_current_password() As Variant
        p_current_password = mvarp_current_password
End Property
Public Property Let p_new_password(ByVal vData As Variant)
    mvarp_new_password = vData
End Property
Public Property Get p_new_password() As Variant
        p_new_password = mvarp_new_password
End Property
Public Property Let p_confirmation(ByVal vData As Variant)
    mvarp_confirmation = vData
End Property
Public Property Get p_confirmation() As Variant
        p_confirmation = mvarp_confirmation
End Property

Public Property Set o_address(ByVal vData As Object)
    Set mvaro_address = vData
End Property
Public Property Get o_address() As Object
    Set o_address = mvaro_address
End Property
Public Property Set o_name(ByVal vData As Object)
    Set mvaro_name = vData
End Property
Public Property Get o_name() As Object
    Set o_name = mvaro_name
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
Public Sub GetAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "associate.cls.GetAssociate"
    
    Set o_name = CreateObject("converge.name")
    Set o_address = CreateObject("converge.address")
    
    Call SelectAssociate
    
    If IsNull(p_assoc_id) = True Then
    Else
        Set o_name.o_trans_data = o_trans_data
        Set o_name.o_error = o_error
        o_name.p_nm_id = Null
        o_name.p_obj_nm = "associate"
        o_name.p_obj_id = p_assoc_id
        o_name.p_nm_type_cd = "STND"
        o_name.p_def_nm_fl = "Y"
        Call o_name.getDefName
        
        Set o_address.o_trans_data = o_trans_data
        Set o_address.o_error = o_error
        o_address.p_addr_id = Null
        o_address.p_obj_nm = "associate"
        o_address.p_obj_id = p_assoc_id
        o_address.p_addr_type_cd = "STND"
        o_address.p_def_addr_fl = "Y"
        Call o_address.getDefAddress
    End If
    
    
    
    Call Err.Clear

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

    
Public Sub AddAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "associate.cls.AddAssociate"
    
    Call ValidateAddAssociate
    Call ValidateAssociate
    p_assoc_id = fGetId("assoc_id", o_error, o_trans_data)
    Call IssueInsertAssociate
    
    o_name.p_obj_nm = "associate"
    o_name.p_obj_id = p_assoc_id
    Call AddName(o_name, o_trans_data, o_error)
    o_address.p_obj_nm = "associate"
    o_address.p_obj_id = p_assoc_id
    Call AddAddr(o_address, o_trans_data, o_error)
    
    Call Err.Clear
    
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


Public Sub ChangeAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "associate.cls.ChangeAssociate"
    
    Call ValidateAssociate
    Call IssueUpdateAssociate
    
    o_name.p_obj_nm = "associate"
    o_name.p_obj_id = p_assoc_id
    
    Call ChangeName(o_name, o_trans_data, o_error)
    
    o_address.p_obj_nm = "associate"
    o_address.p_obj_id = p_assoc_id
    
    Call ChangeAddr(o_address, o_trans_data, o_error)
    
    Call Err.Clear
    
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


Public Sub DeleteAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "associate.cls.DeleteAssociate"
    
    Call IssueDeleteAssociate
    
    o_name.p_obj_nm = "associate"
    o_name.p_obj_id = p_assoc_id
    
    Call DeleteName(o_name, o_trans_data, o_error)
    
    o_address.p_obj_nm = "associate"
    o_address.p_obj_id = p_assoc_id
    
    Call DeleteAddr(o_address, o_trans_data, o_error)
    
    Call Err.Clear
    
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



Private Sub IssueDeleteAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "associate.cls.IssueDeleteAssociate"

    strSql = "delete from associate "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("assoc_id", p_assoc_id, "N", "=")
    
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



Private Sub ValidateAddAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "associate.cls.ValidateAssociate"
    
    If fValuePresent(p_assoc_nbr) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0100"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2130
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

Private Sub ValidateAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "associate.cls.ValidateAssociate"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    If IsNull(p_assoc_nbr) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2130
        Exit Sub
    End If
    
    If IsNull(p_assoc_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM associate "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("assoc_nbr", p_assoc_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM associate "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("assoc_nbr", p_assoc_nbr, "S", "=")
        strSql = strSql & "and" & fWhereVariable("assoc_id", p_assoc_id, "N", "<>")
    End If
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    If objrecordset.EOF = True Then
    Else
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1150"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2201
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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


Private Sub IssueInsertAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "associate.cls.IssueInsertAssociate"
    
    strSql = "insert into associate ( "
    strSql = strSql & "  assoc_id "
    strSql = strSql & ", assoc_nbr "
    strSql = strSql & ", birth_dt "
    strSql = strSql & ", start_dt "
    strSql = strSql & ", status_cd "
    strSql = strSql & ", type_cd "
    strSql = strSql & ", system_pw "
    strSql = strSql & ", security_prof_id "
    strSql = strSql & ", left_portal_1_cd "
    strSql = strSql & ", left_portal_2_cd "
    strSql = strSql & ", left_portal_3_cd "
    strSql = strSql & ", right_portal_1_cd "
    strSql = strSql & ", right_portal_2_cd "
    strSql = strSql & ", right_portal_3_cd "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable(p_assoc_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_birth_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_start_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_system_pw, "S")
    strSql = strSql & "," & fInsertVariable(p_security_prof_id, "N")
    strSql = strSql & "," & fInsertVariable(p_left_portal_1_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_left_portal_2_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_left_portal_3_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_right_portal_1_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_right_portal_2_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_right_portal_3_cd, "S")
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

Private Sub IssueUpdateAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "associate.cls.IssueUpdateAssociate"
    
    strSql = "update associate set "
    strSql = strSql & " " & fUpdateVariable("assoc_nbr", p_assoc_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("birth_dt", p_birth_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("start_dt", p_start_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("type_cd", p_type_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("system_pw", p_system_pw, "S", "=")
    strSql = strSql & "," & fUpdateVariable("security_prof_id", p_security_prof_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("left_portal_1_cd", p_left_portal_1_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("left_portal_2_cd", p_left_portal_2_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("left_portal_3_cd", p_left_portal_3_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("right_portal_1_cd", p_right_portal_1_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("right_portal_2_cd", p_right_portal_2_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("right_portal_3_cd", p_right_portal_3_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("assoc_id", p_assoc_id, "N", "=")
    
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



Private Sub SelectAssociate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "associate.cls.SelectAssociate"

    If IsNull(p_assoc_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM associate "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("assoc_nbr", p_assoc_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM associate "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("assoc_id", p_assoc_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    If objrecordset.EOF = True Then
        objrecordset.Close
        Set objrecordset = Nothing
        p_assoc_id = Null
    Else
        p_assoc_id = objrecordset("assoc_id")
        p_assoc_nbr = objrecordset("assoc_nbr")
        p_birth_dt = objrecordset("birth_dt")
        p_start_dt = objrecordset("start_dt")
        p_status_cd = objrecordset("status_cd")
        p_type_cd = objrecordset("type_cd")
        p_security_prof_id = objrecordset("security_prof_id")
        p_left_portal_1_cd = Trim(objrecordset("left_portal_1_cd"))
        p_left_portal_2_cd = Trim(objrecordset("left_portal_2_cd"))
        p_left_portal_3_cd = Trim(objrecordset("left_portal_3_cd"))
        p_right_portal_1_cd = Trim(objrecordset("right_portal_1_cd"))
        p_right_portal_2_cd = Trim(objrecordset("right_portal_2_cd"))
        p_right_portal_3_cd = Trim(objrecordset("right_portal_3_cd"))
        
        ' Close Connection
        objrecordset.Close
        Set objrecordset = Nothing
        
        p_type_desc = fGetDesc("assoc_type_cd", p_type_cd, o_error, o_trans_data)
        p_status_desc = fGetDesc("assoc_status_cd", p_status_cd, o_error, o_trans_data)
        p_left_portal_1_desc = fGetDesc("portal_cd", p_left_portal_1_cd, o_error, o_trans_data)
        p_left_portal_2_desc = fGetDesc("portal_cd", p_left_portal_2_cd, o_error, o_trans_data)
        p_left_portal_3_desc = fGetDesc("portal_cd", p_left_portal_3_cd, o_error, o_trans_data)
        p_right_portal_1_desc = fGetDesc("portal_cd", p_right_portal_1_cd, o_error, o_trans_data)
        p_right_portal_2_desc = fGetDesc("portal_cd", p_right_portal_2_cd, o_error, o_trans_data)
        p_right_portal_3_desc = fGetDesc("portal_cd", p_right_portal_3_cd, o_error, o_trans_data)
    
    End If
    
    Call Err.Clear
    
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

Public Sub GetAssociates()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Associate.cls.getAssociates"
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    Dim strPOStatus As String
        
    strPOStatus = "assoc_status_cd"
    
    strSql = "SELECT a.*, b.nm, c.security_prof_nm" & _
            "   FROM associate a, name b, security_prof c" & _
            "  WHERE (a.assoc_id = b.obj_id " & _
            "    AND a.security_prof_id = c.security_prof_id " & _
            "    AND b.obj_nm = 'associate') "
    
    If IsNull(p_assoc_nbr) = True Or p_assoc_nbr = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.assoc_nbr", p_assoc_nbr, "S", "=")
    End If
  
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
    
    If IsNull(p_type_cd) = True Or p_type_cd = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.type_cd", p_type_cd, "S", "=")
    End If
    
    If Len(p_security_prof_id) > 0 Then
       strSql = strSql & "  AND " & fWhereVariable("a.security_prof_id", p_security_prof_id, "S", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Err.Clear
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This routine will save the password changed by the user.
'
'Arguments:     None
'
'Returns:       True if valid; False if not
'//////////////////////////////////////////////////////////////////////////////

Private Function ValidData()
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    strSql = "SELECT system_pw " & _
             "FROM associate " & _
             "WHERE " & fWhereVariable("assoc_id", p_assoc_id, "N", "=")
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Err.Clear
        ValidData = False
        Exit Function
    End If
    If objrecordset.EOF Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0150"
        Let o_error.p_message_id = 2130
        Let o_error.p_routine_nm = "ValidData"
        Let o_error.p_desc = "Associate Not Found"
        Let o_error.p_sqlstate_cd = ""
        ValidData = False
        Exit Function
    End If
    
    If UCase(objrecordset("system_pw")) <> UCase(p_current_password) Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0200"
        Let o_error.p_message_id = 2131
        Let o_error.p_routine_nm = "ValidData"
        Let o_error.p_desc = "Password Authentication Failure"
        Let o_error.p_sqlstate_cd = ""
        ValidData = False
        Exit Function
    End If
    objrecordset.Close
    Set objrecordset = Nothing
        
    
    If UCase(p_new_password) <> UCase(p_confirmation) Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0100"
        Let o_error.p_message_id = 2129
        Let o_error.p_routine_nm = "ValidData"
        Let o_error.p_desc = "New password and Confirmation mismatch"
        Let o_error.p_sqlstate_cd = ""
        ValidData = False
        Exit Function
    End If
    
    ValidData = True
End Function

'//////////////////////////////////////////////////////////////////////////////
'Description:   This routine will save the password changed by the user.
'
'Arguments:     None
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub ChangePassword()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm       As String
    Dim strSql              As String
    Dim objrecordset        As ADODB.Recordset
    Dim dblUpdatedQty       As Double
    
    strRoutine_nm = "Associate.cls.ChangePassword"
    
    If Not ValidData() Then
        Exit Sub
    End If
    
    strSql = "UPDATE associate SET " & _
             fUpdateVariable("system_pw", p_new_password, "S", "=") & _
             " WHERE " & fWhereVariable("assoc_id", p_assoc_id, "N", "=")
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



