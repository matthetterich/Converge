VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "trans_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_varGroup_id As Variant
Private m_varOrigin_type_cd As Variant
Private m_varOrigin_nm As Variant
Private m_varStatus_cd As Variant
Private m_varTrans_ts As Variant
Private m_varTrans_id As Variant
Private m_varAssoc_id As Variant
Private m_varFunction_cd As Variant
Private m_varConnectString As Variant
Private m_objConnection As ADODB.Connection
Private m_blnUpdate_fl As Boolean
Private m_varDatabase_nm As Variant
Private m_blnDatabase_Opened_fl As Boolean
Private m_varCodeDesc As Variant
Private m_varStateCode As Variant
Private m_varSecurity As Variant
Private m_varSuppliers As Variant
Private m_varCustomers As Variant
'

'  Customer List Variant Array
Public Property Let p_Customers(ByVal vData As Variant)
    m_varCustomers = vData
End Property
Public Property Get p_Customers() As Variant
    p_Customers = m_varCustomers
End Property

'  Supplier List Variant Array
Public Property Let p_Suppliers(ByVal vData As Variant)
    m_varSuppliers = vData
End Property
Public Property Get p_Suppliers() As Variant
    p_Suppliers = m_varSuppliers
End Property

' Code Description Variant Array
Public Property Let p_CodeDesc(ByVal vData As Variant)
    m_varCodeDesc = vData
End Property
Public Property Get p_CodeDesc() As Variant
    p_CodeDesc = m_varCodeDesc
End Property

' State Code Variant Array
Public Property Let p_StateCode(ByVal vData As Variant)
    m_varStateCode = vData
End Property
Public Property Get p_StateCode() As Variant
    p_StateCode = m_varStateCode
End Property

' Security Variant Array
Public Property Let p_Security(ByVal vData As Variant)
    m_varSecurity = vData
End Property
Public Property Get p_Security() As Variant
    p_Security = m_varSecurity
End Property

Public Property Set o_dbConnection(ByVal vData As ADODB.Connection)
    Set m_objConnection = vData
End Property
Public Property Get o_dbConnection() As ADODB.Connection
    Set o_dbConnection = m_objConnection
End Property
Public Property Let p_database_nm(ByVal vData As Variant)
    m_varDatabase_nm = vData
End Property
Public Property Get p_database_nm() As Variant
    p_database_nm = m_varDatabase_nm
End Property
Public Property Let p_update_fl(ByVal vData As Variant)
    m_blnUpdate_fl = vData
End Property
Public Property Get p_update_fl() As Variant
    p_update_fl = m_blnUpdate_fl
End Property
Public Property Let update_fl(ByVal vData As Variant)
    m_blnUpdate_fl = vData
End Property
Public Property Get update_fl() As Variant
    update_fl = m_blnUpdate_fl
End Property
Public Property Let p_ConnectString(ByVal vData As Variant)
    m_varConnectString = vData
End Property
Public Property Get p_ConnectString() As Variant
    p_ConnectString = m_varConnectString
End Property
Public Property Let p_Function_cd(ByVal vData As Variant)
    m_varFunction_cd = vData
End Property
Public Property Get p_Function_cd() As Variant
    p_Function_cd = m_varFunction_cd
End Property
Public Property Let p_assoc_id(ByVal vData As Variant)
    m_varAssoc_id = vData
End Property
Public Property Get p_assoc_id() As Variant
    p_assoc_id = m_varAssoc_id
End Property
Public Property Let p_trans_id(ByVal vData As Variant)
    m_varTrans_id = vData
End Property
Public Property Get p_trans_id() As Variant
    p_trans_id = m_varTrans_id
End Property
Public Property Let P_trans_ts(ByVal vData As Variant)
    m_varTrans_ts = vData
End Property
Public Property Get P_trans_ts() As Variant
    P_trans_ts = m_varTrans_ts
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    m_varStatus_cd = vData
End Property
Public Property Get p_status_cd() As Variant
    p_status_cd = m_varStatus_cd
End Property
Public Property Let p_origin_nm(ByVal vData As Variant)
    m_varOrigin_nm = vData
End Property
Public Property Get p_origin_nm() As Variant
    p_origin_nm = m_varOrigin_nm
End Property
Public Property Let p_origin_type_cd(ByVal vData As Variant)
    m_varOrigin_type_cd = vData
End Property
Public Property Get p_origin_type_cd() As Variant
    p_origin_type_cd = m_varOrigin_type_cd
End Property
Public Property Let p_group_id(ByVal vData As Variant)
    m_varGroup_id = vData
End Property
Public Property Get p_group_id() As Variant
    p_group_id = m_varGroup_id
End Property
Public Property Let ConnectString(ByVal vData As Variant)
    m_varConnectString = vData
End Property
Public Property Get ConnectString() As Variant
    ConnectString = m_varConnectString
End Property
Public Property Let Function_cd(ByVal vData As Variant)
    m_varFunction_cd = vData
End Property
Public Property Get Function_cd() As Variant
    Function_cd = m_varFunction_cd
End Property
Public Property Let assoc_id(ByVal vData As Variant)
    m_varAssoc_id = vData
End Property
Public Property Get assoc_id() As Variant
    assoc_id = m_varAssoc_id
End Property
Public Property Let trans_id(ByVal vData As Variant)
    m_varTrans_id = vData
End Property
Public Property Get trans_id() As Variant
    trans_id = m_varTrans_id
End Property
Public Property Let trans_ts(ByVal vData As Variant)
    m_varTrans_ts = vData
End Property
Public Property Get trans_ts() As Variant
    trans_ts = m_varTrans_ts
End Property
Public Property Let status_cd(ByVal vData As Variant)
    m_varStatus_cd = vData
End Property
Public Property Get status_cd() As Variant
    status_cd = m_varStatus_cd
End Property
Public Property Let origin_nm(ByVal vData As Variant)
    m_varOrigin_nm = vData
End Property
Public Property Get origin_nm() As Variant
    origin_nm = m_varOrigin_nm
End Property
Public Property Let origin_type_cd(ByVal vData As Variant)
    m_varOrigin_type_cd = vData
End Property
Public Property Get origin_type_cd() As Variant
    origin_type_cd = m_varOrigin_type_cd
End Property
Public Property Let group_id(ByVal vData As Variant)
    m_varGroup_id = vData
End Property
Public Property Get group_id() As Variant
    group_id = m_varGroup_id
End Property
Public Property Let p_database_opened_fl(ByVal vData As Variant)
    m_blnDatabase_Opened_fl = vData
End Property
Public Property Get p_database_opened_fl() As Variant
    p_database_opened_fl = m_blnDatabase_Opened_fl
End Property

Public Sub OpenDatabase(ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.OpenDatabase"
    
    If p_database_opened_fl = True Then
    Else
        p_database_opened_fl = True
        Set o_dbConnection = New ADODB.Connection
        Call o_dbConnection.Open(p_ConnectString)
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


Public Sub BeginTransDatabase(ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.BeginTransDatabase"
    On Error Resume Next
    o_dbConnection.RollbackTrans
    On Error GoTo error_handler
    o_dbConnection.BeginTrans

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

Public Sub OpenRecordset(ByRef r_objRecordset As ADODB.Recordset, ByVal v_strSql As String, ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.OpenRecordset"
    
    Set r_objRecordset = New ADODB.Recordset
    With r_objRecordset
        .ActiveConnection = o_dbConnection
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open (v_strSql)
    End With
    
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



Public Sub ExecuteDatabase(ByVal v_strSql As String, ByRef r_dblUpdateQty As Double, ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.ExecuteDatabase"
    
    If p_update_fl = False Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_message_id = 0
            .p_routine_nm = strRoutine_nm
            .p_sqlstate_cd = o_dbConnection.Errors(0).SQLState
            .p_sql_tx = v_strSql
        End With
        Exit Sub
    End If
    
    Call o_dbConnection.Execute(v_strSql, r_dblUpdateQty)
'    Debug.Print r_dblUpdateQty & " " & v_strSql
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
      .p_sql_tx = v_strSql
    End With
    Err.Clear
End Sub



Public Sub CommitTransDatabase(ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.CommitTransDatabase"
    
    
    o_dbConnection.CommitTrans

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



Public Sub RollBackTransDatabase(ByRef o_error As Object)
    
    On Error GoTo error_handler
    
    Dim strRoutine_nm As String
    strRoutine_nm = "tran_data.cls.RollBackTransDatabase"
    
    o_dbConnection.RollbackTrans

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



