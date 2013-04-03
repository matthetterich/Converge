VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "security_auths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As ADODB.Recordset
Private mvarp_security_prof_id As Variant
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_security_auth_id 'Let' and 'Get' are used to
'               set up the unique security_auth id.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_security_prof_id(ByVal vData As Variant)
    mvarp_security_prof_id = vData
End Property
Public Property Get p_security_prof_id() As Variant
    p_security_prof_id = mvarp_security_prof_id
End Property
'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_record_set 'Set' and 'Get' are used to
'               set up the record set.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_trans_data 'Set' and 'Get' are used to
'               set up the transaction data.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_error 'Set' and 'Get' are used to
'               set up the error object.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:       Gets the list of messages from the database,
'                   calls SELECT function
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetSecurityAuths()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    strRoutine_nm = "security_auths.cls.getSecurityAuths"
    
    strSql = "SELECT a.security_prof_id, a.system_function_cd, " & _
                "    b.system_function_name, b.internal_function_ind " & _
                "FROM security_auth a, system_function b " & _
                "WHERE " & _
                "   a.system_function_cd = b.system_function_cd AND " & _
                "   b.internal_function_ind = 'N' AND " & _
                fWhereVariable("a.security_prof_id", p_security_prof_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
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

