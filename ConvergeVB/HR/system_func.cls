VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "system_func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_system_func_cd As Variant
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_system_fuction_cd 'Let' and 'Get' are used to
'               set up the unique system fuction code.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_system_func_cd(ByVal vData As Variant)
    mvarp_system_func_cd = vData
End Property
Public Property Get p_system_func_cd() As Variant
    p_system_func_cd = mvarp_system_func_cd
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Gets the message from the database, calls SELECT function
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetSystemFunc()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
    
'init
    strRoutine_nm = "system_func.cls.GetSystemFunc"

    strSql = "SELECT *      "
    strSql = strSql & "FROM system_function "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("system_func_cd", p_system_func_cd, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Err.Clear
        Exit Sub
    End If
    
    If objRecordSet.EOF = True Then
        p_system_func_cd = Null
    Else
        p_system_func_cd = objRecordSet("system_func_cd")
    End If
     
    objRecordSet.Close
    Set objRecordSet = Nothing
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Adds message to database, calls Validate, GetMessageId and
'                   IssueInsert
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Add()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String

'init
    strRoutine_nm = "system_func.cls.Add"
    
    Call Validate
    Call IssueInsert
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Chenge message in DB, calls Validate and IssueUpdate
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
'init
    strRoutine_nm = "system_func.cls.Change"
    
    Call Validate
    Call IssueUpdate
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Deletes a message record from Database, calls IssueDelete
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    'init
    strRoutine_nm = "system_func.cls.Delete"
    
    Call IssueDelete
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:       Makes sure there is a message description
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Validate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String

'init
    strRoutine_nm = "system_func.cls.Validate"
    
    If Len(p_system_func_cd) > 0 Then
    Else
        o_error.type_cd = "F"
        o_error.err_cd = "0600"
        o_error.message_id = 0
        o_error.routine_nm = strRoutine_nm
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
'Description:       Puts the new record into Database, calls INSERT
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Integer
    
    strRoutine_nm = "system_func.cls.IssueInsert"
    
    strSql = "INSERT INTO system_function ( " & _
                "  system_func_cd " & _
                ", updated_assoc_id" & _
                ", updated_ts" & _
                ", entered_assoc_id" & _
                ", entered_ts" & _
                " ) " & _
                "VALUES (" & _
                " " & fInsertVariable(p_system_func_cd, "S") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                ")"
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Updates a message record in Database, calls UPDATE
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Integer
    
    strRoutine_nm = "system_func.cls.IssueUpdate"
    
    strSql = "UPDATE system_function SET " & _
               " " & fUpdateVariable("system_func_cd", p_system_func_cd, "S", "=") & _
               "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=") & _
               "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
               " WHERE" & _
               " " & fWhereVariable("system_func_cd", p_system_func_cd, "S", "=")
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:       Deletes record from Database, calls DELETE
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Integer
    
    strRoutine_nm = "system_func.cls.IssueDelete"
    
    strSql = "DELETE FROM system_function "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("system_func_cd", p_system_func_cd, "S", "=")
    
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

