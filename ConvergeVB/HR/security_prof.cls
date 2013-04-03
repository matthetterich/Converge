VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "security_prof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_security_prof_id As Variant
Private mvarp_security_prof_nm As Variant
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_security_prof_id 'Let' and 'Get' are used to
'               set up the unique security_prof id.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_security_prof_id(ByVal vData As Variant)
    mvarp_security_prof_id = vData
End Property
Public Property Get p_security_prof_id() As Variant
    p_security_prof_id = mvarp_security_prof_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_system_fuction_cd 'Let' and 'Get' are used to
'               set up the unique system fuction code.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_security_prof_nm(ByVal vData As Variant)
    mvarp_security_prof_nm = vData
End Property
Public Property Get p_security_prof_nm() As Variant
    p_security_prof_nm = mvarp_security_prof_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_Trans_data 'Set' and 'Get' are used to
'               set up the unique Trans data.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_Error 'Set' and 'Get' are used to
'               set up the unique Error.
'
'//////////////////////////////////////////////////////////////////////////////
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
Public Sub GetSecurityProf()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
'init
    strRoutine_nm = "security_prof.cls.GetSecurityProf"

    strSql = "SELECT *      "
    strSql = strSql & "FROM security_prof "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("security_prof_id", p_security_prof_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_security_prof_id = Null
    Else
        p_security_prof_id = objRecordset("security_prof_id")
        p_security_prof_nm = objRecordset("security_prof_nm")
    End If
        
    Call objRecordset.Close
    Set objRecordset = Nothing
    
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
    strRoutine_nm = "security_prof.cls.Add"
    
    Call Validate
    p_security_prof_id = fGetId("security_prof_id", o_error, o_trans_data)
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
    strRoutine_nm = "security_prof.cls.Change"
    
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
    strRoutine_nm = "security_prof.cls.Delete"
    
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
    strRoutine_nm = "security_prof.cls.Validate"
    
    If Len(p_security_prof_nm) > 0 Then
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
    Dim dblUpdatedQty As Double
    Dim rstData As ADODB.Recordset
    
    strRoutine_nm = "security_prof.cls.IssueInsert"
    
    strSql = "INSERT INTO security_prof ( " & _
                "  security_prof_id " & _
                ", security_prof_nm " & _
                ", updated_assoc_id" & _
                ", updated_ts" & _
                ", entered_assoc_id" & _
                ", entered_ts" & _
                " ) " & _
                "VALUES (" & _
                " " & fInsertVariable(p_security_prof_id, "N") & _
                "," & fInsertVariable(p_security_prof_nm, "S") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                ")"
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
                                        ' Add all internal functions to the
                                        ' security authorization file.
    strSql = "SELECT system_function_cd " & _
                "FROM system_function " & _
                "WHERE internal_function_ind = 'Y'"
    Call o_trans_data.OpenRecordset(rstData, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstData.Close
        Set rstData = Nothing
        Err.Clear
        Exit Sub
    End If
    
    Do While rstData.EOF = False
        strSql = "INSERT INTO security_auth (" & _
                    "  security_prof_id " & _
                    ", system_function_cd " & _
                    ", updated_assoc_id " & _
                    ", updated_ts " & _
                    ", entered_assoc_id " & _
                    ", entered_ts" & _
                    " ) " & _
                    "VALUES (" & _
                    " " & fInsertVariable(p_security_prof_id, "N") & _
                    "," & fInsertVariable(rstData("system_function_cd"), "S") & _
                    "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                    "," & fInsertVariable("current_ts", "D") & _
                    "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                    "," & fInsertVariable("current_ts", "D") & _
                    ")"
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
                                        ' Insert the record
        rstData.MoveNext
    Loop
                                        ' Destroy object
    rstData.Close
    Set rstData = Nothing
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
    
    strRoutine_nm = "security_prof.cls.IssueUpdate"
    
    strSql = "UPDATE security_prof SET " & _
               " " & fUpdateVariable("security_prof_nm", mvarp_security_prof_nm, "S", "=") & _
               "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=") & _
               "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
               " WHERE" & _
               " " & fWhereVariable("security_prof_id", p_security_prof_id, "N", "=")
    
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
    Dim rstData As ADODB.Recordset
    Dim dblUpdatedQty As Double
    
    strRoutine_nm = "security_prof.cls.IssueDelete"
    
    strSql = "SELECT assoc_id " & _
                "FROM associate " & _
                "WHERE " & _
                    fWhereVariable("security_prof_id", p_security_prof_id, "N", "=")
    Call o_trans_data.OpenRecordset(rstData, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstData.Close
        Set rstData = Nothing
        Err.Clear
        Exit Sub
    End If
    If rstData.EOF = False Then
        o_error.type_cd = "E"
        o_error.err_cd = "0100"
        o_error.message_id = 2199
        o_error.routine_nm = strRoutine_nm
        rstData.Close
        Set rstData = Nothing
        Exit Sub
    End If
    rstData.Close
    Set rstData = Nothing
                                        ' Delete all security authoriztion records
                                        ' that are using this profile.
    strSql = "DELETE FROM security_auth " & _
                "WHERE " & _
                fWhereVariable("security_prof_id", p_security_prof_id, "N", "=")
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
                                        ' Delete the security profile
    strSql = "DELETE FROM security_prof "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("security_prof_id", p_security_prof_id, "N", "=")
    
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




