VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_message_id As Variant
Private mvarp_end_message_id As Variant
Private mvarp_message_desc As Variant
Private mvarp_class As Variant
Private mvaro_record_set As ADODB.Recordset 'local copy
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_message_id 'Let' and 'Get' are used to
'               set up the unique message id.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_message_id(ByVal vData As Variant)
    mvarp_message_id = vData
End Property
Public Property Get p_message_id() As Variant
    p_message_id = mvarp_message_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_end_message_id 'Let' and 'Get' are used to
'               set up the unique message id.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_end_message_id(ByVal vData As Variant)
    mvarp_end_message_id = vData
End Property
Public Property Get p_end_message_id() As Variant
    p_end_message_id = mvarp_end_message_id
End Property


'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_message_desc 'Let' and 'Get' are used to
'               set up the unique message desc.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_message_desc(ByVal vData As Variant)
    mvarp_message_desc = vData
End Property
Public Property Get p_message_desc() As Variant
    p_message_desc = mvarp_message_desc
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_class 'Let' and 'Get' are used to
'               set up the unique class.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_class(ByVal vData As Variant)
    mvarp_class = vData
End Property
Public Property Get p_class() As Variant
    p_class = mvarp_class
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_Trans_data 'Let' and 'Get' are used to
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
'Description:   The o_Error 'Let' and 'Get' are used to
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
Public Sub GetMessage()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

'define
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
    
'init
    strRoutine_nm = "message.cls.GetMessage"

    strSql = "SELECT *      "
    strSql = strSql & "FROM error_message "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("message_id", p_message_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    End If
    
    If objRecordSet.EOF = True Then
        p_message_id = Null
    Else
        p_message_id = objRecordSet("message_id")
        p_message_desc = objRecordSet("message_desc")
        p_class = objRecordSet("class")
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

'define
    Dim strRoutine_nm As String

'init
    strRoutine_nm = "message.cls.Add"
    
    Call Validate
    'p_message_id = fGetId("message_id", o_error, o_trans_data)
    Call IssueInsert
    
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

'define
    Dim strRoutine_nm As String
    
'init
    strRoutine_nm = "message.cls.Change"
    
    Call Validate
    Call IssueUpdate
    
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

    'define
    Dim strRoutine_nm As String
    
    'init
    strRoutine_nm = "message.cls.Delete"
    
    Call IssueDelete
    
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


'define
    Dim strRoutine_nm As String

'init
    strRoutine_nm = "message.cls.Validate"
    
    If Len(p_message_desc) > 0 Then
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

'define
    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
'init
    strRoutine_nm = "message.cls.IssueInsert"
    
    strSql = "INSERT INTO error_message ( " & _
                "  message_id " & _
                ", message_desc " & _
                ", class " & _
                ", updated_assoc_id" & _
                ", updated_ts" & _
                ", entered_assoc_id" & _
                ", entered_ts" & _
                " ) " & _
                "VALUES (" & _
                " " & fInsertVariable(p_message_id, "N") & _
                "," & fInsertVariable(p_message_desc, "S") & _
                "," & fInsertVariable(p_class, "S") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                ")"
    
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

'define
    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
'init
    strRoutine_nm = "message.cls.IssueUpdate"
    
    strSql = "UPDATE error_message SET " & _
               " " & fUpdateVariable("message_desc", mvarp_message_desc, "S", "=") & _
               "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=") & _
               "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
               " WHERE" & _
               " " & fWhereVariable("message_id", p_message_id, "N", "=")
    
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

'define
    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    
'init
    strRoutine_nm = "message.cls.IssueDelete"
    
    strSql = "DELETE FROM error_message "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("message_id", p_message_id, "N", "=")
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:       Gets the list of messages from the database,
'                   calls SELECT function
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetMessages()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "messages.cls.getMessages"
    Dim objRecordSet As ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT * FROM error_message where 1 = 1 "
    
    If IsNull(p_message_id) = True Or p_message_id = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("message_id", p_message_id, "N", "=")
    End If
    
    If IsNull(p_end_message_id) = True Or p_end_message_id = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("end_message_id", p_end_message_id, "N", "=")
    End If
 
    strSql = strSql & " ORDER BY message_id "

    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordSet = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordSet


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


