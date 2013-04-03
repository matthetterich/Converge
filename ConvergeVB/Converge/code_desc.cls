VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "code_desc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As adodb.Recordset 'local copy
Private mvarp_field_nm As Variant
Private mvarp_seq_nbr As Variant
Private mvarp_code As Variant
Private mvarp_code_desc As Variant
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_record_set 'Set' and 'Get' are used to
'               set up the unique Recordset.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
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
'Description:   The p_field_nm 'Let' and 'Get' are used to
'               set up the unique Field Name.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_field_nm(ByVal vData As Variant)
    Let mvarp_field_nm = vData
End Property
Public Property Get p_field_nm() As Variant
    Let p_field_nm = mvarp_field_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_seq_nbr 'Let' and 'Get' are used to
'               set up the unique Sequence Number.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_seq_nbr(ByVal vData As Variant)
    Let mvarp_seq_nbr = vData
End Property
Public Property Get p_seq_nbr() As Variant
    Let p_seq_nbr = mvarp_seq_nbr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_code 'Let' and 'Get' are used to
'               set up the unique Code.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_code(ByVal vData As Variant)
    Let mvarp_code = vData
End Property
Public Property Get p_code() As Variant
    Let p_code = mvarp_code
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_code_desc 'Let' and 'Get' are used to
'               set up the unique Code.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_code_desc(ByVal vData As Variant)
    Let mvarp_code_desc = vData
End Property
Public Property Get p_code_desc() As Variant
    Let p_code_desc = mvarp_code_desc
End Property

Public Sub GetCodeDescs()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strFieldNm As String
    strRoutine_nm = "code_desc.cls.GetCodeDescs"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = "SELECT * from code_desc WHERE 1 = 1"
    
    If Len(p_field_nm) > 0 Then
        If InStr(1, p_field_nm, "*", vbBinaryCompare) > 0 Then
            strFieldNm = Left(p_field_nm, InStr(1, p_field_nm, "*", vbBinaryCompare) - 1)
            strSql = strSql & " AND field_nm LIKE '" & strFieldNm & "%'"
        Else
            strSql = strSql & " AND " & fWhereVariable("field_nm", p_field_nm, "S", "=")
        End If
    End If
        
    strSql = strSql & "ORDER BY field_nm, seq_nbr"
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
    
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


Public Sub GetCodeDesc()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "code_desc.cls.GetCodeDesc"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = " SELECT * from code_desc " & _
             " WHERE " & fWhereVariable("field_nm", p_field_nm, "S", "=") & _
             " AND " & fWhereVariable("seq_nbr", p_seq_nbr, "N", "=")
             
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_field_nm = Null
        p_seq_nbr = Null
    Else
        p_field_nm = objRecordset("field_nm")
        p_seq_nbr = objRecordset("seq_nbr")
        p_code = objRecordset("code")
        p_code_desc = objRecordset("code_desc")
    End If
    
    objRecordset.Close
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

Public Sub Change()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "code_desc.cls.change"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = " UPDATE code_desc SET "
    strSql = strSql & " " & fUpdateVariable("code", p_code, "S", "=")
    strSql = strSql & "," & fUpdateVariable("code_desc", p_code_desc, "S", "=")
    strSql = strSql & " WHERE " & fWhereVariable("field_nm", p_field_nm, "S", "=")
    strSql = strSql & " AND " & fWhereVariable("seq_nbr", p_seq_nbr, "N", "=")
    
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

Public Sub Delete()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "code_desc.cls.delete"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = " DELETE FROM code_desc "
    strSql = strSql & " WHERE " & fWhereVariable("field_nm", p_field_nm, "S", "=")
    strSql = strSql & " AND " & fWhereVariable("seq_nbr", p_seq_nbr, "N", "=")
    
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

Public Sub Add()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "code_desc.cls.add"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    Call ValidateAdd
    Call IssueInsertData
    
    
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

Private Sub ValidateAdd()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "code_desc.cls.ValidateAdd"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = " SELECT * " & _
             " FROM code_desc " & _
             " WHERE " & fWhereVariable("field_nm", p_field_nm, "S", "=") & _
             " AND " & fWhereVariable("seq_nbr", p_seq_nbr, "N", "=")
             
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF And objRecordset.BOF Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "5400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3425
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

Private Sub IssueInsertData()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "code_desc.cls.IssueInsertData"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = " INSERT INTO code_desc ( "
    strSql = strSql & "  field_nm "
    strSql = strSql & ", seq_nbr "
    strSql = strSql & ", code "
    strSql = strSql & ", code_desc "
    strSql = strSql & " ) VALUES ( "
    strSql = strSql & " " & fInsertVariable(p_field_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_seq_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_code, "S")
    strSql = strSql & "," & fInsertVariable(p_code_desc, "S")
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
