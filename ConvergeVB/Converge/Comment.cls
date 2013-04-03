VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_objError As Object
Private m_objTrans_data As Object
Private m_objUpdated_assoc As Object
Private m_objEntered_assoc As Object
Private m_varComment_id As Variant
Private m_varComment_tx As Variant
Private m_varComment_desc As Variant
Private m_varObj_nm As Variant
Private m_varObj_id As Variant
Private m_varEntered_ts As Variant
Private m_varUpdated_ts As Variant
Private m_varEntered_dt As Variant
Private m_varUpdated_dt As Variant
Private m_varSortByDesc As Variant
Private mvaro_record_set As adodb.Recordset
Public Property Let setSortByDesc(ByVal inBoolean As Variant)
    m_varSortByDesc = inBoolean
End Property


Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_updated_dt(ByVal vData As Variant)
    m_varUpdated_dt = vData
End Property
Public Property Get p_updated_dt() As Variant
        p_updated_dt = m_varUpdated_dt
End Property
Public Property Let p_entered_dt(ByVal vData As Variant)
    m_varEntered_dt = vData
End Property
Public Property Get p_entered_dt() As Variant
        p_entered_dt = m_varEntered_dt
End Property
Public Property Let p_comment_desc(ByVal vData As Variant)
    m_varComment_desc = vData
End Property
Public Property Get p_comment_desc() As Variant
        p_comment_desc = m_varComment_desc
End Property
Public Property Set o_entered_assoc(ByVal vData As Object)
    Set m_objEntered_assoc = vData
End Property
Public Property Get o_entered_assoc() As Object
    Set o_entered_assoc = m_objEntered_assoc
End Property
Public Property Set o_updated_assoc(ByVal vData As Object)
    Set m_objUpdated_assoc = vData
End Property
Public Property Get o_updated_assoc() As Object
    Set o_updated_assoc = m_objUpdated_assoc
End Property
Public Property Let p_updated_ts(ByVal vData As Variant)
    m_varUpdated_ts = vData
End Property
Public Property Get p_updated_ts() As Variant
        p_updated_ts = m_varUpdated_ts
End Property
Public Property Let p_entered_ts(ByVal vData As Variant)
    m_varEntered_ts = vData
End Property
Public Property Get p_entered_ts() As Variant
        p_entered_ts = m_varEntered_ts
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
Public Property Let p_comment_id(ByVal vData As Variant)
    m_varComment_id = vData
End Property
Public Property Get p_comment_id() As Variant
        p_comment_id = m_varComment_id
End Property
Public Property Let p_comment_tx(ByVal vData As Variant)
    m_varComment_tx = vData
End Property
Public Property Get p_comment_tx() As Variant
        p_comment_tx = m_varComment_tx
End Property
Public Property Let p_obj_nm(ByVal vData As Variant)
    m_varObj_nm = vData
End Property
Public Property Get p_obj_nm() As Variant
        p_obj_nm = m_varObj_nm
End Property
Public Property Let p_obj_id(ByVal vData As Variant)
    m_varObj_id = vData
End Property
Public Property Get p_obj_id() As Variant
        p_obj_id = m_varObj_id
End Property







Public Property Let updated_dt(ByVal vData As Variant)
    m_varUpdated_dt = vData
End Property
Public Property Get updated_dt() As Variant
        updated_dt = m_varUpdated_dt
End Property
Public Property Let entered_dt(ByVal vData As Variant)
    m_varEntered_dt = vData
End Property
Public Property Get entered_dt() As Variant
        entered_dt = m_varEntered_dt
End Property
Public Property Let comment_desc(ByVal vData As Variant)
    m_varComment_desc = vData
End Property
Public Property Get comment_desc() As Variant
        comment_desc = m_varComment_desc
End Property
Public Property Set entered_assoc(ByVal vData As Object)
    Set m_objEntered_assoc = vData
End Property
Public Property Get entered_assoc() As Object
    Set entered_assoc = m_objEntered_assoc
End Property
Public Property Set updated_assoc(ByVal vData As Object)
    Set m_objUpdated_assoc = vData
End Property
Public Property Get updated_assoc() As Object
    Set updated_assoc = m_objUpdated_assoc
End Property
Public Property Let updated_ts(ByVal vData As Variant)
    m_varUpdated_ts = vData
End Property
Public Property Get updated_ts() As Variant
        updated_ts = m_varUpdated_ts
End Property
Public Property Let entered_ts(ByVal vData As Variant)
    m_varEntered_ts = vData
End Property
Public Property Get entered_ts() As Variant
        entered_ts = m_varEntered_ts
End Property
Public Property Set Trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get Trans_data() As Object
    Set Trans_data = m_objTrans_data
End Property
Public Property Set error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get error() As Object
    Set error = m_objError
End Property
Public Property Let comment_id(ByVal vData As Variant)
    m_varComment_id = vData
End Property
Public Property Get comment_id() As Variant
        comment_id = m_varComment_id
End Property
Public Property Let comment_tx(ByVal vData As Variant)
    m_varComment_tx = vData
End Property
Public Property Get comment_tx() As Variant
        comment_tx = m_varComment_tx
End Property
Public Property Let obj_nm(ByVal vData As Variant)
    m_varObj_nm = vData
End Property
Public Property Get obj_nm() As Variant
        obj_nm = m_varObj_nm
End Property
Public Property Let obj_id(ByVal vData As Variant)
    m_varObj_id = vData
End Property
Public Property Get obj_id() As Variant
        obj_id = m_varObj_id
End Property


Public Sub GetComment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comment.cls.GetComment"
    Dim strSql As String
    Dim objRecordset As adodb.Recordset

    Set o_updated_assoc = CreateObject("converge_hr.associate")
    Set o_entered_assoc = CreateObject("converge_hr.associate")
    
    strSql = "SELECT *      "
    strSql = strSql & "FROM comment "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("comment_id", p_comment_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If objRecordset.EOF = True Then
        p_comment_id = Null
    Else
        p_comment_id = objRecordset("comment_id")
        p_comment_desc = objRecordset("comment_desc")
        p_comment_tx = objRecordset("comment_tx")
        p_obj_nm = objRecordset("obj_nm")
        p_obj_id = objRecordset("obj_id")
        p_entered_ts = objRecordset("entered_ts")
        p_entered_dt = Format(m_varEntered_ts, "short Date")
        p_updated_ts = objRecordset("updated_ts")
        p_updated_dt = Format(m_varUpdated_ts, "short Date")
        
        Set o_entered_assoc.o_error = o_error
        Set o_entered_assoc.o_trans_data = o_trans_data
        o_entered_assoc.p_assoc_id = objRecordset("entered_assoc_id")
        Call o_entered_assoc.GetAssociate
        
        Set o_updated_assoc.o_error = o_error
        Set o_updated_assoc.o_trans_data = o_trans_data
        o_updated_assoc.p_assoc_id = objRecordset("updated_assoc_id")
        Call o_updated_assoc.GetAssociate
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


    
Public Sub AddComment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comment.cls.AddComment"
    
    Call ValidateComment
    p_comment_id = fGetId("comment_id", o_error, o_trans_data)
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


Public Sub ChangeComment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comment.cls.ChangeComment"
    
    Call ValidateComment
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


Public Sub DeleteComment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comment.cls.DeleteComment"
    
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



Private Sub ValidateComment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comment.cls.ValidateComment"
    
    If Len(p_comment_desc) > 0 Then
    Else
        o_error.type_cd = "F"
        o_error.err_cd = "0500"
        o_error.message_id = 0
        o_error.routine_nm = strRoutine_nm
    End If
    
    If Len(p_comment_tx) > 0 Then
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


Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "comment.cls.IssueInsert"
    
    strSql = "insert into comment ( "
    strSql = strSql & "  comment_id "
    strSql = strSql & ", comment_tx "
    strSql = strSql & ", comment_desc "
    strSql = strSql & ", obj_nm "
    strSql = strSql & ", obj_id "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_comment_id, "N")
    strSql = strSql & "," & fInsertVariable(p_comment_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_comment_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_obj_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_obj_id, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
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
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "comment.cls.IssueUpdate"
    
    strSql = "update comment set "
    strSql = strSql & " " & fUpdateVariable("comment_tx", p_comment_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("comment_desc", p_comment_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("comment_id", p_comment_id, "N", "=")
    
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
    strRoutine_nm = "comment.cls.IssueDelete"
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strSql = "delete from comment "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("comment_id", p_comment_id, "N", "=")
    
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


Public Sub GetComments()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "comment.cls.getCommentss"
    
    strSql = "select a.*, b.assoc_nbr from comment a, associate b "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("a.obj_nm", p_obj_nm, "S", "=")
    strSql = strSql & " and" & fWhereVariable("a.obj_id", p_obj_id, "N", "=")
    strSql = strSql & " and a.updated_assoc_id = b.assoc_id "
    If m_varSortByDesc = "true" Then
        strSql = strSql & " order by a.comment_desc"
    Else
        strSql = strSql & " order by a.entered_ts desc"
    End If

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
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



