VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cust_group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_cust_grp_id As Variant
Private mvarp_cust_grp_nbr As Variant
Private mvarp_cust_grp_nm As Variant
Private mvaro_record_set As adodb.Recordset
'

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_cust_grp_nm(ByVal vData As Variant)
    mvarp_cust_grp_nm = vData
End Property
Public Property Set p_cust_grp_nm(ByVal vData As Variant)
    Set mvarp_cust_grp_nm = vData
End Property
Public Property Get p_cust_grp_nm() As Variant
    p_cust_grp_nm = mvarp_cust_grp_nm
End Property
Public Property Let p_cust_grp_nbr(ByVal vData As Variant)
    mvarp_cust_grp_nbr = vData
End Property
Public Property Get p_cust_grp_nbr() As Variant
    p_cust_grp_nbr = mvarp_cust_grp_nbr
End Property
Public Property Let p_cust_grp_id(ByVal vData As Variant)
    mvarp_cust_grp_id = vData
End Property
Public Property Get p_cust_grp_id() As Variant
    p_cust_grp_id = mvarp_cust_grp_id
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
Public Sub GetCustGrp()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.GetCustGrp"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If IsNull(p_cust_grp_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM cust_group "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_grp_nbr", p_cust_grp_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM cust_group "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_cust_grp_id = Null
    Else
        p_cust_grp_id = objRecordset("cust_grp_id")
        p_cust_grp_nbr = objRecordset("cust_grp_nbr")
        p_cust_grp_nm = objRecordset("cust_grp_nm")
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
    Err.clear
End Sub
    
Public Sub Add()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.Add"
    
    Call validate("add")
    p_cust_grp_id = fGetId("cust_grp_id", o_error, o_trans_data)
    Call IssueInsert
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.Change"
    
    Call validate("change")
    Call IssueUpdate
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0300"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.delete"
    
    Call validate("delete")
    Call IssueDelete
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "cust_grp.cls.IssueDelete"

    strSql = "delete from cust_Group "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0500"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub validate(ByVal strFunction_cd As String)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.Validate"
    
    If strFunction_cd = "add" Then
        Call Validateadd
        Call ValidateCust_grp_nbr
    End If
    
    If strFunction_cd = "change" Then
        If fValuePresent(p_cust_grp_nbr) = True Then
            Call ValidateCust_grp_nbr
        End If
    
        If IsNull(p_cust_grp_nbr) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "0600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1066
        End If
    
        If IsNull(p_cust_grp_nm) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "0700"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1067
        End If
    End If
    
    If strFunction_cd = "delete" Then
        Call ValidateDelete
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub Validateadd()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "cust_grp.cls.ValidateAdd"
    
    If fValuePresent(p_cust_grp_nbr) = False Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0850"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1065
    End If
    
    If fValuePresent(p_cust_grp_nm) = False Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0900"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1066
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0950"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ValidateDelete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.ValidateDelete"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strSql = "SELECT cust_id      "
    strSql = strSql & "FROM customer "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
        
    If objRecordset.EOF = True Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0975"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1072
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub ValidateCust_grp_nbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "cust_grp.cls.ValidateCust_grp_nbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If IsNull(p_cust_grp_nbr) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1065
        Exit Sub
    End If
    
    If IsNull(p_cust_grp_id) = True Then
        strSql = "SELECT *      "
        strSql = strSql & "FROM cust_group "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_grp_nbr", p_cust_grp_nbr, "S", "=")
    Else
        strSql = "SELECT *      "
        strSql = strSql & "FROM cust_group "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("cust_grp_nbr", p_cust_grp_nbr, "S", "=")
        strSql = strSql & "and" & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "<>")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "1300"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1070
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "cust_grp.cls.IssueInsert"
    
    strSql = "insert into cust_group ( "
    strSql = strSql & "  cust_grp_id "
    strSql = strSql & ", cust_grp_nbr "
    strSql = strSql & ", cust_grp_nm "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_cust_grp_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_grp_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_cust_grp_nm, "S")
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
      .p_err_cd = "1500"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "cust_grp.cls.IssueUpdate"
    
    strSql = "update cust_group set "
    strSql = strSql & " " & fUpdateVariable("cust_grp_nbr", p_cust_grp_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_grp_nm", p_cust_grp_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub getCustGrps()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strRoutine_nm = "cust_group.cls.getCustGrps"
    
    strSql = "SELECT  " & _
                " a.cust_grp_nm, " & _
                " a.cust_grp_nbr " & _
             " FROM cust_group a where 1 = 1 " _
    
    If IsNull(p_cust_grp_id) = True Or p_cust_grp_id = 0 Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.cust_grp_id", p_cust_grp_id, "N", "=")
    End If
                
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
      .p_err_cd = "1700"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


