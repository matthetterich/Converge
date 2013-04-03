VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarp_nm_id As Variant
Private mvarp_nm As Variant
Private mvarp_nm_type_cd As Variant
Private mvarp_def_nm_fl As Variant
Private mvaro_error As Object
Private mvarp_obj_nm As Variant
Private mvaro_trans_data As Object
Private mvarp_obj_id As Variant
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

Public Property Let p_obj_nm(ByVal vData As Variant)
    mvarp_obj_nm = vData
End Property

Public Property Get p_obj_nm() As Variant
        p_obj_nm = mvarp_obj_nm
End Property

Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Property Let p_def_nm_fl(ByVal vData As Variant)
    mvarp_def_nm_fl = vData
End Property
Public Property Get p_def_nm_fl() As Variant
        p_def_nm_fl = mvarp_def_nm_fl
End Property
Public Property Let p_nm(ByVal vData As Variant)
    mvarp_nm = vData
End Property
Public Property Get p_nm() As Variant
        p_nm = mvarp_nm
End Property
Public Property Let p_nm_type_cd(ByVal vData As Variant)
    mvarp_nm_type_cd = vData
End Property
Public Property Get p_nm_type_cd() As Variant
        p_nm_type_cd = mvarp_nm_type_cd
End Property
Public Property Let p_nm_id(ByVal vData As Variant)
    mvarp_nm_id = vData
End Property
Public Property Get p_nm_id() As Variant
        p_nm_id = mvarp_nm_id
End Property

Public Sub getDefName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "name.cls.GetDefName"
    
    Dim strSql
    Dim objRecordset As adodb.Recordset

    strSql = "SELECT * "
    strSql = strSql & "FROM Name "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("obj_id", p_obj_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
    strSql = strSql & "and" & fWhereVariable("nm_type_cd", p_nm_type_cd, "S", "=")
    strSql = strSql & "and" & fWhereVariable("def_nm_fl", "Y", "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    If objRecordset.EOF = True Then
        p_nm_id = Null
    Else
        p_nm_id = objRecordset("nm_id")
        p_nm = objRecordset("nm")
        p_nm_type_cd = objRecordset("nm_type_cd")
        p_def_nm_fl = objRecordset("def_nm_fl")
        p_obj_nm = objRecordset("obj_nm")
        p_obj_id = objRecordset("obj_id")
    End If

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


Public Sub AddName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "name.cls.AddName"
    
    Set objUtilities = CreateObject("converge.utilities")
    
    Call ValidateName
    Call IssueInsertName

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


Public Sub ChangeName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "name.cls.ChangeName"
        
    Call ValidateName
    Call IssueUpdateName

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


Public Sub DeleteName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "name.cls.DeleteName"
    
    Call IssueDeleteName

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
    strRoutine_nm = "name.cls.DeleteName"
    
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_obj_id, "N", "=")
    strSql = strSql & " and" & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
   
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


    
Private Sub ValidateName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "name.cls.ValidateCustNm"
    
    If p_nm = Empty Or IsNull(p_nm) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0500"
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


Private Sub IssueInsertName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "name.cls.IssueInsertName"
    
    p_nm_id = fGetId("name_id", o_error, o_trans_data)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    strSql = "insert into name ( "
    strSql = strSql & " nm_id"
    strSql = strSql & ",nm_type_cd"
    strSql = strSql & ",def_nm_fl"
    strSql = strSql & ",nm"
    strSql = strSql & ",obj_nm"
    strSql = strSql & ",obj_id"
    strSql = strSql & ",updated_assoc_id"
    strSql = strSql & ",updated_ts"
    strSql = strSql & ",entered_assoc_id"
    strSql = strSql & ",entered_ts"
    strSql = strSql & ")"
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_nm_id, "N")
    strSql = strSql & "," & fInsertVariable(p_nm_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_def_nm_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_nm, "S")
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


Private Sub IssueUpdateName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "name.cls.IssueUpdateName"
    
    strSql = "update name set "
    strSql = strSql & " " & fUpdateVariable("nm", p_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("def_nm_fl", p_def_nm_fl, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & "where "
    strSql = strSql & "   " & fWhereVariable("nm_id", p_nm_id, "N", "=")
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

Private Sub IssueDeleteName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "name.cls.IssueDeletename"
    
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("nm_id", p_nm_id, "N", "=")
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

