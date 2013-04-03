VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_bin_id As Variant
Private mvarp_bin_desc As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_inv_loc_id As Variant
Private mvaro_record_set As adodb.Recordset
'


Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
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
Public Property Let p_bin_id(ByVal vData As Variant)
    mvarp_bin_id = vData
End Property
Public Property Get p_bin_id() As Variant
        p_bin_id = mvarp_bin_id
End Property
Public Property Let p_bin_desc(ByVal vData As Variant)
    mvarp_bin_desc = vData
End Property
Public Property Get p_bin_desc() As Variant
        p_bin_desc = mvarp_bin_desc
End Property

Public Sub GetInventoryBin()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.GetInventorybin"
  
    Call SelectInventorybin
    
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
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.Add"
    
 
    Call Validate("add")
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


Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.Change"
    
    Call Validate("change")
    Call IssueUpdateData
    
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
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.Delete"
    
    Call IssueDeleteData
    
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


Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "bin.cls.IssueDeleteData"

    strSql = "delete from inventory_bin "
    strSql = strSql & "where " & fWhereVariable("bin_id", p_bin_id, "S", "=")
    
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

Private Sub Validate(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.Validate"

    If sFunction_cd = "add" Then
        If fValuePresent(p_bin_id) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2202
        End If
        Call ValidateID
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

Private Sub ValidateID()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "bin.cls.ValidateID"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    If IsNull(p_bin_desc) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1900"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2211
        Exit Sub
    End If
    
    strSql = "SELECT * " & _
             "FROM inventory_bin " & _
            "WHERE " & fWhereVariable("bin_id", p_bin_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "2100"
        o_error.p_message_id = 2191
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_desc = "The Bin that you entered already exists.  Please " & _
                            "click the BACK button and correct the problem."
        o_error.p_sqlstate_cd = ""
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

Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "bin.cls.IssueInsertData"
    
    strSql = "insert into inventory_bin ( "
    strSql = strSql & "  bin_id "
    strSql = strSql & ", bin_desc "
    strSql = strSql & ", inv_loc_id "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_bin_id, "S")
    strSql = strSql & "," & fInsertVariable(p_bin_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_id, "S")
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

Private Sub IssueUpdateData()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "bin.cls.IssueUpdateData"
    
    strSql = "update inventory_bin set "
    strSql = strSql & " " & fUpdateVariable("inv_loc_id", p_inv_loc_id, "S", "=")
    strSql = strSql & "," & fUpdateVariable("bin_desc", p_bin_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("bin_id", p_bin_id, "S", "=")
    
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


Private Sub SelectInventorybin()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "bin.cls.SelectInventorybin"
    
    strSql = "SELECT * " & _
             "FROM inventory_bin " & _
            "WHERE " & fWhereVariable("bin_id", p_bin_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_bin_id = Null
    Else
        p_bin_id = objRecordset("bin_id")
        p_bin_desc = objRecordset("bin_desc")
        p_inv_loc_id = objRecordset("inv_loc_id")
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


Public Sub GetInventoryBins()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "inventory_bin.cls.getInventoryBins"
    
    strSql = "select * from inventory_bin "

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
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


