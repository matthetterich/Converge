VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "item_inspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_item_id As Variant
Private mvarp_inspect_crit As Variant
Private mvaro_record_set As adodb.Recordset
Private mvaro_record_set_input As adodb.Recordset
'

Public Property Set o_record_set_input(ByVal vData As adodb.Recordset)
    Set mvaro_record_set_input = vData
End Property
Public Property Get o_record_set_input() As adodb.Recordset
    Set o_record_set_input = mvaro_record_set_input
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
End Property
Public Property Let p_inspect_crit(ByVal vData As Variant)
    mvarp_inspect_crit = vData
End Property
Public Property Get p_inspect_crit() As Variant
        p_inspect_crit = mvarp_inspect_crit
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
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

Public Sub GetItemInspects()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "item_inspect_crit.cls.getItemInspects"
    
    strSql = "SELECT inspection.inspect_nm, item_inspection.item_id, inspection.inspect_desc, item_inspection.item_inspect_desc " & _
             " FROM inspection LEFT OUTER JOIN " & _
             " item_inspection ON " & _
             " inspection.inspect_nm = item_inspection.inspect_nm " & _
             " AND " & fWhereVariable("item_inspection.item_id", p_item_id, "N", "=") & _
             " WHERE inspection.inspect_nm IS NOT NULL "
             
    
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


Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "Item_inspect.cls.Change"
    
    'Call validate("change")
    Call IssueDelete
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


Private Sub IssueDelete()
    
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Item_inspect.cls.IssueDelete"
    
    strSql = "DELETE FROM item_inspection " & _
             " WHERE " & _
             " " & fWhereVariable("item_id", p_item_id, "N", "=")
             
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

Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Item_inspect.cls.IssueInsert"
    
    If o_record_set_input.EOF And o_record_set_input.BOF Then
    Else
        o_record_set_input.MoveFirst
    End If
        
    Do Until o_record_set_input.EOF
        
        strSql = "INSERT INTO item_inspection (" & _
                 " item_id, " & _
                 " inspect_nm, " & _
                 " item_inspect_desc " & _
                 " ) " & _
                 " VALUES (" & _
                 " " & fInsertVariable(p_item_id, "N") & "," & _
                 " " & fInsertVariable(o_record_set_input("inspect_nm"), "S") & "," & _
                 " " & fInsertVariable(o_record_set_input("item_inspect_desc"), "S") & _
                 ")"
        
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        o_record_set_input.MoveNext
                
    Loop
             
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
