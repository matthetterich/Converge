VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lot_inspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_lot_id As Variant
Private mvarp_inspect_crit As Variant
Private mvarp_inspect_comment As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set_input As ADODB.Recordset
'

Public Property Set o_record_set_input(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set_input = vData
End Property
Public Property Get o_record_set_input() As ADODB.Recordset
        Set o_record_set_input = mvaro_record_set_input
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
Public Property Let p_inspect_comment(ByVal vData As Variant)
    mvarp_inspect_comment = vData
End Property
Public Property Get p_inspect_comment() As Variant
    p_inspect_comment = mvarp_inspect_comment
End Property
Public Property Let p_inspect_crit(ByVal vData As Variant)
    mvarp_inspect_crit = vData
End Property
Public Property Set p_inspect_crit(ByVal vData As Variant)
    Set mvarp_inspect_crit = vData
End Property
Public Property Get p_inspect_crit() As Variant
    p_inspect_crit = mvarp_inspect_crit
End Property
Public Property Let p_lot_id(ByVal vData As Variant)
    mvarp_lot_id = vData
End Property
Public Property Get p_lot_id() As Variant
    p_lot_id = mvarp_lot_id
End Property


Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "lot_inspect.cls.Change"
    
    'Call Validate("change")
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
    
    strSql = "DELETE FROM lot_inspect " & _
             " WHERE " & _
             " " & fWhereVariable("lot_id", p_lot_id, "N", "=")
             
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
    strRoutine_nm = "lot_inspect.cls.IssueInsert"
    
    If o_record_set_input Is Nothing Then
        Exit Sub
    End If
    
    If o_record_set_input.EOF And o_record_set_input.BOF Then
    Else
        o_record_set_input.MoveFirst
    End If
    
    Do Until o_record_set_input.EOF
        
        strSql = "INSERT INTO lot_inspect (" & _
                 " lot_id, " & _
                 " inspect_nm, " & _
                 " inspect_comment " & _
                 " ) " & _
                 " VALUES (" & _
                 " " & fInsertVariable(p_lot_id, "N") & "," & _
                 " " & fInsertVariable(o_record_set_input("inspect_nm"), "S") & "," & _
                 " " & fInsertVariable(o_record_set_input("inspect_comment"), "S") & _
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

