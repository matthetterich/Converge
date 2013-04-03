VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comp_lot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_lot_id As Variant
Private mvarp_comp_lot_id As Variant
Private mvaro_record_set_comp_lot As ADODB.RecordSet
'

Public Property Set o_record_set_comp_lot(ByVal vdata As Object)
    Set mvaro_record_set_comp_lot = vdata
End Property
Public Property Get o_record_set_comp_lot() As Object
    Set o_record_set_comp_lot = mvaro_record_set_comp_lot
End Property

Public Property Let p_comp_lot_id(ByVal vdata As Variant)
    mvarp_comp_lot_id = vdata
End Property
Public Property Get p_comp_lot_id() As Variant
    p_comp_lot_id = mvarp_comp_lot_id
End Property
Public Property Let p_lot_id(ByVal vdata As Variant)
    mvarp_lot_id = vdata
End Property
Public Property Get p_lot_id() As Variant
    p_lot_id = mvarp_lot_id
End Property
Public Property Set o_trans_data(ByVal vdata As Object)
    Set mvaro_trans_data = vdata
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
Public Property Set o_error(ByVal vdata As Object)
    Set mvaro_error = vdata
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property


Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comp_lot.cls.Add"
    
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
    
    strSql = "DELETE FROM comp_lot " & _
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
    Dim objLot As Object
    strRoutine_nm = "comp_lot.cls.IssueInsert"
    Set objLot = CreateObject("converge_inv.lot")
    
    If o_record_set_comp_lot Is Nothing Then
        Exit Sub
    End If
    
    If o_record_set_comp_lot.EOF And o_record_set_comp_lot.BOF Then
    Else
        o_record_set_comp_lot.MoveFirst
    End If
    
    Do Until o_record_set_comp_lot.EOF
    
        objLot.p_lot_id = o_record_set_comp_lot("comp_lot")
        Set objLot.o_error = o_error
        Set objLot.o_trans_data = o_trans_data
        Call objLot.GetLot
        
        If objLot.p_item_id = o_record_set_comp_lot("item_id") Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "0100"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 3305
            End With
            Exit Sub
        End If
        
        strSql = "INSERT INTO comp_lot (" & _
                 " lot_id, " & _
                 " comp_lot_id " & _
                 " ) " & _
                 " VALUES (" & _
                 " " & fInsertVariable(p_lot_id, "N") & "," & _
                 " " & fInsertVariable(o_record_set_comp_lot("comp_lot"), "N") & _
                 ")"
        
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        o_record_set_comp_lot.MoveNext
                
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


Public Sub GetParentLots( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByRef r_objRecordset As ADODB.RecordSet _
            , ByVal v_dblLotId As Double _
            )
    
    On Error GoTo error_handler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "comp_lot.cls.getParentLots"
        
    Dim strSql As String
    
    strSql = "SELECT * " & _
             " FROM comp_lot " & _
             " WHERE " & _
             " " & fWhereVariable("comp_lot_id", v_dblLotId, "N", "=")

    Call r_objTransData.OpenRecordset(r_objRecordset, strSql, r_objError)
    
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





