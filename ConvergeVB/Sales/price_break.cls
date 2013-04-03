VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "price_break"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarp_cust_id As Variant
Private mvarp_item_id As Variant
Private mvarp_line_nbr As Variant
Private mvarp_quantity As Variant
Private mvarp_item_price As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As Object
'

Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
        p_cust_id = mvarp_cust_id
End Property

Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property

Public Property Let p_line_nbr(ByVal vData As Variant)
    mvarp_line_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
        p_line_nbr = mvarp_line_nbr
End Property

Public Property Let p_quantity(ByVal vData As Variant)
    mvarp_quantity = vData
End Property
Public Property Get p_quantity() As Variant
        p_quantity = mvarp_quantity
End Property

Public Property Let p_item_price(ByVal vData As Variant)
    mvarp_item_price = vData
End Property
Public Property Get p_item_price() As Variant
        p_item_price = mvarp_item_price
End Property

Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property

Public Sub GetPriceBreak()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "price_break.cls.GetPriceBreak"
    
    Call SelectPriceBreak
    
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

Private Sub SelectPriceBreak()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
        
    strRoutine_nm = "price_break.cls.SelectPriceBreak"
    
    strSql = " SELECT * FROM price_break "
    strSql = strSql & " WHERE " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("item_id", p_item_id, "N", "=")
    If Len(p_line_nbr) > 0 Then
        strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordSet = Nothing
        Exit Sub
    End If
    
    If objRecordSet.EOF = True Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        '' TODO: add error message here
        'o_error.p_type_cd = "E"
        'o_error.p_err_cd = "0500"
        'o_error.p_routine_nm = strRoutine_nm
        'o_error.p_message_id = 2133
        Exit Sub
    End If
    
    p_cust_id = objRecordSet("cust_id")
    p_item_id = objRecordSet("item_id")
    p_line_nbr = objRecordSet("line_nbr")
    p_quantity = objRecordSet("quantity")
    p_item_price = objRecordSet("item_price")
    
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

Public Sub Change()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "price_break.cls.Change"
    
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
    Err.clear
End Sub

Private Sub Validate()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "price_break.cls.Validate"
    
    If Len(p_cust_id) > 0 Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3429
        Exit Sub
    End If
    
    If Len(p_item_id) > 0 Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0300"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3429
        Exit Sub
    End If
    
    If Len(p_line_nbr) > 0 Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3430
        Exit Sub
    End If
    
    If CDbl(p_quantity) > 0 Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0500"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3431
        Exit Sub
    End If
    
    If CDbl(p_item_price) > 0 Then
    Else
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0600"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3432
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
    Err.clear
End Sub

Private Sub IssueUpdate()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "price_break.cls.IssueUpdate"
    
    strSql = " UPDATE price_break "
    strSql = strSql & " SET "
    strSql = strSql & " " & fUpdateVariable("quantity", p_quantity, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_price", p_item_price, "N", "=")
    strSql = strSql & " WHERE " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("item_id", p_item_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
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
    Err.clear
End Sub

Public Sub Add()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
    
    strRoutine_nm = "price_break.cls.Add"
    
    strSql = " SELECT max(line_nbr) as max_line FROM price_break "
    strSql = strSql & " WHERE " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If Len(objRecordSet("max_line")) > 0 Then
        p_line_nbr = CDbl(objRecordSet("max_line")) + 1
    Else
        p_line_nbr = 1
    End If
    
    Call Validate
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
    Err.clear
End Sub

Private Sub IssueInsert()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "price_break.cls.IssueInsert"
    
    strSql = " INSERT INTO price_break "
    strSql = strSql & "( "
    strSql = strSql & "cust_id ,"
    strSql = strSql & "item_id,"
    strSql = strSql & "line_nbr,"
    strSql = strSql & "quantity,"
    strSql = strSql & "item_price,"
    strSql = strSql & "entered_ts,"
    strSql = strSql & "entered_assoc_id,"
    strSql = strSql & "updated_ts,"
    strSql = strSql & "updated_assoc_id"
    strSql = strSql & " ) "
    strSql = strSql & " VALUES ( "
    strSql = strSql & " " & fInsertVariable(p_cust_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_quantity, "N")
    strSql = strSql & "," & fInsertVariable(p_item_price, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & ")"
    
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
    Err.clear
End Sub

Public Sub Delete()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "price_break.cls.GetPriceBreak"
    
    strSql = " DELETE FROM price_break "
    strSql = strSql & " WHERE " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("item_id", p_item_id, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
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
    Err.clear
End Sub

Public Sub GetPriceBreaks()
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordSet As ADODB.Recordset
    
    strRoutine_nm = "price_break.cls.GetPriceBreaks"
    
    strSql = " SELECT * FROM price_break "
    strSql = strSql & " WHERE 1 = 1 "
    
    If Len(p_cust_id) > 0 Then
        strSql = strSql & " AND " & fWhereVariable("cust_id", p_cust_id, "N", "=")
    End If
    
    If Len(p_item_id) > 0 Then
        strSql = strSql & " AND " & fWhereVariable("item_id", p_item_id, "N", "=")
    End If
    
    strSql = strSql & " ORDER BY quantity"
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordSet.Close
        Set objRecordSet = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordSet
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
    Err.clear
End Sub

