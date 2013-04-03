VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "purchase_order_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_po_return_id As Variant
Private mvarp_po_id As Variant
Private mvarp_return_auth_nbr As Variant
Private mvarp_return_dt As Variant
Private mvarp_reason_cd As Variant
Private mvarp_reason_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_supplier_id As Variant
Private mvarp_supplier_nbr As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As ADODB.Recordset
'

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
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
Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Set p_supplier_id(ByVal vData As Variant)
    Set mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_reason_desc(ByVal vData As Variant)
    mvarp_reason_desc = vData
End Property
Public Property Get p_reason_desc() As Variant
        p_reason_desc = mvarp_reason_desc
End Property
Public Property Let p_reason_cd(ByVal vData As Variant)
    mvarp_reason_cd = vData
End Property
Public Property Get p_reason_cd() As Variant
        p_reason_cd = mvarp_reason_cd
End Property
Public Property Let p_return_dt(ByVal vData As Variant)
    mvarp_return_dt = vData
End Property
Public Property Get p_return_dt() As Variant
        p_return_dt = mvarp_return_dt
End Property
Public Property Let p_return_auth_nbr(ByVal vData As Variant)
    mvarp_return_auth_nbr = vData
End Property
Public Property Get p_return_auth_nbr() As Variant
        p_return_auth_nbr = mvarp_return_auth_nbr
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
        p_po_id = mvarp_po_id
End Property
Public Property Let p_po_return_id(ByVal vData As Variant)
    mvarp_po_return_id = vData
End Property
Public Property Get p_po_return_id() As Variant
        p_po_return_id = mvarp_po_return_id
End Property
Public Sub GetPurchaseOrderRet()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_Order_Ret.cls.GetPurchaseOrderRet"
    
    Call SelectPurchaseOrderRet
    
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
    strRoutine_nm = "Purchase_Order_Ret.cls.Add"
    
    Call ValidateData("add")
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
  strRoutine_nm = "Purchase_Order_Ret.cls.Change"
    
    Call ValidateData("change")
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
  strRoutine_nm = "Purchase_Order_Ret.cls.Delete"
    
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
    Dim dblUpdatedQty As Double
  strRoutine_nm = "Purchase_Order_Ret.cls.IssueDeleteData"

    strSql = "delete from Purchase_Order_Ret " & _
             " where " & fWhereVariable("po_return_id", p_po_return_id, "N", "=")
    
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
    Err.Clear
End Sub

Private Sub ValidateAdd()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    
    strRoutine_nm = "Purchase_Order_Ret.cls.ValidateAdd"
    
    If fValuePresent(p_return_dt) = False Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "0200"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2040
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

Private Sub ValidateData(sFunction_cd)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
  strRoutine_nm = "Purchase_Order_Ret.cls.ValidateData"
    
    If fValuePresent(p_status_cd) = True And _
       fGetDesc("po_return_status_cd", p_status_cd, o_error, o_trans_data) = Empty Then
      o_error.p_type_cd = "F"
      o_error.p_err_cd = "0300"
      o_error.p_routine_nm = strRoutine_nm
      o_error.p_message_id = 2010
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
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "Purchase_Order_Ret.cls.IssueInsertData"
    
    p_po_return_id = fGetId("po_return_id", o_error, o_trans_data)
    
    strSql = "insert into purchase_order_return ( "
    strSql = strSql & "  po_return_id"
    strSql = strSql & ", po_id"
    strSql = strSql & ", return_auth_nbr"
    strSql = strSql & ", return_dt"
    strSql = strSql & ", reason_cd"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", supplier_id"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_po_return_id, "N")
    strSql = strSql & "," & fInsertVariable(p_po_id, "N")
    strSql = strSql & "," & fInsertVariable(p_return_auth_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_return_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_reason_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
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
    Err.Clear
End Sub

Private Sub IssueUpdateData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    strRoutine_nm = "PurchaseOrderRec.cls.IssueUpdateData"
      
    strSql = "UPDATE pruchase_order_return set "
    strSql = strSql & " " & fUpdateVariable("return_dt", p_return_dt, "D", "=")
    strSql = strSql & " " & fUpdateVariable("return_auth_nbr", p_return_auth_nbr, "S", "=")
    strSql = strSql & " " & fUpdateVariable("reason_cd", p_reason_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("status_cd", p_status_cd, "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("po_return_id", p_po_return_id, "N", "=")
      
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
    Err.Clear
End Sub

Private Sub SelectPurchaseOrderRet()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strRoutine_nm = "Purchase_Order_Ret.cls.SelectPurchaseOrderRet"
    
    If IsNull(p_po_return_id) = True Then
     p_po_return_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM purchase_order_return "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("po_return_id", p_po_return_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        'objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    Else
        If objRecordset.EOF = True Then
            objRecordset.Close
            Set objRecordset = Nothing
            p_po_return_id = Null
        Else
            p_po_return_id = objRecordset("po_return_id")
            p_return_auth_nbr = objRecordset("return_auth_nbr")
            p_po_id = objRecordset("po_id")
            p_return_dt = objRecordset("return_dt")
            p_supplier_id = objRecordset("supplier_id")
            p_status_cd = objRecordset("status_cd")
            p_reason_cd = objRecordset("reason_cd")
            
            objRecordset.Close
            Set objRecordset = Nothing
                               
            p_supplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
            p_status_desc = fGetDesc("po_return_status_cd", p_status_cd, o_error, o_trans_data)
            p_reason_desc = fGetDesc("po_return_reason_cd", p_reason_cd, o_error, o_trans_data)
        End If
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


Public Sub GetPurchaseOrderRets()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    
    strRoutine_nm = "purchase_order_rets.cls.getPurchaseOrderRets"
    strSql = "SELECT " & _
                " a.po_id, " & _
                " a.po_return_id, " & _
                " b.supplier_nbr, " & _
                " a.return_auth_nbr, " & _
                " a.return_dt, a.status_cd " & _
                " FROM purchase_order_return a, supplier b " & _
                " WHERE a.supplier_id = b.supplier_id "
                
    If IsNull(p_supplier_id) = True Or p_supplier_id = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("a.supplier_id", p_supplier_id, "N", "=")
    End If
    
    If IsNull(p_po_id) = True Or p_po_id = "" Then
    Else
        strSql = strSql & "  AND " & fWhereVariable("a.po_id", p_po_id, "N", "=")
    End If
        
    strSql = strSql & " Order BY a.supplier_id, a.po_id "
    
    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        o_record_set.Close
        Set o_record_set = Nothing
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



