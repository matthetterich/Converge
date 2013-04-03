VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "purchase_order_rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarp_po_rec_id As Variant
Private mvarp_po_id As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_supplier_id As Variant
Private mvaro_error As Object
Private mvaro_record_set As ADODB.Recordset
Private mvaro_trans_data As Object
Private mvaro_purchase_order As Object
Private mvarp_supplier_nbr As Variant
Private mvarp_receipt_dt As Variant
'

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

Public Property Set o_purchase_order(ByVal vData As Object)
    Set mvaro_purchase_order = vData
End Property
Public Property Get o_purchase_order() As Object
    Set o_purchase_order = mvaro_purchase_order
End Property

Public Property Let p_receipt_dt(ByVal vData As Variant)
    mvarp_receipt_dt = vData
End Property
Public Property Get p_receipt_dt() As Variant
        p_receipt_dt = mvarp_receipt_dt
End Property

Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Sets the error object up for reporting error encountered
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Supplier ID is one of the keys for purchase order
'               Returns
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
   End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Status description
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Status Code
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Purchase Order ID
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
        p_po_id = mvarp_po_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Purchase Order Receipt ID
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_po_rec_id(ByVal vData As Variant)
    mvarp_po_rec_id = vData
End Property
Public Property Get p_po_rec_id() As Variant
        p_po_rec_id = mvarp_po_rec_id
End Property
Public Sub GetPurchaseOrderRec()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "Purchase_Order_Rec.cls.GetPurchaseOrderRec"
    
    Call SelectPurchaseOrderRec
    
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
    strRoutine_nm = "Purchase_Order_Rec.cls.Add"
    
    Call getPurchaseOrder
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
    strRoutine_nm = "Purchase_Order_Rec.cls.Change"
    
    Call getPurchaseOrder
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
    strRoutine_nm = "Purchase_Order_Rec.cls.Delete"
    
    Call Validate("delete")
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
    strRoutine_nm = "Purchase_Order_Rec.cls.IssueDeleteData"
    
    strSql = "DELETE FROM purchase_order_rec " & _
             " WHERE " & fWhereVariable("po_receipt_id", p_po_rec_id, "N", "=")
        
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

Private Sub Validate(sFunction_cd)

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim objRecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "Purchase_Order_Rec.cls.Validate"
    
    If sFunction_cd = "add" Then
        If o_purchase_order.p_drop_ship_fl = "Y" Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "0100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1023
        End If
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        If fValuePresent(p_status_cd) = True And _
           fValuePresent(fGetDesc("sales_order_status_cd", p_status_cd, o_error, o_trans_data)) = False Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "0100"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 2010
            Exit Sub
        End If
    End If
        
    If sFunction_cd = "delete" Then
        strSql = "select inv_adj_id from inventory_adj " & _
                 " where " & fWhereVariable("po_receipt_id", p_po_rec_id, "N", "=")
        
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Set objRecordset = Nothing
            Exit Sub
        End If

        If objRecordset.EOF = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2148
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

Private Sub IssueInsertData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedQty As Double
    
    strRoutine_nm = "Purchase_Order_Rec.cls.IssueInsertData"
    p_po_rec_id = fGetId("po_receipt_id", o_error, o_trans_data)
    
    strSql = "insert into purchase_order_rec ( " & _
                     "  po_receipt_id" & _
                     ", po_id" & _
                     ", receipt_dt" & _
                     ", status_cd" & _
                     ", supplier_id" & _
                     ", updated_assoc_id" & _
                     ", updated_ts" & _
                     ", entered_assoc_id" & _
                     ", entered_ts" & _
                     " ) " & _
                     "values (" & _
                     " " & fInsertVariable(p_po_rec_id, "N") & _
                     "," & fInsertVariable(p_po_id, "N") & _
                     "," & fInsertVariable(p_receipt_dt, "D") & _
                     "," & fInsertVariable("OP", "S") & _
                     "," & fInsertVariable(p_supplier_id, "N") & _
                     "," & fInsertVariable(o_trans_data.p_assoc_id, "N") & _
                     "," & fInsertVariable("current_ts", "D") & _
                     "," & fInsertVariable(o_trans_data.p_assoc_id, "N") & _
                     "," & fInsertVariable("current_ts", "D") & _
                     ")"
  
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
    
    strSql = "UPDATE purchase_order_rec set " & _
             " " & fUpdateVariable("receipt_dt", p_receipt_dt, "D", "=") & _
             "," & fUpdateVariable("status_cd", p_status_cd, "S", "=") & _
             " WHERE " & fWhereVariable("po_receipt_id", p_po_rec_id, "N", "=")
                
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

Private Sub SelectPurchaseOrderRec()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
    
    strRoutine_nm = "Purchase_Order_Rec.cls.SelectPurchaseOrderRec"


    
    If IsNull(p_po_rec_id) = True Then
     p_po_rec_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM purchase_order_rec "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("po_receipt_id", p_po_rec_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    Else
        If objRecordset.EOF = True Then
           p_po_rec_id = Null
           ' Close the Record Set
           objRecordset.Close
           Set objRecordset = Nothing
        Else
            p_po_rec_id = objRecordset("po_receipt_id")
            p_po_id = objRecordset("po_id")
            p_receipt_dt = objRecordset("receipt_dt")
            p_supplier_id = objRecordset("supplier_id")
            p_status_cd = objRecordset("status_cd")
            
            ' Close the Record Set
            objRecordset.Close
            Set objRecordset = Nothing
            
            p_supplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
            p_status_desc = fGetDesc("po_receipt_status_cd", p_status_cd, o_error, o_trans_data)
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


Public Sub GetPurchaseOrderRecs()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_recs.cls.getPurchaseOrderRecs"
    Dim strSql As String
    Dim objRecordset As ADODB.Recordset
        
    strSql = "SELECT " & _
                " a.po_id, " & _
                " a.po_receipt_id, " & _
                " a.receipt_dt, " & _
                " b.supplier_nbr, " & _
                " c.code_desc " & _
                " FROM purchase_order_rec a, supplier b, code_desc c" & _
                " WHERE a.supplier_id = b.supplier_id " & _
                "   AND (" & fWhereVariable("c.field_nm", "po_receipt_status_cd", "S", "=") & _
                "   AND c.code = a.status_cd)"
               
    If IsNull(p_supplier_id) = True Or p_supplier_id = "" Then
    Else
     strSql = strSql & "  AND " & fWhereVariable("a.supplier_id", p_supplier_id, "N", "=")
    End If
    
    If IsNull(p_po_id) = True Or p_po_id = "" Then
    Else
     strSql = strSql & "  AND " & fWhereVariable("a.po_id", p_po_id, "N", "=")
    End If
        
    If IsNull(p_status_cd) = True Or p_status_cd = "" Then
    Else
     strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
        
    strSql = strSql & " Order BY a.po_receipt_id DESC "
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    Else
        Set o_record_set = objRecordset
    End If

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


Private Sub getPurchaseOrder()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "purchase_order_recs.cls.GetPurchaseOrder"

    If o_purchase_order Is Nothing Then
        Set o_purchase_order = CreateObject("converge_purchase.purchase_order")
        Set o_purchase_order.o_error = o_error
        Set o_purchase_order.o_trans_data = o_trans_data
        o_purchase_order.p_po_id = p_po_id
        o_purchase_order.getPurchaseOrder
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
