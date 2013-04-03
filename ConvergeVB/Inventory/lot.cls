VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_varp_Lot_id As Variant
Private m_varp_item_id As Variant
Private m_varp_item_nbr As Variant
Private m_varp_supplier_lot_id As Variant
Private m_varp_status_cd As Variant
Private m_varp_po_receipt_id As Variant
Private m_varp_po_id As Variant
Private m_varp_po_line_nbr As Variant
Private m_varp_assem_ord_id As Variant
Private m_varp_prod_run_id As Variant
Private m_varp_phy_inv_id As Variant
Private m_varo_error As Object
Private m_varo_trans_data As Object
Private m_varp_unit_cost As Variant
Private m_varp_status_desc As Variant
Private m_varp_record_set_count As Variant
Private m_varo_record_set As ADODB.Recordset
Private m_varo_record_set_input As ADODB.Recordset
Private m_varo_record_set_comp_lot As ADODB.Recordset
'

Public Property Set o_record_set_comp_lot(ByVal vData As ADODB.Recordset)
    Set m_varo_record_set_comp_lot = vData
End Property
Public Property Get o_record_set_comp_lot() As ADODB.Recordset
        Set o_record_set_comp_lot = m_varo_record_set_comp_lot
End Property

Public Property Set o_record_set_input(ByVal vData As ADODB.Recordset)
    Set m_varo_record_set_input = vData
End Property
Public Property Get o_record_set_input() As ADODB.Recordset
        Set o_record_set_input = m_varo_record_set_input
End Property

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set m_varo_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
        Set o_record_set = m_varo_record_set
End Property
Public Property Let p_record_set_count(ByVal vData As Variant)
    m_varp_record_set_count = vData
End Property
Public Property Get p_record_set_count() As Variant
        p_record_set_count = m_varp_record_set_count
End Property

Public Property Let p_status_desc(ByVal vData As Variant)
    m_varp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = m_varp_status_desc
End Property
Public Property Let p_unit_cost(ByVal vData As Variant)
    m_varp_unit_cost = vData
End Property
Public Property Get p_unit_cost() As Variant
        p_unit_cost = m_varp_unit_cost
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set m_varo_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_varo_trans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_varo_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_varo_error
End Property
Public Property Let p_assem_ord_id(ByVal vData As Variant)
    m_varp_assem_ord_id = vData
End Property
Public Property Get p_assem_ord_id() As Variant
        p_assem_ord_id = m_varp_assem_ord_id
End Property
Public Property Let p_prod_run_id(ByVal vData As Variant)
    m_varp_prod_run_id = vData
End Property
Public Property Get p_prod_run_id() As Variant
        p_prod_run_id = m_varp_prod_run_id
End Property
Public Property Let p_phy_inv_id(ByVal vData As Variant)
    m_varp_phy_inv_id = vData
End Property
Public Property Get p_phy_inv_id() As Variant
        p_phy_inv_id = m_varp_phy_inv_id
End Property
Public Property Let p_po_line_nbr(ByVal vData As Variant)
    m_varp_po_line_nbr = vData
End Property
Public Property Get p_po_line_nbr() As Variant
        p_po_line_nbr = m_varp_po_line_nbr
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    m_varp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
        p_po_id = m_varp_po_id
End Property
Public Property Let p_po_receipt_id(ByVal vData As Variant)
    m_varp_po_receipt_id = vData
End Property
Public Property Get p_po_receipt_id() As Variant
        p_po_receipt_id = m_varp_po_receipt_id
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    m_varp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = m_varp_status_cd
End Property
Public Property Let p_supplier_lot_id(ByVal vData As Variant)
    m_varp_supplier_lot_id = vData
End Property
Public Property Get p_supplier_lot_id() As Variant
        p_supplier_lot_id = m_varp_supplier_lot_id
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    m_varp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = m_varp_item_id
End Property
Public Property Let p_item_nbr(ByVal vData As Variant)
    m_varp_item_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
        p_item_nbr = m_varp_item_nbr
End Property
Public Property Let p_lot_id(ByVal vData As Variant)
    m_varp_Lot_id = vData
End Property
Public Property Get p_lot_id() As Variant
        p_lot_id = m_varp_Lot_id
End Property

Public Sub GetLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.GetLot"
  
    If IsNumeric(p_lot_id) = True Then
    Else
        p_lot_id = Null
    End If
    
    Call SelectLot
    
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.Add"
    
    Dim dblAssemblyCost As Double
    
    Call Validate("add")
    
    If IsNumeric(Me.p_po_id) = True And IsNumeric(Me.p_po_line_nbr) = True Then
        Me.p_unit_cost = Me.getPOICost(Me.p_po_id, Me.p_po_line_nbr)
    Else
        Call Me.getItemCost(Me.p_item_id, Me.p_unit_cost, dblAssemblyCost)
    End If
    
    Call IssueInsertData
    
    If o_record_set_input Is Nothing Then
    Else
        If (o_record_set_input.BOF And o_record_set_input.EOF) Then
        Else
            Call ChangeLotInspects
        End If
    End If
    
    If o_record_set_comp_lot Is Nothing Then
    Else
        If (o_record_set_comp_lot.BOF And o_record_set_comp_lot.EOF) Then
        Else
            Call Me.AddCompLots
            Me.p_unit_cost = Me.getAssemCost(Me.p_item_id, Me.p_lot_id)
            Call UpdateCost
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


Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.Change"
    
    Call Validate("change")
    Call IssueUpdateData
    If o_record_set_input Is Nothing Then
    Else
        If o_record_set_input.BOF And o_record_set_input.EOF Then
        Else
            Call ChangeLotInspects
        End If
    End If
                                        'in case the lot cost has changed,
                                        'we need to update the cost on all parent lots
    If IsNumeric(Me.p_unit_cost) = True Then
        Call updateParentLots(o_error, o_trans_data, Me.p_lot_id)
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

Public Sub Delete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objLotInspect As Object
    strRoutine_nm = "Lot.cls.Delete"
    
    Call IssueDeleteData
    
    Set objLotInspect = CreateObject("converge_inv.lot_inspect")
    Set objLotInspect.o_error = o_error
    Set objLotInspect.o_trans_data = o_trans_data
    objLotInspect.p_lot_id = p_lot_id
    objLotInspect.Change
    
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Lot.cls.IssueDeleteData"

    strSql = "DELETE FROM lot " & _
             " WHERE " & fWhereVariable("lot_id", p_lot_id, "N", "=")
    
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim objSupplierLot As Object
    Dim strSql As String
    strRoutine_nm = "Lot.cls.Validate"

    If sFunction_cd = "add" Then
        If fValuePresent(p_item_nbr) = True Then
            p_item_id = fGetItemId(p_item_nbr, o_error, o_trans_data)
            If IsNull(p_item_id) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "1100"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 2115
            End If
        Else
            p_item_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
            If IsNull(p_item_nbr) = True Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "1120"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 2115
            End If
        End If
        
    End If
    
    If sFunction_cd = "add" Or sFunction_cd = "change" Then
        p_status_desc = fGetDesc("lot_status_cd", p_status_cd, o_error, o_trans_data)
        If fValuePresent(p_status_desc) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2214
        End If
    
        If fValuePresent(p_supplier_lot_id) Then
            Set objSupplierLot = CreateObject("converge_inv.supplier_lot")
            Set objSupplierLot.o_error = o_error
            Set objSupplierLot.o_trans_data = o_trans_data
            objSupplierLot.p_supplier_lot_id = p_supplier_lot_id
            objSupplierLot.GetSupplierLot
            If IsNull(objSupplierLot.p_supplier_lot_id) = True Then
                o_error.p_type_cd = "E"
                o_error.p_err_cd = "1100"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 2215
            End If
        End If
    End If
    
    If sFunction_cd = "delete" Then
        strSql = "select inv_adj_id from inventory_adj " & _
                 " where " & fWhereVariable("lot_id", p_lot_id, "N", "=")
        
        Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            objrecordset.Close
            Set objrecordset = Nothing
            Exit Sub
        End If

        If objrecordset.EOF = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 0
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Lot.cls.IssueInsertData"
    
    p_lot_id = fGetId("lot_id", o_error, o_trans_data)
    
    strSql = "insert into Lot ( "
    strSql = strSql & "  lot_id"
    strSql = strSql & ", item_id"
    strSql = strSql & ", supplier_lot_id"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", po_receipt_id"
    strSql = strSql & ", po_id"
    strSql = strSql & ", po_line_nbr"
    strSql = strSql & ", assem_ord_id"
    strSql = strSql & ", prod_run_id"
    strSql = strSql & ", phy_inv_id"
    strSql = strSql & ", unit_cost"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_lot_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_item_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_supplier_lot_id, "S")
    strSql = strSql & ", " & fInsertVariable(p_status_cd, "S")
    strSql = strSql & ", " & fInsertVariable(p_po_receipt_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_po_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_po_line_nbr, "N")
    strSql = strSql & ", " & fInsertVariable(p_assem_ord_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_prod_run_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_phy_inv_id, "N")
    strSql = strSql & ", " & fInsertVariable(p_unit_cost, "N")
    strSql = strSql & ", " & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & ", " & fInsertVariable("current_ts", "D")
    strSql = strSql & ", " & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & ", " & fInsertVariable("current_ts", "D")
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Lot.cls.IssueUpdateData"
    
    strSql = "update Lot set "
    strSql = strSql & " " & fUpdateVariable("status_cd", p_status_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("unit_cost", p_unit_cost, "N", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_lot_id", p_supplier_lot_id, "S", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("lot_id", p_lot_id, "N", "=")
    
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


Private Sub SelectLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "Lot.cls.SelectLot"
    
    strSql = "SELECT * " & _
             "from Lot " & _
            "WHERE " & fWhereVariable("lot_id", p_lot_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If
    

    If objrecordset.EOF = True Then
        p_lot_id = Null
    Else
        p_lot_id = objrecordset("lot_id")
        p_item_id = objrecordset("item_id")
        p_supplier_lot_id = objrecordset("supplier_lot_id")
        p_status_cd = objrecordset("status_cd")
        p_po_receipt_id = objrecordset("po_receipt_id")
        p_po_id = objrecordset("po_id")
        p_po_line_nbr = objrecordset("po_line_nbr")
        p_unit_cost = objrecordset("unit_cost")
        p_assem_ord_id = objrecordset("assem_ord_id")
        
        ' close the record set
        objrecordset.Close
        Set objrecordset = Nothing
        
        p_status_desc = fGetDesc("lot_status_cd", p_status_cd, o_error, o_trans_data)
        p_item_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
    
    End If
     
'    objRecordset.Close
    'Set objrecordset = Nothing
    
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


Public Sub GetLots()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    Dim lngCount As Long
    Dim strItemNbr As String
    
    
    strRoutine_nm = "Lot.cls.getLots"
    
    strSql = "select " & _
             " a.lot_id, " & _
             " b.item_nbr, " & _
             " a.status_cd, " & _
             " a.supplier_lot_id, " & _
             " a.po_receipt_id, " & _
             " a.po_id, " & _
             " a.assem_ord_id, " & _
             " a.prod_run_id, " & _
             " a.phy_inv_id " & _
             " FROM lot a, item b " & _
             " WHERE " & _
             " a.item_id = b.item_id "
    
    'If IsNull(p_item_id) = True Or p_item_id = "" Then
    'Else
    '   strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    'End If
    
    If IsNull(p_item_nbr) = True Or p_item_nbr = "" Then
    Else
       lngCount = InStr(1, p_item_nbr, "*", vbBinaryCompare)
       If lngCount > 0 Then
            strItemNbr = Left(p_item_nbr, lngCount - 1)
            strSql = strSql & " AND item_nbr LIKE '" & strItemNbr & "%'"
       Else
            strSql = strSql & "  AND " & fWhereVariable("b.item_nbr", p_item_nbr, "S", "=")
       End If
    End If
    
    If IsNull(p_po_id) = True Or p_po_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.po_id", p_po_id, "N", "=")
    End If
    
    If IsNumeric(Me.p_po_line_nbr) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.po_line_nbr", Me.p_po_line_nbr, "N", "=")
    End If
    
    If IsNull(p_po_receipt_id) = True Or p_po_receipt_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.po_receipt_id", p_po_receipt_id, "N", "=")
    End If
    
    If IsNull(p_assem_ord_id) = True Or p_assem_ord_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.assem_ord_id", p_assem_ord_id, "N", "=")
    End If
    
    If IsNull(p_prod_run_id) = True Or p_prod_run_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.prod_run_id", p_prod_run_id, "N", "=")
    End If
    
    If IsNull(p_phy_inv_id) = True Or p_phy_inv_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.phy_inv_id", p_phy_inv_id, "N", "=")
    End If
    
    If IsNull(p_supplier_lot_id) = True Or p_supplier_lot_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_lot_id", p_supplier_lot_id, "S", "=")
    End If
             
    strSql = strSql & " ORDER BY a.lot_id"
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

Public Sub GetInventoryItembyLots()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
   
    strRoutine_nm = "Lot.cls.getLots"
    strSql = "select * from lot " & _
                "WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

Public Sub GetInventoryItembyLotsbyBin()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "Lot.cls.GetInventoryItembyLotsbyBin"
    
    strSql = "select a.lot_id, " & _
                     " a.supplier_lot_id, " & _
                     " a.unit_cost, " & _
                     " a.status_cd " & _
                     " b.bin_id " & _
                     " from lot a, inventory_bin b " & _
                "WHERE " & fWhereVariable("item_id", p_item_id, "N", "=") & _
                "  AND a.lot_id = b.lot_id " & _
                " ORDER BY a.lot_id, b.bin_id "

    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

Public Sub GetLotInspects()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "Lot.cls.getLotInspects"
    
    strSql = "SELECT a.inspect_nm, a.inspect_desc, b.item_inspect_desc, c.inspect_comment " & _
             " FROM item_inspection b LEFT OUTER JOIN " & _
             " lot_inspect c ON b.inspect_nm = c.inspect_nm AND " & _
             " " & fWhereVariable("c.lot_id", p_lot_id, "N", "=") & " RIGHT OUTER JOIN " & _
             " inspection a ON b.inspect_nm = a.inspect_nm " & _
             " WHERE " & _
             " " & fWhereVariable("b.item_id", p_item_id, "N", "=")

    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

Public Sub ChangeLotInspects()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objLotInspect As Object
    strRoutine_nm = "Item_inspect.cls.Change"
    
    Set objLotInspect = CreateObject("converge_inv.lot_inspect")
    Set objLotInspect.o_record_set_input = o_record_set_input
    objLotInspect.p_lot_id = p_lot_id
    Set objLotInspect.o_error = o_error
    Set objLotInspect.o_trans_data = o_trans_data
           
    Call objLotInspect.Change
           
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

Public Sub GetCompLots()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "Lot.cls.getCompLots"
    
    strSql = "SELECT * " & _
             " FROM comp_lot " & _
             " WHERE " & _
             " " & fWhereVariable("lot_id", p_lot_id, "N", "=")

    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objrecordset
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

Public Sub AddCompLots()
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim objCompLot
    Dim strRoutine_nm As String
    
    strRoutine_nm = "Lot.cls.AddCompLots"
    Set objCompLot = CreateObject("converge_prod.comp_lot")
    Set objCompLot.o_record_set_comp_lot = o_record_set_comp_lot
    objCompLot.p_lot_id = p_lot_id
    Set objCompLot.o_error = o_error
    Set objCompLot.o_trans_data = o_trans_data
    
    Call objCompLot.Add
        
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


Public Sub GetLotInspectsForItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    
    strRoutine_nm = "Lot.cls.getLotInspectsForItem"
    
    strSql = " SELECT a.inspect_nm, a.inspect_desc, b.item_inspect_desc " & _
             " FROM item_inspection b RIGHT OUTER JOIN " & _
             " inspection a ON b.inspect_nm = a.inspect_nm " & _
             " WHERE " & _
             " " & fWhereVariable("b.item_id", p_item_id, "N", "=")
                         
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        'objRecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objrecordset
   
    'If isnull(o_record_set) then

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

Private Sub UpdateCost()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Lot.cls.UpdateCost"
    
    strSql = "update Lot set "
    strSql = strSql & " " & fUpdateVariable("unit_cost", p_unit_cost, "N", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("lot_id", p_lot_id, "N", "=")
    
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

Public Function getAssemCost( _
              ByVal v_dblItemId As Double _
            , ByVal v_dblLotId As Double _
            ) As Double
                    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "lot.cls.getAssemCost"
        
    Dim dblItemCost As Double
    Dim dblAssemblyCost As Double
    
    Call getItemCost(v_dblItemId, dblItemCost, dblAssemblyCost)
    
    Dim strSqlTx As String
    Dim objrecordset As ADODB.Recordset
    
    strSqlTx = _
        "Select SUM(d.item_qty * c.unit_cost) as unit_cost " & _
        "  from lot a " & _
        "  inner join comp_lot b " & _
        "  on a.lot_id = b.lot_id " & _
        "  inner join lot c " & _
        "  on c.lot_id = b.comp_lot_id " & _
        "  inner join bill_of_material d " & _
        "  on d.item_id = a.item_id " & _
        "  and d.comp_item_id = c.item_id " & _
        "  where " & fWhereVariable("a.lot_id", v_dblLotId, "N", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSqlTx, o_error)
        
    If objrecordset.EOF = False Then
        getAssemCost = objrecordset("unit_cost") + dblAssemblyCost
    Else
        getAssemCost = dblItemCost + dblAssemblyCost
    End If
    
    Exit Function
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
End Function


Public Function getPOICost(ByVal v_dblPoId As Double, ByVal v_dblLineNbr) As Double
     
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    Dim objPurchaseOrderItem As Object
    strRoutine_nm = "lot.cls.getPOICost"
    
    Set objPurchaseOrderItem = CreateObject("converge_purchase.purchase_order_item")
    Set objPurchaseOrderItem.o_error = o_error
    Set objPurchaseOrderItem.o_trans_data = o_trans_data
    objPurchaseOrderItem.p_line_nbr = v_dblLineNbr
    objPurchaseOrderItem.p_po_id = v_dblPoId
    
    objPurchaseOrderItem.GetPurchaseOrderItem
    
    getPOICost = objPurchaseOrderItem.p_item_cost
    
    Exit Function
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
End Function


Public Sub getItemCost(ByVal v_dblItemId As Double, ByRef r_dblItemCost As Double, ByRef r_dblAssemblyCost As Double)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "lot.cls.getItemCost"
    
    Dim strSql As String
    Dim objItem As Object
    
    ' Get the item cost informaiton
    Set objItem = CreateObject("converge_item.item")
    objItem.p_item_id = v_dblItemId
    Set objItem.o_error = o_error
    Set objItem.o_trans_data = o_trans_data
    Call objItem.GetItem
    r_dblAssemblyCost = objItem.p_assem_cost
    r_dblItemCost = objItem.p_item_cost
    Set objItem = Nothing
           
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


Private Sub updateParentLots( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblLotId As Double _
            )
    
    On Error GoTo error_handler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "lot.cls.updateParentLots"
    
    Dim objLot As lot
    Dim objCompLot As Object
    Dim rstParentLots As ADODB.Recordset
    Dim dblItemCost As Double
    Set objLot = New lot
    Set objCompLot = CreateObject("converge_prod.comp_lot")
    
    Call objCompLot.GetParentLots(r_objError, r_objTransData, rstParentLots, v_dblLotId)
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Do While rstParentLots.EOF = False And _
        r_objError.type_cd <> "E" And _
        r_objError.type_cd <> "F"
        Set objLot.o_error = r_objError
        Set objLot.o_trans_data = r_objTransData
        objLot.p_lot_id = rstParentLots("lot_id")
        Call objLot.GetLot
        dblItemCost = objLot.getAssemCost(objLot.p_item_id, objLot.p_lot_id)
        objLot.p_unit_cost = dblItemCost
        Call objLot.Change
        Call rstParentLots.MoveNext
    Loop
    
    Exit Sub
error_handler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub




