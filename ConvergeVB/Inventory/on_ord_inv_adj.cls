VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "on_ord_inv_adj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_trans_data As Object
Private mvaro_error As Object
Private mvarp_on_ord_inv_adj_id As Variant
Private mvarp_item_id As Variant
Private mvarp_po_id As Variant
Private mvarp_po_line_nbr As Variant
Private mvarp_assem_ord_id As Variant
Private mvarp_ac_line_nbr As Variant
Private mvarp_on_ord_qty As Variant
Private mvarp_inv_loc_id As Variant
Private mvaro_record_set As ADODB.Recordset

Public Property Let p_on_ord_inv_adj_id(ByVal vData As Variant)
    mvarp_on_ord_inv_adj_id = vData
End Property
Public Property Get p_on_ord_inv_adj_id() As Variant
    p_on_ord_inv_adj_id = mvarp_on_ord_inv_adj_id
End Property

Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
    p_inv_loc_id = mvarp_inv_loc_id
End Property

Public Property Let p_on_ord_qty(ByVal vData As Variant)
    mvarp_on_ord_qty = vData
End Property
Public Property Get p_on_ord_qty() As Variant
    p_on_ord_qty = mvarp_on_ord_qty
End Property

Public Property Let p_assem_ord_id(ByVal vData As Variant)
    mvarp_assem_ord_id = vData
End Property
Public Property Get p_assem_ord_id() As Variant
    p_assem_ord_id = mvarp_assem_ord_id
End Property
Public Property Let p_ac_line_nbr(ByVal vData As Variant)
    mvarp_ac_line_nbr = vData
End Property
Public Property Get p_ac_line_nbr() As Variant
    p_ac_line_nbr = mvarp_ac_line_nbr
End Property


Public Property Let p_po_line_nbr(ByVal vData As Variant)
    mvarp_po_line_nbr = vData
End Property
Public Property Get p_po_line_nbr() As Variant
    p_po_line_nbr = mvarp_po_line_nbr
End Property

Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
    p_po_id = mvarp_po_id
End Property

Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = mvarp_item_id
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

Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Get on order inventory adjustment
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetOnOrderInventoryAdj()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "on_ord_inv_adj.cls.GetOnOrderInventoryAdj"
    Dim objrecordset As ADODB.Recordset
  
    Call SelectOnOrdInvAdj(p_on_ord_inv_adj_id, objrecordset)
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
    Else
        If objrecordset.EOF = True Then
            p_on_ord_inv_adj_id = Null
        Else
            p_on_ord_inv_adj_id = objrecordset("on_ord_inv_adj_id")
            p_item_id = objrecordset("item_id")
            p_inv_loc_id = objrecordset("inv_loc_id")
            p_on_ord_qty = objrecordset("on_ord_qty")
            p_po_id = objrecordset("po_id")
            p_po_line_nbr = objrecordset("po_line_nbr")
            p_assem_ord_id = objrecordset("assem_ord_id")
        End If
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing

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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Add on order inventory adjustment
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "on_ord_inv_adj.cls.Add"
    Dim objInvLoc As Object
    Set objInvLoc = CreateObject("converge_inv.inventory_loc")
    
    Call Validate("add", objInvLoc)
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
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "on_ord_inv_adj.cls.change"
    Dim objInvLoc As Object
    Set objInvLoc = CreateObject("converge_inv.inventory_loc")
    
    Call IssueDeleteData(p_po_id, p_po_line_nbr, p_assem_ord_id, p_ac_line_nbr)
    
    If p_on_ord_qty > 0 Then
        Call Validate("add", objInvLoc)
        Call IssueInsertData
    End If
    
    Dim objInventoryUtil As Object
    Set objInventoryUtil = CreateObject("converge_inv.inventory_util")
    Set objInventoryUtil.o_trans_data = o_trans_data
    Set objInventoryUtil.o_error = o_error
    objInventoryUtil.p_item_id = p_item_id
    objInventoryUtil.ReallocateItemOnOrdInv
    Set objInventoryUtil = Nothing
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to get the available inventory
'  Arguments:   VarInv_loc_id
'               varItem_id
'               o_error
'
'//////////////////////////////////////////////////////////////////////////////
Public Function fAvailableQty(ByVal varInv_loc_id As Variant, ByVal varItem_id As Variant, ByRef r_objError As Object, ByRef r_objTrans_data As Object) As Double
    
    On Error GoTo error_handler
    
    Set o_error = r_objError
    Set o_trans_data = r_objTrans_data

    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "on_ord_inv_adj.cls.fAvailableQty"
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    Dim dblInventoryQty As Double
    Dim dblAllocatedQty As Double
    
    strSql = "SELECT sum(on_ord_qty) AS [p_on_ord_qty] " & _
            "FROM on_ord_inv_adj " & _
            "WHERE " & _
            "    " & fWhereVariable("item_id", varItem_id, "N", "=") & _
            " and " & fWhereVariable("inv_loc_id", varInv_loc_id, "S", "=")

    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_on_ord_qty")) = True Then
        dblInventoryQty = 0
    Else
        dblInventoryQty = objrecordset("p_on_ord_qty")
    End If
    
    objrecordset.Close
    
    strSql = "SELECT sum(alloc_qty) as [p_alloc_qty] " & _
            "FROM on_ord_inv_alloc " & _
            "WHERE " & _
            "    " & fWhereVariable("item_id", varItem_id, "N", "=") & _
            "and " & fWhereVariable("inv_loc_id", varInv_loc_id, "S", "=")

    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objrecordset("p_alloc_qty")) = True Then
        dblAllocatedQty = 0
    Else
        dblAllocatedQty = objrecordset(0)
    End If
    
    fAvailableQty = dblInventoryQty - dblAllocatedQty
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to delete on order inventory adjustment
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Delete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "on_ord_inv_adj.cls.Add"
    
    Call IssueDeleteData(p_po_id, p_po_line_nbr, p_assem_ord_id, p_ac_line_nbr)
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to insert into on order inventory adjustment
'  Arguments:   sFunction_cd
'               objInvLoc
'               o_error
'               o_trans_data
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Validate(ByVal sFunction_cd As Variant, ByRef objInvLoc As Object)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    
    ' Initialize
    strRoutine_nm = "on_ord_inv_adj.cls.validate"
    
    If sFunction_cd = "add" Then
        If p_on_ord_qty = 0 Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "2100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2217
            Exit Sub
        End If
        
        ' validate location
        Set objInvLoc.o_trans_data = o_trans_data
        Set objInvLoc.o_error = o_error
        objInvLoc.p_inv_loc_id = p_inv_loc_id
        objInvLoc.GetInventoryLoc
        
        If IsNull(objInvLoc.p_inv_loc_id) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "2300"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1175
            Exit Sub
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to insert into on order inventory adjustment
'  Arguments:   o_error
'               o_trans_data
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueInsertData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "on_ord_inv_adj.cls.IssueInsertData"
    
    If p_po_id > 0 Then
        p_assem_ord_id = Null
    Else
        p_po_id = Null
        p_po_line_nbr = Null
    End If
    
    p_on_ord_inv_adj_id = fGetId("on_ord_inv_adj_id", o_error, o_trans_data)
    
    strSql = "insert into on_ord_inv_adj ( "
    strSql = strSql & "  on_ord_inv_adj_id "
    strSql = strSql & ", item_id "
    strSql = strSql & ", inv_loc_id "
    strSql = strSql & ", on_ord_qty "
    strSql = strSql & ", po_id "
    strSql = strSql & ", po_line_nbr "
    strSql = strSql & ", assem_ord_id "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_on_ord_inv_adj_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_inv_loc_id, "S")
    strSql = strSql & "," & fInsertVariable(p_on_ord_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_po_id, "N")
    strSql = strSql & "," & fInsertVariable(p_po_line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_assem_ord_id, "N")
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Delete on order inventory adjustment
'  Arguments:   varOnOrdInvAdj
'               o_error
'               o_trans_data
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueDeleteData(ByVal v_intPoId As Variant, ByVal v_intPoLineNbr As Variant, ByVal v_intAssemOrdId As Variant, ByVal v_intAcLineNbr As Variant)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "on_ord_inv_adj.cls.IssueDeleteData"
    
    If v_intPoId > 0 Then
        strSql = " DELETE FROM on_ord_inv_adj " & _
                    "WHERE " & fWhereVariable("po_id", v_intPoId, "N", "=") & _
                    " and " & fWhereVariable("po_line_nbr", v_intPoLineNbr, "N", "=")
    Else
        strSql = " DELETE FROM on_ord_inv_adj " & _
                    "WHERE " & fWhereVariable("assem_ord_id", v_intAssemOrdId, "N", "=") & _
                    " "
    End If

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


'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to Delete on order inventory adjustment
'  Arguments:   varOnOrdInvAdjId
'               o_error
'               o_trans_data
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelectOnOrdInvAdj(ByVal intOnOrdInvAdjId As Double, ByRef objrecordset As ADODB.Recordset)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    strRoutine_nm = "on_ord_inv_adj.cls.SelectInventory"

    strSql = "SELECT * " & _
             "FROM on_ord_inv_adj " & _
            "WHERE " & fWhereVariable("on_ord_inv_adj_id", intOnOrdInvAdjId, "N", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
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


Public Sub GetOnOrderInventoryAdjs()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "on_ord_inv_allocs.cls.getOnOrderInventoryAdjs"
    
    strSql = "SELECT a.on_ord_inv_adj_id, a.on_ord_qty, a.inv_loc_id, a.po_id, "
        strSql = strSql & " a.po_line_nbr, a.assem_ord_id, b.item_nbr, b.item_desc,"
        strSql = strSql & " d.status_cd AS ao_status, d.due_dt,"
        strSql = strSql & " d.entered_dt AS ao_entered_dt, c.status_cd AS po_status,"
        strSql = strSql & " c.entered_dt AS po_entered_dt"
    strSql = strSql & " FROM on_ord_inv_adj a LEFT OUTER JOIN"
        strSql = strSql & " purchase_order c ON a.po_id = c.po_id LEFT OUTER JOIN"
        strSql = strSql & " assembly_order d ON"
        strSql = strSql & " a.assem_ord_id = d.assem_ord_id LEFT OUTER JOIN"
        strSql = strSql & " item b ON a.item_id = b.item_id "
    strSql = strSql & " WHERE 1 = 1 "
                
    If fValuePresent(p_item_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If fValuePresent(p_inv_loc_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
    If fValuePresent(p_po_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.po_id", p_po_id, "S", "=")
    End If
    
    If fValuePresent(p_assem_ord_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.assem_ord_id", p_assem_ord_id, "S", "=")
    End If
    
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



