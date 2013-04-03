VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "on_ord_inv_alloc"
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
Private mvarp_mat_req_id As Variant
Private mvarp_alloc_qty As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_alloc_fl As Variant ' N means that after this allocation, there is not more inventory available
Private mvarp_on_ord_inv_alloc_id As Variant
Private mvarp_unallocated_qty As Variant
Private mvaro_record_set As ADODB.Recordset
'
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property
'

Public Property Let p_unallocated_qty(ByVal vData As Variant)
    mvarp_unallocated_qty = vData
End Property
Public Property Get p_unallocated_qty() As Variant
        p_unallocated_qty = mvarp_unallocated_qty
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_on_ord_inv_alloc_id 'Let' and 'Get' are used to
'               set up the unique on order inventory allocation ID.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_on_ord_inv_alloc_id(ByVal vData As Variant)
    mvarp_on_ord_inv_alloc_id = vData
End Property
Public Property Get p_on_ord_inv_alloc_id() As Variant
        p_on_ord_inv_alloc_id = mvarp_on_ord_inv_alloc_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_alloc_fl 'Let' and 'Get' are used to tell if all of the
'               inventory has been allocated.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_alloc_fl(ByVal vData As Variant)
    mvarp_alloc_fl = vData
End Property
Public Property Get p_alloc_fl() As Variant
        p_alloc_fl = mvarp_alloc_fl
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_inv_loc_id 'Let' and 'Get' are used to keep the inventory
'               Location ID
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_inv_loc_id 'Let' and 'Get' are used to keep allocation
'               Quantity
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_alloc_qty(ByVal vData As Variant)
    mvarp_alloc_qty = vData
End Property
Public Property Get p_alloc_qty() As Variant
    p_alloc_qty = mvarp_alloc_qty
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_inv_loc_id 'Let' and 'Get' are used to keep Sales
'               Order Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_mat_req_id(ByVal vData As Variant)
    mvarp_mat_req_id = vData
End Property
Public Property Get p_mat_req_id() As Variant
    p_mat_req_id = mvarp_mat_req_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_inv_loc_id 'Let' and 'Get' are used to keep Item
'               Identifier
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = mvarp_item_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_inv_loc_id 'Set' and 'Get' are used to load the
'               Transaction Data Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_error 'Set' and 'Get' are used to load the
'               Error Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property


'//////////////////////////////////////////////////////////////////////////////
'Description:   The allocate method will allocate on order inventory
'               and will not allow them to overallocate inventory
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Allocate()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim ObjOnOrdInvAdj As Object
    Dim dblAvailableQty As Double
    Dim dblOnOrderAvailableQty As Double
    
    strRoutine_nm = "inventory_alloc.cls.Allocate"
    
    If p_alloc_qty < 0 Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "2100"
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_message_id = 2210
        Exit Sub
    End If
    
    Set ObjOnOrdInvAdj = CreateObject("converge_inv.on_ord_inv_adj")
    
    Call DeleteAllocation
    dblAvailableQty = ObjOnOrdInvAdj.fAvailableQty(p_inv_loc_id, p_item_id, o_error, o_trans_data)
        
    If dblAvailableQty <= 0 Then
        p_unallocated_qty = p_alloc_qty
        p_alloc_fl = "N"
    Else
        If p_alloc_qty <= dblAvailableQty Then
            Call InsertAllocation(p_alloc_qty)
            p_unallocated_qty = 0
        Else
            Call InsertAllocation(dblAvailableQty)
            p_unallocated_qty = p_alloc_qty - dblAvailableQty
            p_alloc_fl = "N"
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


Public Sub Deallocate()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
  
    strRoutine_nm = "on_ord_inv_alloc.cls.DeAllocate"
    
    Call DeleteAllocation
    
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
'Description:   Function to Allocate inventory and on order inventory
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub InsertAllocation(ByVal v_dblAllocQty As Double)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "on_ord_inv_alloc.cls.InsertAllocation"

    p_on_ord_inv_alloc_id = fGetId("on_ord_inv_alloc_id", o_error, o_trans_data)
    
    strSql = " INSERT INTO on_ord_inv_alloc ( " & _
                    "  on_ord_inv_alloc_id" & _
                    ", inv_loc_id" & _
                    ", item_id" & _
                    ", mat_req_id" & _
                    ", alloc_qty" & _
                    ", updated_assoc_id" & _
                    ", updated_ts" & _
                    ", entered_assoc_id" & _
                    ", entered_ts" & _
                    " ) " & _
                    "VALUES (" & _
                    "  " & fInsertVariable(p_on_ord_inv_alloc_id, "N") & _
                    ", " & fInsertVariable(p_inv_loc_id, "S") & _
                    ", " & fInsertVariable(p_item_id, "N") & _
                    ", " & fInsertVariable(p_mat_req_id, "N") & _
                    ", " & fInsertVariable(v_dblAllocQty, "N")
      strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N") & _
                    "," & fInsertVariable("current_ts", "D") & _
                    "," & fInsertVariable(o_trans_data.p_assoc_id, "N") & _
                    "," & fInsertVariable("current_ts", "D") & _
                    ")"
    
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
'Description:   Function to Allocate inventory and on order inventory
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub DeleteAllocation()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "on_ord_inv_alloc.cls.DeleteAllocation"

    strSql = "DELETE FROM on_ord_inv_alloc " & _
           " WHERE " & fWhereVariable("mat_req_id", p_mat_req_id, "N", "=")
    
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


Private Function fGetAllocatedQty() As Double

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "inventory_alloc.cls.fGetAllocatedQty"

    strSql = "Select alloc_qty FROM on_ord_inv_alloc" & _
           "  WHERE " & fWhereVariable("mat_req_id", p_mat_req_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Function
    End If

    If IsNull(objrecordset(0)) = True Then
        fGetAllocatedQty = 0
    Else
        fGetAllocatedQty = objrecordset(0)
    End If
     
    Set o_record_set = objrecordset

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
'Description:   Gets the inventory allocation class
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetOnOrderInventoryAllocs()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "on_ord_inv_allocs.cls.getOnOrderInventoryallocs"
    
    strSql = "SELECT " & _
                " a.on_ord_inv_alloc_id, a.alloc_qty, a.mat_req_id, " & _
                " a.inv_loc_id, b.assem_ord_id, b.sales_ord_id, c.item_nbr, " & _
                "  c.item_desc, customer.cust_nbr, b.required_dt" & _
               "  FROM customer INNER JOIN" & _
                "   sales_order ON" & _
                "   customer.cust_id = sales_order.cust_id RIGHT OUTER JOIN" & _
                "   on_ord_inv_alloc a INNER JOIN " & _
                "    mat_req b ON a.mat_req_id = b.mat_req_id INNER JOIN" & _
                "   item c ON b.item_id = c.item_id ON" & _
                "   sales_order.sales_ord_id = b.sales_ord_id" & _
                " WHERE  a.alloc_qty > 0 "
                
    If fValuePresent(p_item_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If fValuePresent(p_inv_loc_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
    If fValuePresent(p_mat_req_id) Then
       strSql = strSql & "  AND " & fWhereVariable("a.mat_req_id", p_mat_req_id, "N", "=")
    End If
    
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If

    Set o_record_set = objrecordset

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



