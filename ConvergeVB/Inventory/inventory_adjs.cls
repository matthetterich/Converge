VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inventory_adjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As ADODB.Recordset
Private mvarp_lot_id As Variant
Private mvarp_po_receipt_id As Variant
Private mvarp_po_line_nbr As Variant
Private mvarp_item_id As Variant
Private mvarp_inv_loc_id As Variant
Private mvarp_phy_inv_id As Variant
Private mvarp_prod_run_id As Variant
Private mvarp_bin_id As Variant
Private mvarp_count As Variant
'

Public Property Let p_count(ByVal vData As Variant)
    mvarp_count = vData
End Property
Public Property Get p_count() As Variant
        p_count = mvarp_count
End Property
Public Property Let p_prod_run_id(ByVal vData As Variant)
    mvarp_prod_run_id = vData
End Property
Public Property Get p_prod_run_id() As Variant
        p_prod_run_id = mvarp_prod_run_id
End Property
Public Property Let p_phy_inv_id(ByVal vData As Variant)
    mvarp_phy_inv_id = vData
End Property
Public Property Get p_phy_inv_id() As Variant
        p_phy_inv_id = mvarp_phy_inv_id
End Property
Public Property Let p_po_line_nbr(ByVal vData As Variant)
    mvarp_po_line_nbr = vData
End Property
Public Property Get p_po_line_nbr() As Variant
        p_po_line_nbr = mvarp_po_line_nbr
End Property

Public Property Let p_bin_id(ByVal vData As Variant)
    mvarp_bin_id = vData
End Property
Public Property Get p_bin_id() As Variant
        p_bin_id = mvarp_bin_id
End Property
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property
Public Property Let p_po_receipt_id(ByVal vData As Variant)
    mvarp_po_receipt_id = vData
End Property
Public Property Get p_po_receipt_id() As Variant
        p_po_receipt_id = mvarp_po_receipt_id
End Property
Public Property Let p_lot_id(ByVal vData As Variant)
    mvarp_lot_id = vData
End Property
Public Property Get p_lot_id() As Variant
        p_lot_id = mvarp_lot_id
End Property
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
Public Sub GetInventoryAdjs()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "inventory_adjs.cls.getInventoryAdjs"
    
    strSql = "SELECT " & _
                " sum(adj_qty) as [p_adj_qty], " & _
                " a.bin_id, " & _
                " a.inv_loc_id, " & _
                " a.lot_id, " & _
                " c.unit_cost, " & _
                " b.item_nbr, " & _
                " b.item_desc" & _
                " FROM inventory_adj a, item b, lot c " & _
                " WHERE a.item_id = b.item_id AND a.lot_id = c.lot_id  "
                
    If IsNull(p_item_id) = True Or p_item_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If IsNull(p_lot_id) = True Or p_lot_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.lot_id", p_lot_id, "N", "=")
    End If
    
    If IsNull(p_inv_loc_id) = True Or p_inv_loc_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
    If IsNull(p_bin_id) = True Or p_bin_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.bin_id", p_bin_id, "S", "=")
    End If
    
    strSql = strSql & " GROUP BY a.inv_loc_id, a.bin_id, a.lot_id, c.unit_cost, b.item_nbr, b.item_desc "
    
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

Public Sub GetPORecInventoryAdjs()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "inventory_adjs.cls.getPORecInventoryAdjs"
    
    strSql = "SELECT inv_loc_id, bin_id, lot_id, sum(adj_qty) as sum_adj_qty from inventory_adj " & _
                "WHERE " & fWhereVariable("po_line_nbr", p_po_line_nbr, "N", "=") & _
                "  AND " & fWhereVariable("po_receipt_id", p_po_receipt_id, "N", "=") & _
                "group by inv_loc_id, bin_id, lot_id "

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


Public Sub GetPhyInvAdjs()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "inventory_adjs.cls.getPhyInvAdjs"
    
    strSql = "SELECT " & _
                " a.inv_adj_id, " & _
                " a.lot_id, " & _
                " a.inv_loc_id, " & _
                " a.bin_id, " & _
                " a.adj_qty, " & _
                " a.entered_assoc_id, " & _
                " Upper(c.assoc_nbr) as upper_assoc_nbr, " & _
                " b.item_nbr, " & _
                " b.item_desc " & _
                " FROM inventory_adj a, item b, associate c " & _
                "WHERE " & fWhereVariable("phy_inv_id", p_phy_inv_id, "N", "=") & _
                " AND a.item_id = b.item_id AND a.entered_assoc_id = c.assoc_id"
                
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


Public Sub Chart()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "inventory_adjs.cls.getPhyInvAdjs"
    
    strSql = "SELECT " & _
                " a.adj_qty, " & _
                " b.item_nbr " & _
                " FROM inventory_adj a, item b " & _
                "WHERE a.item_id = b.item_id "
                
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If

    Set o_record_set = objrecordset
    p_count = 0
    
    Do Until o_record_set.EOF = True
        p_count = p_count + 1
        o_record_set.MoveNext
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


Public Sub GetTotalInventoryAdjs()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    strRoutine_nm = "inventory_adjs.cls.getTotalInventoryAdjs"
    
    strSql = "SELECT " & _
                " a.inv_adj_id, " & _
                " a.bin_id, " & _
                " a.inv_loc_id, " & _
                " a.lot_id, " & _
                " a.adj_qty, " & _
                " a.ship_id, " & _
                " a.ship_ret_id, " & _
                " a.sales_ord_id, " & _
                " a.so_line_nbr, " & _
                " a.po_receipt_id, " & _
                " a.po_return_id, " & _
                " a.po_id, " & _
                " a.po_line_nbr, " & _
                " a.assem_ord_id, " & _
                " a.ac_line_nbr, " & _
                " a.phy_inv_id, " & _
                " a.prod_run_id, " & _
                " a.mat_req_id, " & _
                " b.item_nbr, " & _
                " b.item_desc" & _
                " FROM inventory_adj a, item b " & _
                " WHERE a.item_id = b.item_id "
                
    If IsNull(p_item_id) = True Or p_item_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If IsNull(p_inv_loc_id) = True Or p_inv_loc_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
    strSql = strSql & " ORDER BY a.inv_adj_id "
    
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


Public Sub GetProdRunInventoryAdjs()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Exit Sub
    End If
        
    Dim strRoutine_nm As String
    Dim objrecordset As ADODB.Recordset
    Dim strSql As String
    strRoutine_nm = "inventory_adjs.cls.getProdRunInventoryAdjs"
    
    strSql = "SELECT inv_loc_id, bin_id, lot_id, sum(adj_qty) as sum_adj_qty from inventory_adj " & _
                "WHERE " & fWhereVariable("prod_run_id", p_prod_run_id, "N", "=") & " And ac_line_nbr Is Null " & _
                "group by inv_loc_id, bin_id, lot_id "

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


