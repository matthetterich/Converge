VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "shipment_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"shipment"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_objError As Object
Private m_objTrans_data As Object
Private m_objShipment As shipment
Private m_objSalesOrderItem As sales_order_item
Private m_objOldShipmentItem As Object
Private m_varShip_id As Variant
Private m_varItem_id As Variant
Private m_varItem_nbr As Variant
Private m_varSales_ord_id As Variant
Private m_varLine_nbr As Variant
Private m_varLoaded_qty As Variant
Private m_varOutstanding_qty As Variant
Private m_varPlanned_qty As Variant
Private m_varDrop_ship_fl As Variant
Private m_varShip_type_cd As Variant
Private m_varShip_type_Desc As Variant
Private mvaro_record_set As adodb.Recordset
'

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Let item_nbr(ByVal vData As Variant)
    m_varItem_nbr = vData
End Property
Public Property Get item_nbr() As Variant
    item_nbr = m_varItem_nbr
End Property
Public Property Let item_id(ByVal vData As Variant)
    m_varItem_id = vData
End Property
Public Property Get item_id() As Variant
    item_id = m_varItem_id
End Property
Public Property Let loaded_qty(ByVal vData As Variant)
    m_varLoaded_qty = vData
End Property
Public Property Get loaded_qty() As Variant
        loaded_qty = m_varLoaded_qty
End Property
Public Property Let outstanding_qty(ByVal vData As Variant)
    m_varOutstanding_qty = vData
End Property
Public Property Get outstanding_qty() As Variant
        outstanding_qty = m_varOutstanding_qty
End Property
Public Property Let planned_qty(ByVal vData As Variant)
    m_varPlanned_qty = vData
End Property
Public Property Get planned_qty() As Variant
        planned_qty = m_varPlanned_qty
End Property
Public Property Let ship_id(ByVal vData As Variant)
    m_varShip_id = vData
End Property
Public Property Get ship_id() As Variant
    ship_id = m_varShip_id
End Property
Public Property Let sales_ord_id(ByVal vData As Variant)
    m_varSales_ord_id = vData
End Property
Public Property Get sales_ord_id() As Variant
    sales_ord_id = m_varSales_ord_id
End Property
Public Property Let line_nbr(ByVal vData As Variant)
    m_varLine_nbr = vData
End Property
Public Property Get line_nbr() As Variant
    line_nbr = m_varLine_nbr
End Property
Public Property Let drop_ship_fl(ByVal vData As Variant)
    m_varDrop_ship_fl = vData
End Property
Public Property Get drop_ship_fl() As Variant
    drop_ship_fl = m_varDrop_ship_fl
End Property
Public Property Let ship_type_cd(ByVal vData As Variant)
    m_varShip_type_cd = vData
End Property
Public Property Get ship_type_cd() As Variant
    ship_type_cd = m_varShip_type_cd
End Property
Public Property Let ship_type_desc(ByVal vData As Variant)
    m_varShip_type_Desc = vData
End Property
Public Property Get ship_type_desc() As Variant
    ship_type_desc = m_varShip_type_Desc
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
Public Property Set o_old_shipment_item(ByVal vData As Object)
    Set m_objOldShipmentItem = vData
End Property
Public Property Get o_old_shipment_item() As Object
    Set o_old_shipment_item = m_objOldShipmentItem
End Property
Public Property Let p_item_nbr(ByVal vData As Variant)
    m_varItem_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
    p_item_nbr = m_varItem_nbr
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    m_varItem_id = vData
End Property
Public Property Get p_item_id() As Variant
    p_item_id = m_varItem_id
End Property
Public Property Let p_loaded_qty(ByVal vData As Variant)
    m_varLoaded_qty = vData
End Property
Public Property Get p_loaded_qty() As Variant
    p_loaded_qty = m_varLoaded_qty
End Property
Public Property Let p_outstanding_qty(ByVal vData As Variant)
    m_varOutstanding_qty = vData
End Property
Public Property Get p_outstanding_qty() As Variant
    p_outstanding_qty = m_varOutstanding_qty
End Property
Public Property Let p_planned_qty(ByVal vData As Variant)
    m_varPlanned_qty = vData
End Property
Public Property Get p_planned_qty() As Variant
    p_planned_qty = m_varPlanned_qty
End Property
Public Property Let p_ship_id(ByVal vData As Variant)
    m_varShip_id = vData
End Property
Public Property Get p_ship_id() As Variant
    p_ship_id = m_varShip_id
End Property
Public Property Let p_sales_ord_id(ByVal vData As Variant)
    m_varSales_ord_id = vData
End Property
Public Property Get p_sales_ord_id() As Variant
    p_sales_ord_id = m_varSales_ord_id
End Property
Public Property Let p_line_nbr(ByVal vData As Variant)
    m_varLine_nbr = vData
End Property
Public Property Get p_line_nbr() As Variant
    p_line_nbr = m_varLine_nbr
End Property
Public Property Let p_drop_ship_fl(ByVal vData As Variant)
    m_varDrop_ship_fl = vData
End Property
Public Property Get p_drop_ship_fl() As Variant
    p_drop_ship_fl = m_varDrop_ship_fl
End Property
Public Property Let p_ship_type_cd(ByVal vData As Variant)
    m_varShip_type_cd = vData
End Property
Public Property Get p_ship_type_cd() As Variant
    p_ship_type_cd = m_varShip_type_cd
End Property
Public Property Let p_ship_type_desc(ByVal vData As Variant)
    m_varShip_type_Desc = vData
End Property
Public Property Get p_ship_type_desc() As Variant
    p_ship_type_desc = m_varShip_type_Desc
End Property

Public Property Set SalesOrderItem(ByVal vData As Object)
    Set m_objSalesOrderItem = vData
End Property
Public Property Get SalesOrderItem() As Object
    Set SalesOrderItem = m_objSalesOrderItem
End Property
Public Property Set o_SalesOrderItem(ByVal vData As sales_order_item)
    Set m_objSalesOrderItem = vData
End Property
Public Property Get o_SalesOrderItem() As sales_order_item
    Set o_SalesOrderItem = m_objSalesOrderItem
End Property
Public Property Set shipment(ByVal vData As Object)
    Set m_objShipment = vData
End Property
Public Property Get shipment() As Object
    Set shipment = m_objShipment
End Property
Public Property Set o_shipment(ByVal vData As shipment)
    Set m_objShipment = vData
End Property
Public Property Get o_shipment() As shipment
    Set o_shipment = m_objShipment
End Property

Public Sub GetShipmentItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.GetshipmentItem"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    Set o_SalesOrderItem = New sales_order_item
    
    Call SelectShipmentItem(objRecordset)
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
    Else
        If objRecordset.EOF = True Then
            objRecordset.Close
            Set objRecordset = Nothing
            Call clear
        Else
            line_nbr = objRecordset("line_nbr")
            sales_ord_id = objRecordset("sales_ord_id")
            ship_id = objRecordset("ship_id")
            item_id = objRecordset("item_id")
            drop_ship_fl = objRecordset("drop_ship_fl")
            ship_type_cd = objRecordset("ship_type_cd")
            loaded_qty = objRecordset("loaded_qty")
            planned_qty = objRecordset("planned_qty")
            outstanding_qty = planned_qty - loaded_qty
            
            objRecordset.Close
            Set objRecordset = Nothing
            
            item_nbr = fGetItemNbr(item_id, o_error, o_trans_data)
            ship_type_desc = fGetDesc("shipment_type_cd", ship_type_cd, o_error, o_trans_data)
            Set o_SalesOrderItem.o_trans_data = o_trans_data
            Set o_SalesOrderItem.o_error = o_error
            o_SalesOrderItem.p_sales_ord_id = sales_ord_id
            o_SalesOrderItem.p_line_nbr = line_nbr
            Call o_SalesOrderItem.GetSalesOrderLineItem
        End If
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Public Sub Add()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.Add"
    
    Call validate("add")
    Call IssueInsert
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Public Sub Update()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.Update"
    
    Set Me.o_shipment = New shipment
    Set Me.o_shipment.o_error = Me.o_error
    Set Me.o_shipment.o_trans_data = Me.o_trans_data
    Me.o_shipment.p_ship_id = ship_id
    Me.o_shipment.p_bypass_desc_fl = "Y"
    Me.o_shipment.GetShipment
    
    Call validate("change")
    Call IssueUpdate
    
    Set Me.o_SalesOrderItem = New sales_order_item
    Call Me.o_SalesOrderItem.UpdateQuantities( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.sales_ord_id _
                            , Me.line_nbr _
                            )
    
    If Me.o_shipment.p_status_cd = "OP" Or _
       Me.o_shipment.p_status_cd = "IP" Then
       Me.o_shipment.p_status_cd = "OP"
       Call Me.o_shipment.changeStatus
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.Delete"
    
    If o_shipment Is Nothing Then
        Set o_shipment = New shipment
        Set o_shipment.o_error = o_error
        Set o_shipment.o_trans_data = o_trans_data
        o_shipment.p_ship_id = ship_id
        o_shipment.p_sales_ord_id = p_sales_ord_id
        o_shipment.p_bypass_desc_fl = "Y"
        o_shipment.GetShipment
    End If
    
    Call validate("delete")
    Call IssueDeleteData
    
    Set o_SalesOrderItem = New sales_order_item
    Call o_SalesOrderItem.UpdateQuantities( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.sales_ord_id _
                            , Me.line_nbr _
                            )
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Public Sub UpdateQuantities(ByVal v_varInv_loc_consignment_fl As Variant)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.SetLoadedQty"
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    Set o_shipment = New shipment
    Set o_shipment.o_error = o_error
    Set o_shipment.o_trans_data = o_trans_data
    o_shipment.p_ship_id = Me.ship_id
    o_shipment.p_bypass_desc_fl = "Y"
    o_shipment.GetShipment
    
                    ' note: 1st leg shipments create 2 inventory adjustments 1 negative adjustment from the warehouse,
                    ' one positive adjustment into the consignment warehouse.  If this is a 1st Leg shipment and the inventory location consignment flag = "Y" then
                    ' the inventory adjustment is an adjustment into a consigment location.  Do not update shipment quantities
                    ' for adjustments into the consignment location on 1st leg shipments.
    If v_varInv_loc_consignment_fl = "Y" Then
        If o_shipment.p_type_cd = "1L" Then
            Exit Sub
        End If
    End If
    
    strSql = "update shipment_item set " & _
                " " & fUpdateVariable("loaded_qty", getLoadedQty(), "N", "=") & _
                " where" & _
                "   " & fWhereVariable("ship_id", p_ship_id, "N", "=") & _
                "and" & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=") & _
                "and" & fWhereVariable("line_nbr", p_line_nbr, "N", "=")

    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If

    Set o_SalesOrderItem = New sales_order_item
    Call o_SalesOrderItem.UpdateQuantities( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.sales_ord_id _
                            , Me.line_nbr _
                            )
    
    If o_shipment.p_status_cd = "OP" Or _
       o_shipment.p_status_cd = "IP" Then
       Call o_shipment.setStatus( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.ship_id _
                            )
    End If
    
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub validate(ByVal sFunction_cd As Variant)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.Validate"
    
    If sFunction_cd = "change" Or sFunction_cd = "add" Then
        If drop_ship_fl = "Y" Or drop_ship_fl = "N" Then
        Else
            o_error.type_cd = "F"
            o_error.err_cd = "1200"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1275
        End If
    End If
    
    If sFunction_cd = "change" Then
        If drop_ship_fl = "Y" And loaded_qty < 0 Then
            o_error.type_cd = "F"
            o_error.err_cd = "1400"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1290
        End If
        
        
        
    End If
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Public Sub clear()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.Clear"
    p_line_nbr = Null
    p_sales_ord_id = Null
    p_ship_id = Null
    p_loaded_qty = Empty
    p_planned_qty = Empty
    p_outstanding_qty = Empty
    p_drop_ship_fl = Empty

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "1800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub IssueUpdate()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "shipment_item.cls.IssueUpdate"
    
    If o_shipment.p_drop_ship_fl = "N" Then
        loaded_qty = Empty
    End If
    
    strSql = "update shipment_item set "
    strSql = strSql & " " & fUpdateVariable("planned_qty", planned_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("loaded_qty", loaded_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("ship_id", ship_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("sales_ord_id", sales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", line_nbr, "N", "=")

    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "shipment_item.cls.IssueInsert"
    
    If IsNumeric(p_planned_qty) = False Then
        p_planned_qty = 0
    End If
        
    strSql = "insert into shipment_item ( "
    strSql = strSql & "  ship_id "
    strSql = strSql & ", sales_ord_id "
    strSql = strSql & ", line_nbr "
    strSql = strSql & ", item_id "
    strSql = strSql & ", drop_ship_fl "
    strSql = strSql & ", ship_type_cd "
    strSql = strSql & ", loaded_qty "
    strSql = strSql & ", planned_qty "
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(ship_id, "N")
    strSql = strSql & "," & fInsertVariable(sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(line_nbr, "N")
    strSql = strSql & "," & fInsertVariable(item_id, "N")
    strSql = strSql & "," & fInsertVariable(drop_ship_fl, "S")
    strSql = strSql & "," & fInsertVariable(ship_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(0, "N")
    strSql = strSql & "," & fInsertVariable(p_planned_qty, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "Shipment_item.cls.IssueDeleteData"

    strSql = "delete from Shipment_item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Function getLoadedQty()
    
    Dim objShipment As Object
        
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.getLoadedQty"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    Set objShipment = CreateObject("converge_sales.shipment")
    
    objShipment.p_ship_id = p_ship_id
    Set objShipment.o_error = o_error
    Set objShipment.o_trans_data = o_trans_data
    
    Call objShipment.GetShipment
        
    strSql = "SELECT sum(adj_qty) "
    strSql = strSql & "FROM inventory_adj "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("so_line_nbr", p_line_nbr, "N", "=")
    If LCase(objShipment.p_type_cd) = LCase("1L") Then
        strSql = strSql & "and (inv_loc_id = (SELECT inv_loc_id from sales_order where sales_ord_id = " & p_sales_ord_id & "))"
    '    strSql = strSql & " and " & fWhereVariable("inv_loc_id", "timco", "S", "=")
    End If
    
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If IsNull(objRecordset(0)) = True Then
        getLoadedQty = 0
    Else
        getLoadedQty = objRecordset(0) * -1
    End If
    
    objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Function
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Function


Private Sub SelectShipmentItem(ByRef r_objRecordset As adodb.Recordset)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment_item.cls.SelectShipmentItem"
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    strSql = "SELECT *      "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", p_line_nbr, "N", "=")
    
    Call o_trans_data.OpenRecordset(r_objRecordset, strSql, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "2800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Public Sub getShipmentItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.getShipmentItems"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
     
    strSql = "SELECT  " & _
                "  ship_id " & _
                " ,sales_ord_id " & _
                " ,line_nbr " & _
                " FROM shipment_item " & _
                " WHERE " & _
                "       " & fWhereVariable("ship_id", p_ship_id, "N", "=") & _
                "   AND " & fWhereVariable("sales_ord_id", p_sales_ord_id, "N", "=")
                
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
      .p_err_cd = "3000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

