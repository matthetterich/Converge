VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalesComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const M_STR_CLASS_NM As String = "clsSalesComm"


Public Sub extract()

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "extract"
    
    Dim objConnection As ADODB.Connection, _
        strConnectStringTx As String
    
    Set objConnection = New ADODB.Connection
    
    strConnectStringTx = "FILEDSN=c:/gss/load/rebuild.dsn"

    Call objConnection.Open(strConnectStringTx)
    Call delete(objConnection)
    Call insertNonDropShipPurchasedItems(objConnection)
    Call insertNonDropShipAssembledItems(objConnection)
    Call insertDropShipPurchasedItems(objConnection)
    Call objConnection.Close
    Set objConnection = Nothing
    Exit Sub
    
errHandler:
    Set objConnection = Nothing
    Call Err.Raise(1)
End Sub


Private Sub delete(ByRef r_objConnection As ADODB.Connection)
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "delete"
    
    Dim strSqlTx As String
    
    strSqlTx = "Delete from dwh_sales_comm"
    Call r_objConnection.Execute(strSqlTx)
    Exit Sub
errHandler:
    Call Err.Raise(1)
End Sub

Private Sub insertNonDropShipPurchasedItems(ByRef r_objConnection As ADODB.Connection)
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipPurchasedItems"
          
    Dim objRecordset As ADODB.Recordset, _
        strSqlTx As String, _
        dblGrossMarginPct As Double, _
        dblCommPct As Double, _
        dblCommAm As Double, _
        dteCustRegDt As Date, _
        dteOrderedDt As Date, _
        dblItemPrice As Double, _
        dblItemCost As Double, _
        dblItemQty As Double

    strSqlTx = _
        "select " & _
        "  associate.assoc_id " & _
        " ,associate.assoc_Nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,purchase_order.po_id " & _
        " ,purchase_order_item.line_nbr as po_line_nbr " & _
        " ,purchase_order_item.item_cost as po_item_cost " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct " & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_cost as item_item_cost "
    
    strSqlTx = strSqlTx & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd " & _
        " ,lot.lot_id " & _
        " ,lot.unit_cost as lot_item_cost " & _
        " ,inventory_adj.adj_qty "
    
    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join inventory_adj on inventory_adj.ship_id = shipment_item.ship_id " & _
        "  and inventory_adj.so_line_nbr = shipment_item.line_nbr " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join sales_order on sales_order.sales_ord_id = shipment.sales_ord_id " & _
        "inner join sales_order_item on sales_order_item.sales_ord_id = sales_order.sales_ord_id " & _
        "  and sales_order_item.line_nbr = shipment_item.line_nbr " & _
        "inner join item on item.item_id = shipment_item.item_id " & _
        "inner join lot on lot.lot_id = inventory_adj.lot_id " & _
        "inner join purchase_order on purchase_order.po_id = lot.po_id " & _
        "inner join purchase_order_item on purchase_order_item.po_id = lot.po_id " & _
        "  and purchase_order_item.line_nbr = lot.po_line_nbr " & _
        "inner join customer on customer.cust_id = sales_order.cust_id " & _
        "inner join sales_invoice on sales_invoice.ship_id = shipment_item.ship_id " & _
        "inner join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
        "  and sales_invoice_item.line_nbr = shipment_item.line_nbr " & _
        "inner join associate on associate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' "
        
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        dteCustRegDt = objRecordset("registration_dt")
        dteOrderedDt = objRecordset("Ordered_dt")
        dblItemPrice = CDbl(objRecordset("so_item_price"))
        dblItemCost = CDbl(objRecordset("lot_item_cost"))
        dblItemQty = CDbl(objRecordset("adj_qty") * -1)
        dblGrossMarginPct = getGrossMarginPct(dblItemCost, dblItemPrice)
        dblCommPct = getCommPct(dblGrossMarginPcg, dteCustRegDt, dteOrderedDt)
        dblCommAm = dblItemQty * dblItemPrice * dblCommPct
        
        strSqlTx = _
            "insert into dwh_sales_comm (" & _
            "  sales_rep_assoc_id " & _
            " ,sales_rep_assoc_nbr " & _
            " ,sales_rep_assoc_nm " & _
            " ,cust_id " & _
            " ,cust_nbr " & _
            " ,cust_nm " & _
            " ,cust_registration_dt " & _
            " ,sales_ord_id " & _
            " ,sales_ord_ordered_dt " & _
            " ,sales_ord_line_nbr " & _
            " ,sales_ord_item_price " & _
            " ,po_id " & _
            " ,po_line_nbr " & _
            " ,po_item_cost " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd " & _
            " ,sales_inv_id " & _
            " ,sales_inv_discount_pct " & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price "
        strSqlTx = strSqlTx & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,lot_id " & _
            " ,lot_unit_cost " & _
            " ,assem_ord_id " & _
            " ,comm_item_price " & _
            " ,comm_item_cost " & _
            " ,comm_item_qty " & _
            " ,comm_margin_pct " & _
            " ,comm_pct " & _
            " ,comm_am " & _
            " ) "
        
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(objRecordset("assoc_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assoc_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("registration_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ordered_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_line_nbr"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_line_nbr"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_inv_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("discount_pct"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(dblItemPrice, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemCost, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemQty, "N") & _
            ", " & modUtilites.fInsertVariable(dblGrossMarginPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommAm, "N") & ")"
        
    
        Call r_objConnection.Execute(strSqlTx)
        Call objRecordset.MoveNext
    Loop
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    Call Err.Raise(1)
End Sub


Private Sub insertNonDropShipAssembledItems(ByRef r_objConnection As ADODB.Connection)
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipAssembledItems"
          
    Dim objRecordset As ADODB.Recordset, _
        strSqlTx As String, _
        dblGrossMarginPct As Double, _
        dblCommPct As Double, _
        dblCommAm As Double, _
        dteCustRegDt As Date, _
        dteOrderedDt As Date, _
        dblItemPrice As Double, _
        dblItemCost As Double, _
        dblItemQty As Double

    strSqlTx = _
        "select " & _
        "  associate.assoc_id " & _
        " ,associate.assoc_Nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,assembly_order.assem_ord_id " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct " & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_cost as item_item_cost "
    
    strSqlTx = strSqlTx & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.assem_cost " & _
        " ,item.commodity_cd " & _
        " ,lot.lot_id " & _
        " ,lot.unit_cost as lot_item_cost " & _
        " ,inventory_adj.adj_qty "
        
    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join inventory_adj on inventory_adj.ship_id = shipment_item.ship_id " & _
        "  and inventory_adj.so_line_nbr = shipment_item.line_nbr " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join sales_order on sales_order.sales_ord_id = shipment.sales_ord_id " & _
        "inner join sales_order_item on sales_order_item.sales_ord_id = sales_order.sales_ord_id " & _
        "  and sales_order_item.line_nbr = shipment_item.line_nbr " & _
        "inner join item on item.item_id = shipment_item.item_id " & _
        "inner join lot on lot.lot_id = inventory_adj.lot_id " & _
        "inner join customer on customer.cust_id = sales_order.cust_id " & _
        "inner join sales_invoice on sales_invoice.ship_id = shipment_item.ship_id " & _
        "inner join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
        "  and sales_invoice_item.line_nbr = shipment_item.line_nbr " & _
        "left join assembly_order on assembly_order.assem_ord_id = lot.assem_ord_id " & _
        "inner join associate on associate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' "
        
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        dteCustRegDt = objRecordset("registration_dt")
        dteOrderedDt = objRecordset("Ordered_dt")
        dblItemPrice = CDbl(objRecordset("so_item_price"))
        dblItemCost = CDbl(objRecordset("lot_item_cost"))
        dblItemQty = CDbl(objRecordset("adj_qty") * -1)
        dblGrossMarginPct = getGrossMarginPct(dblItemCost, dblItemPrice)
        dblCommPct = getCommPct(dblGrossMarginPcg, dteCustRegDt, dteOrderedDt)
        dblCommAm = dblItemQty * dblItemPrice * dblCommPct
        
        strSqlTx = _
            "insert into dwh_sales_comm (" & _
            "  sales_rep_assoc_id " & _
            " ,sales_rep_assoc_nbr " & _
            " ,sales_rep_assoc_nm " & _
            " ,cust_id " & _
            " ,cust_nbr " & _
            " ,cust_nm " & _
            " ,cust_registration_dt " & _
            " ,sales_ord_id " & _
            " ,sales_ord_ordered_dt " & _
            " ,sales_ord_line_nbr " & _
            " ,sales_ord_item_price " & _
            " ,po_id " & _
            " ,po_line_nbr " & _
            " ,po_item_cost " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd " & _
            " ,sales_inv_id " & _
            " ,sales_inv_discount_pct " & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price "
        strSqlTx = strSqlTx & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,lot_id " & _
            " ,lot_unit_cost " & _
            " ,assem_ord_id " & _
            " ,comm_item_price " & _
            " ,comm_item_cost " & _
            " ,comm_item_qty " & _
            " ,comm_margin_pct " & _
            " ,comm_pct " & _
            " ,comm_am " & _
            " ) "
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(objRecordset("assoc_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assoc_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("registration_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ordered_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_line_nbr"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_inv_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("discount_pct"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assem_cost"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assem_ord_id"), "S") & _
            ", " & modUtilites.fInsertVariable(dblItemPrice, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemCost, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemQty, "N") & _
            ", " & modUtilites.fInsertVariable(dblGrossMarginPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommAm, "N") & ")"
    
        Call r_objConnection.Execute(strSqlTx)
        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    Call Err.Raise(1)
End Sub

Private Sub insertDropShipPurchasedItems(ByRef r_objConnection As ADODB.Connection)
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertDropShipPurchasedItems"
          
    Dim objRecordset As ADODB.Recordset, _
        strSqlTx As String, _
        dblGrossMarginPct As Double, _
        dblCommPct As Double, _
        dblCommAm As Double, _
        dteCustRegDt As Date, _
        dteOrderedDt As Date, _
        dblItemPrice As Double, _
        dblItemCost As Double, _
        dblItemQty As Double

    strSqlTx = _
        "select " & _
        "  associate.assoc_id " & _
        " ,associate.assoc_Nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,purchase_order.po_id " & _
        " ,purchase_order_item.line_nbr as po_line_nbr " & _
        " ,purchase_order_item.item_cost as po_item_cost " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,shipment_item.loaded_qty " & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct " & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_cost as item_item_cost "
    
    strSqlTx = strSqlTx & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd "
    
    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join inventory_adj on inventory_adj.ship_id = shipment_item.ship_id " & _
        "  and inventory_adj.so_line_nbr = shipment_item.line_nbr " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join sales_order on sales_order.sales_ord_id = shipment.sales_ord_id " & _
        "inner join sales_order_item on sales_order_item.sales_ord_id = sales_order.sales_ord_id " & _
        "  and sales_order_item.line_nbr = shipment_item.line_nbr " & _
        "inner join item on item.item_id = shipment_item.item_id " & _
        "inner join lot on lot.lot_id = inventory_adj.lot_id " & _
        "inner join purchase_order on purchase_order.po_id = lot.po_id " & _
        "inner join purchase_order_item on purchase_order_item.po_id = lot.po_id " & _
        "  and purchase_order_item.line_nbr = lot.po_line_nbr " & _
        "inner join customer on customer.cust_id = sales_order.cust_id " & _
        "inner join sales_invoice on sales_invoice.ship_id = shipment_item.ship_id " & _
        "inner join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
        "  and sales_invoice_item.line_nbr = shipment_item.line_nbr " & _
        "inner join associate on associate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' "
        
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        dteCustRegDt = objRecordset("registration_dt")
        dteOrderedDt = objRecordset("Ordered_dt")
        dblItemPrice = CDbl(objRecordset("so_item_price"))
        dblItemCost = CDbl(objRecordset("po_item_cost"))
        dblItemQty = CDbl(objRecordset("loaded_qty"))
        dblGrossMarginPct = getGrossMarginPct(dblItemCost, dblItemPrice)
        dblCommPct = getCommPct(dblGrossMarginPcg, dteCustRegDt, dteOrderedDt)
        dblCommAm = dblItemQty * dblItemPrice * dblCommPct

        strSqlTx = _
            "insert into dwh_sales_comm (" & _
            "  sales_rep_assoc_id " & _
            " ,sales_rep_assoc_nbr " & _
            " ,sales_rep_assoc_nm " & _
            " ,cust_id " & _
            " ,cust_nbr " & _
            " ,cust_nm " & _
            " ,cust_registration_dt " & _
            " ,sales_ord_id " & _
            " ,sales_ord_ordered_dt " & _
            " ,sales_ord_line_nbr " & _
            " ,sales_ord_item_price " & _
            " ,po_id " & _
            " ,po_line_nbr " & _
            " ,po_item_cost " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd " & _
            " ,sales_inv_id " & _
            " ,sales_inv_discount_pct " & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price "
        strSqlTx = strSqlTx & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,lot_id " & _
            " ,lot_unit_cost " & _
            " ,assem_ord_id " & _
            " ,comm_item_price " & _
            " ,comm_item_cost " & _
            " ,comm_item_qty " & _
            " ,comm_margin_pct " & _
            " ,comm_pct " & _
            " ,comm_am " & _
            " ) "
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(objRecordset("assoc_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assoc_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("registration_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ordered_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_line_nbr"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("so_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_line_nbr"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("po_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_inv_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("discount_pct"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(dblItemPrice, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemCost, "N") & _
            ", " & modUtilites.fInsertVariable(dblItemQty, "N") & _
            ", " & modUtilites.fInsertVariable(dblGrossMarginPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommPct, "N") & _
            ", " & modUtilites.fInsertVariable(dblCommAm, "N") & ")"
    
        Call r_objConnection.Execute(strSqlTx)
        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    Call Err.Raise(1)
End Sub



Private Function getGrossMarginPct( _
                          ByVal v_dblItemCost As Double _
                        , ByVal v_dblItemPrice As Double _
                        ) As Double


    If r_dblItemCost > 0 Then
        getGrossMarginPct = (v_dblItemPrice - r_dblItemCost) / r_dblItemCost
    Else
        dblGrossProfit = 9999
    End If
    
End Function

Private Function getCommPct( _
              ByVal v_dblGrossMarginPct As Double _
            , ByVal v_dteCustRegistrationDt As Date _
            , ByVal v_dteOrderedDt As Date _
            )

    If DateDiff("yyyy", v_dteOrderedDt, v_dteCustRegistrationDt) <= 1 Then
        If v_dblGrossMarginPct >= 0.25 Then
            dblCommPct = 0.07
        ElseIf v_dblGrossMarginPct >= 0.22 Then
            dblCommPct = 0.05
        ElseIf v_dblGrossMarginPct >= 0.2 Then
            dblCommPct = 0.04
        ElseIf v_dblGrossMarginPct >= 0.15 Then
            dblCommPct = 0.02
        Else
            dblCommPct = 0.01
        End If
    ElseIf DateDiff("yyyy", v_dteOrderedDt, v_dteCustRegistrationDt) <= 2 Then
        If v_dblGrossMarginPct >= 0.25 Then
            dblCommPct = 0.07
        ElseIf v_dblGrossMarginPct >= 0.22 Then
            dblCommPct = 0.05
        ElseIf v_dblGrossMarginPct >= 0.2 Then
            dblCommPct = 0.04
        ElseIf v_dblGrossMarginPct >= 0.15 Then
            dblCommPct = 0.02
        Else
            dblCommPct = 0.01
        End If
    ElseIf DateDiff("yyyy", v_dteOrderedDt, v_dteCustRegistrationDt) <= 3 Then
        If v_dblGrossMarginPct >= 0.25 Then
            dblCommPct = 0.0525
        ElseIf v_dblGrossMarginPct >= 0.22 Then
            dblCommPct = 0.0375
        ElseIf v_dblGrossMarginPct >= 0.2 Then
            dblCommPct = 0.03
        ElseIf v_dblGrossMarginPct >= 0.15 Then
            dblCommPct = 0.01
        Else
            dblCommPct = 0
        End If
    Else
        If v_dblGrossMarginPct >= 0.25 Then
            dblCommPct = 0.035
        ElseIf v_dblGrossMarginPct >= 0.22 Then
            dblCommPct = 0.025
        ElseIf v_dblGrossMarginPct >= 0.2 Then
            dblCommPct = 0.02
        ElseIf v_dblGrossMarginPct >= 0.15 Then
            dblCommPct = 0.01
        Else
            dblCommPct = 0
        End If
    End If


End Function

