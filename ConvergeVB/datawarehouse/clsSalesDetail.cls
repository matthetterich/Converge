VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalesDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const M_STR_CLASS_NM As String = "clsSalesDetail"

Public Sub extract( _
              ByRef r_objError As Variant _
            , ByRef r_conDwh As adodb.Connection _
            , ByRef r_conConverge As adodb.Connection _
            , ByVal v_strUnloadFileNm As String _
            )

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "extract"
    
    Dim lngFileId As Long _
        , strSqlTx As String
                                        'Clean up previous warehouse table
    Call dropTable(r_objError, r_conDwh)
    
                                        'Create New DWH Table
    Call createTable(r_objError, r_conDwh)
    
    lngFileId = FreeFile
    Open v_strUnloadFileNm For Output As #lngFileId
    
                                        'Insert shipment records
    Call insertNonDropShipPurchasedItems(r_objError, r_conConverge, lngFileId)
                                        'Insert shipment records
    Call insertNonDropShipAssembledItems(r_objError, r_conConverge, lngFileId)
                                        'Insert shipment records
    Call insertDropShipPurchasedItems(r_objError, r_conConverge, lngFileId)
    
    Close #lngFileId
    
    strSqlTx = _
      "BULK INSERT dwh_sales_detail " _
    & "From '" & v_strUnloadFileNm & "' " _
    & "WITH ( " _
    & "  DATAFILETYPE = 'char' " _
    & ", FIELDTERMINATOR = '<col>' " _
    & ", ROWTERMINATOR = '<row>\n' " _
    & ", TABLOCK " _
    & ", KEEPNULLS " _
    & ") "
    
    strSqlTx = _
      "BULK INSERT dwh_sales_detail " _
    & "From 'c:\data shares\dwhextract\CONVERGE_DWH_LOAD.TXT' " _
    & "WITH ( " _
    & "  DATAFILETYPE = 'char' " _
    & ", FIELDTERMINATOR = '<col>' " _
    & ", ROWTERMINATOR = '<row>\n' " _
    & ", TABLOCK " _
    & ", KEEPNULLS " _
    & ") "
    

    r_conDwh.CommandTimeout = 7200
    Call r_conDwh.Execute(strSqlTx)
    Kill v_strUnloadFileNm
    Exit Sub
    
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
End Sub


Private Sub dropTable( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "delete"
    
    Dim strSqlTx As String
    
                                        'drop the existing table
    strSqlTx = "drop table [dbo].[dwh_sales_detail]"
    On Error Resume Next
    Call r_objConn.Execute(strSqlTx)
    On Error GoTo 0
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub

Private Sub createTable( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "createTable"
    
    Dim strSqlTx As String
    
                                        'create a new table
    strSqlTx = _
        "CREATE TABLE [dbo].[dwh_sales_detail] ( " & _
        "[sales_rep_assoc_id] [int] NOT NULL , " & _
        "[sales_rep_assoc_nbr] [nvarchar] (50) NOT NULL , " & _
        "[sales_rep_assoc_nm] [nvarchar] (200)  NULL , " & _
        "[service_rep_assoc_id] [int] NOT NULL , " & _
        "[service_rep_assoc_nbr] [nvarchar] (50) NOT NULL , " & _
        "[service_rep_assoc_nm] [nvarchar] (200)  NULL , " & _
        "[cust_id] [int] NOT NULL , " & _
        "[cust_nbr] [nvarchar] (50)  NOT NULL , " & _
        "[cust_nm] [nvarchar] (200)  NULL , " & _
        "[cust_registration_dt] [datetime] NULL , " & _
        "[cust_grp_id] [int] NULL , " & _
        "[cust_grp_nbr] [nvarchar] (50) NULL , " & _
        "[cust_grp_nm] [nvarchar] (200) NULL , " & _
        "[cust_item_nbr] [nvarchar](200) NULL , " & _
        "[cust_item_price] [decimal](20, 5) NULL , " & _
        "[cust_item_yrs] [numeric](10, 4) null , " & _
        "[sales_ord_id] [int] NOT NULL , " & _
        "[sales_ord_ordered_dt] [datetime] NOT NULL , " & _
        "[sales_ord_line_nbr] [int] NOT NULL , " & _
        "[sales_ord_item_price] [decimal](20, 5) NOT NULL , " & _
        "[po_id] [int] NULL , " & _
        "[po_line_nbr] [int] NULL , "
    strSqlTx = strSqlTx & _
        "[po_item_cost] [decimal](20, 5) NULL , " & _
        "[ship_id] [int] NOT NULL , " & _
        "[ship_shipped_dt] [datetime] NOT NULL , " & _
        "[ship_drop_ship_fl] [char] (1)  NOT NULL , " & _
        "[ship_type_cd] [nvarchar] (50)  NOT NULL , " & _
        "[ship_planned_ship_dt] [datetime] NULL , " & _
        "[ship_planned_del_dt] [datetime] NULL , " & _
        "[sales_inv_id] [int] NOT NULL , " & _
        "[sales_inv_discount_pct] [numeric](18, 4) NULL , " & _
        "[item_id] [int] NOT NULL , "
    strSqlTx = strSqlTx & _
        "[item_nbr] [nvarchar] (50)  NOT NULL , " & _
        "[item_price] [decimal](20, 5) NOT NULL , " & _
        "[item_cost] [decimal](20, 5) NOT NULL , " & _
        "[item_assem_cost] [decimal](20, 5) NULL , " & _
        "[item_type_cd] [nvarchar] (50)  NOT NULL , " & _
        "[item_commodity_cd] [nvarchar] (50)  NULL , " & _
        "[item_desc] [nvarchar] (200)  NULL , " & _
        "[item_initial_purchase_dt] [datetime] NOT NULL , " & _
        "[lot_id] [int] NULL , " & _
        "[lot_unit_cost] [numeric](18, 5) NULL , " & _
        "[inv_loc_id] [nvarchar] (100) NOT NULL,  " & _
        "[assem_ord_id] [int] NULL , " & _
        "[sales_mo] [int] NULL , "
    strSqlTx = strSqlTx & _
        "[sales_dt] [datetime] NOT NULL , " & _
        "[sales_item_price] [decimal](20, 5) NOT NULL , " & _
        "[sales_item_cost] [decimal](20, 5) NOT NULL ," & _
        "[sales_item_qty] [numeric](18, 0) NOT NULL , " & _
        "[sales_MarkUp_pct] [numeric](18, 4) NOT NULL , " & _
        "[sales_comm_pct] [numeric](18, 4) NOT NULL , " & _
        "[sales_comm_am] [decimal](20, 5) NOT NULL , " & _
        "[sales_price] [decimal](20, 2) NOT NULL , " & _
        "[sales_cost] [decimal](20, 2) NOT NULL , " & _
        "[sales_profit_am] [decimal](20, 2) NOT NULL  "
'    strSqlTx = strSqlTx & _
'        "[lot_mat_cost] [decimal](20, 5) NOT NULL  " & _
'        "[lot_labor_cost] [decimal](20, 5) NOT NULL  " & _
'        "[lot_labor_hour_ct] [decimal](20, 5) NOT NULL  " & _
'        "[sales_mat_cost] [decimal](20, 2) NOT NULL  " & _
'        "[sales_labor_cost] [decimal](20, 2) NOT NULL  " & _
'        "[sales_labor_hour_ct] [decimal](20, 5) NOT NULL  "
    strSqlTx = strSqlTx & _
        ") ON [primary]"
    Call r_objConn.Execute(strSqlTx)
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub

Private Sub insertNonDropShipPurchasedItems( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            , ByVal v_lngFileId As Long _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipPurchasedItems"
          
    Dim objRecordset As adodb.Recordset, _
        strSqlTx As String, _
        dblSalesMarkUpPct As Double, _
        dblCommMarkUpPct As Double, _
        dblSalesCommPct As Double, _
        dblSalesCommAm As Double, _
        dteCustRegDt As Date, _
        dteInitialPurchaseDt As Date, _
        dteSalesDt As Date, _
        dblSalesItemPrice As Double, _
        dblSalesItemCost As Double, _
        dblOrigSalesItemCost As Double, _
        dblSalesItemQty As Double, _
        dblSalesPrice As Double, _
        dblSalesCost As Double, _
        dblSalesProfitAm As Double, _
        dblCustItemYrs As Double, _
        dblLotUnitCost As Double, _
        dblItemCost As Double, _
        dblItemPrice As Double, _
        dblSalesMo As Double, _
        lngItemId As Double, _
        strLoadRecord As String

                                        'create the sql statement
    strSqlTx = _
        "select " & _
        "  salesRepAssociate.assoc_id as sales_rep_assoc_id " & _
        " ,salesRepAssociate.assoc_nbr as sales_rep_assoc_nbr " & _
        " ,serviceRepAssociate.assoc_id as service_rep_assoc_id " & _
        " ,serviceRepAssociate.assoc_nbr as service_rep_assoc_nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,cust_item.initial_purchase_dt " & _
        " ,cust_item.cust_item_nbr " & _
        " ,cust_item.cust_item_price " & _
        " ,cust_group.cust_grp_id " & _
        " ,cust_group.cust_grp_nbr " & _
        " ,cust_group.cust_grp_nm "
    strSqlTx = strSqlTx & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,purchase_order.po_id " & _
        " ,purchase_order.supplier_id " & _
        " ,purchase_order_item.line_nbr as po_line_nbr " & _
        " ,purchase_order_item.item_cost as po_item_cost " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,sales_invoice_item.item_price as si_item_price " & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct "
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd " & _
        " ,lot.lot_id " & _
        " ,lot.unit_cost as lot_unit_cost " & _
        " ,inventory_adj.adj_qty " & _
        " ,inventory_adj.inv_loc_id " & _
        " ,an.nm as sales_rep_assoc_nm " & _
        " ,sn.nm as service_rep_assoc_nm " & _
        " ,cn.nm as cust_nm " & _
        " ,jls_inv_loc_id "

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
        "left join purchase_order on purchase_order.po_id = lot.po_id " & _
        "left join purchase_order_item on purchase_order_item.po_id = lot.po_id " & _
        "  and purchase_order_item.line_nbr = lot.po_line_nbr " & _
        "inner join customer on customer.cust_id = sales_order.cust_id " & _
        "inner join sales_invoice on sales_invoice.ship_id = shipment_item.ship_id " & _
        "inner join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
        "  and sales_invoice_item.line_nbr = shipment_item.line_nbr " & _
        "left join associate salesRepAssociate on salesRepAssociate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "left join associate serviceRepAssociate on serviceRepAssociate.assoc_id = customer.cust_rep_assoc_id  "
    strSqlTx = strSqlTx & _
        "left join cust_item on cust_item.cust_id = customer.cust_id " & _
        "  and cust_item.item_id = item.item_id "
    strSqlTx = strSqlTx & _
        "left join cust_group on cust_group.cust_grp_id = customer.cust_grp_id "
    strSqlTx = strSqlTx & _
        "left join name as an on an.obj_id = salesRepAssociate.assoc_id " & _
        "  and an.obj_nm = 'associate' " & _
        "  and an.nm_type_cd = 'STND' " & _
        "  and an.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as cn on cn.obj_id = customer.cust_id " & _
        "  and cn.obj_nm = 'customer' " & _
        "  and cn.nm_type_cd = 'STND' " & _
        "  and cn.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as sn on sn.obj_id = serviceRepAssociate.assoc_id " & _
        "  and sn.obj_nm = 'associate' " & _
        "  and sn.nm_type_cd = 'STND' " & _
        "  and sn.def_nm_fl = 'Y' "
    
    strSqlTx = strSqlTx & _
        "left join (SELECT lot_id, inv_loc_id AS jls_inv_loc_id from inventory_adj " & _
        "WHERE inv_loc_id = 'JLS' group by lot_id, inv_loc_id) jls_inv_adj " & _
        "ON jls_inv_adj.lot_id = inventory_adj.lot_id "
    
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' " & _
        "and sales_invoice.status_cd <> 'CA' "

                                        'execute the sql statement
    r_objConn.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenStatic
    Set objRecordset.ActiveConnection = r_objConn
    Call objRecordset.Open(strSqlTx)
                                        'loop thru recordset
    Do While objRecordset.EOF = False
        If IsNull(objRecordset("registration_dt")) = True Then
            dteCustRegDt = CDate("01/01/1901")
        Else
            dteCustRegDt = objRecordset("registration_dt")
        End If
        
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
                    
        dblSalesItemPrice = CDbl(objRecordset("si_item_price"))
        dblSalesItemCost = CDbl(objRecordset("lot_unit_cost"))
        dblOrigSalesItemCost = dblSalesItemCost
        If (isJLSItem(objRecordset("inv_loc_id"), replaceNull2(objRecordset("supplier_id")), replaceNull(objRecordset("jls_inv_loc_id")), dteSalesDt)) Then
            dblSalesItemCost = dblSalesItemCost * getJSLIncreaseFactor()
        End If
            
                    'if customer is goodmn, cust_id = 1160, check
                    'the itemid to see if we need to add in additional cost
                    'for the component item.
        lngItemId = objRecordset("item_id")
        If objRecordset("cust_id") = 1160 Or _
           objRecordset("cust_id") = 2325 Or _
           objRecordset("cust_id") = 1020 Or _
           objRecordset("cust_id") = 2271 Then
            If lngItemId = 25792 Or _
               lngItemId = 25793 Or _
               lngItemId = 25794 Or _
               lngItemId = 25795 Or _
               lngItemId = 25796 Or _
               lngItemId = 25797 Or _
               lngItemId = 25798 Or _
               lngItemId = 25799 Then
                dblSalesItemCost = dblSalesItemCost + addInAdditionalCost(r_objError, r_objConn, lngItemId)
            End If
        End If
        
        dblSalesItemQty = CDbl(objRecordset("adj_qty") * -1)
        dblSalesMarkUpPct = getMarkUpPct(r_objError, dblSalesItemCost, dblSalesItemPrice)
        If dteSalesDt >= CDate("1/1/2012") Then
            dblCommMarkUpPct = dblSalesMarkUpPct
        Else
            dblCommMarkUpPct = getMarkUpPct(r_objError, dblOrigSalesItemCost, dblSalesItemPrice)
        End If
        dblSalesPrice = dblSalesItemPrice * dblSalesItemQty
        dblSalesCost = dblSalesItemCost * dblSalesItemQty
        dblSalesProfitAm = dblSalesPrice - dblSalesCost
        
        dblLotUnitCost = objRecordset("lot_unit_cost")
        dblItemPrice = objRecordset("item_item_price")
        dblItemCost = objRecordset("item_item_cost")
                    
                    'if customer is Timco(1368) then set all the values to zero.
                    'this is because they use sales orders to timco to just transfer
                    'inventory from PCI location to Timco location.
        
        If objRecordset("cust_id") = 1368 Then
            dblSalesItemPrice = 0
            dblSalesItemCost = 0
            dblSalesMarkUpPct = 0
            dblSalesPrice = 0
            dblSalesCost = 0
            dblSalesProfitAm = 0
            dblLotUnitCost = 0
            dblItemPrice = 0
            dblItemCost = 0
        End If
        
        If IsNull(objRecordset("initial_purchase_dt")) = False Then
            dteInitialPurchaseDt = objRecordset("initial_purchase_dt")
        Else
            dteInitialPurchaseDt = CDate("12/31/2000")
        End If
        
        If dteSalesDt >= CDate("1/1/2012") Then
            dblSalesCommPct = getLatestCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), objRecordset("commodity_cd"))
        Else
            dblSalesCommPct = getCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), objRecordset("inv_loc_id"), replaceNull2(objRecordset("supplier_id")), replaceNull(objRecordset("jls_inv_loc_id")), objRecordset("commodity_cd"))
        End If
        dblSalesCommAm = Round(dblSalesItemQty * dblSalesItemPrice * dblSalesCommPct, 2)
        
        strLoadRecord = ""
        strLoadRecord = strLoadRecord _
          & replaceNull2(objRecordset("sales_rep_assoc_id")) & "<col>" _
          & replaceNull(objRecordset("sales_rep_assoc_nbr")) & "<col>" _
          & replaceNull(objRecordset("sales_rep_assoc_nm")) & "<col>" _
          & replaceNull2(objRecordset("service_rep_assoc_id")) & "<col>" _
          & replaceNull(objRecordset("service_rep_assoc_nbr")) & "<col>" _
          & replaceNull(objRecordset("service_rep_assoc_nm")) & "<col>" _
          & objRecordset("cust_id") & "<col>" _
          & objRecordset("cust_nbr") & "<col>" _
          & objRecordset("cust_nm") & "<col>" _
          & objRecordset("registration_dt") & "<col>"
        strLoadRecord = strLoadRecord _
          & objRecordset("cust_grp_id") & "<col>" _
          & objRecordset("cust_grp_nbr") & "<col>" _
          & objRecordset("cust_grp_nm") & "<col>" _
          & objRecordset("cust_item_nbr") & "<col>" _
          & objRecordset("cust_item_price") & "<col>" _
          & dblCustItemYrs & "<col>" _
          & objRecordset("sales_ord_id") & "<col>" _
          & objRecordset("ordered_dt") & "<col>" _
          & objRecordset("so_line_nbr") & "<col>" _
          & objRecordset("so_item_price") & "<col>" _
          & objRecordset("po_id") & "<col>" _
          & objRecordset("po_line_nbr") & "<col>" _
          & objRecordset("po_item_cost") & "<col>" _
          & objRecordset("ship_id") & "<col>" _
          & objRecordset("shipped_dt") & "<col>" _
          & objRecordset("drop_ship_fl") & "<col>" _
          & objRecordset("ship_type_cd") & "<col>" _
          & objRecordset("planned_ship_dt") & "<col>" _
          & objRecordset("planned_del_dt") & "<col>" _
          & objRecordset("sales_inv_id") & "<col>" _
          & objRecordset("discount_pct") & "<col>"
        strLoadRecord = strLoadRecord _
          & objRecordset("item_id") & "<col>" _
          & objRecordset("item_nbr") & "<col>" _
          & dblItemPrice & "<col>" _
          & dblItemCost & "<col>" _
          & "<col>" _
          & objRecordset("item_type_cd") & "<col>" _
          & objRecordset("commodity_cd") & "<col>" _
          & objRecordset("item_desc") & "<col>" _
          & dteInitialPurchaseDt & "<col>" _
          & objRecordset("lot_id") & "<col>" _
          & dblLotUnitCost & "<col>" _
          & objRecordset("inv_loc_id") & "<col>" _
          & "<col>" _
          & dblSalesMo & "<col>" _
          & dteSalesDt & "<col>" _
          & dblSalesItemPrice & "<col>" _
          & dblSalesItemCost & "<col>" _
          & dblSalesItemQty & "<col>" _
          & dblSalesMarkUpPct & "<col>" _
          & dblSalesCommPct & "<col>" _
          & dblSalesCommAm & "<col>" _
          & Round(dblSalesPrice, 2) & "<col>" _
          & Round(dblSalesCost, 2) & "<col>" _
          & Round(dblSalesProfitAm, 2) & "<row>"

        Print #v_lngFileId, strLoadRecord
        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Private Sub insertNonDropShipAssembledItems( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            , ByVal v_lngFileId As Long _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipAssembledItems"
          
    Dim objRecordset As adodb.Recordset, _
        strSqlTx As String, _
        dblSalesMarkUpPct As Double, _
        dblCommMarkUpPct As Double, _
        dblSalesCommPct As Double, _
        dblSalesCommAm As Double, _
        dteCustRegDt As Date, _
        dteInitialPurchaseDt As Date, _
        dteSalesDt As Date, _
        dblSalesItemPrice As Double, _
        dblSalesItemCost As Double, _
        dblOrigSalesItemCost As Double, _
        dblSalesItemQty As Double, _
        dblSalesPrice As Double, _
        dblSalesCost As Double, _
        dblSalesProfitAm As Double, _
        dblCustItemYrs As Double, _
        dblSalesMo As Double, _
        dblLotUnitCost As Double, _
        dblItemPrice As Double, _
        dblItemCost As Double, _
        lngItemId As Double, _
        strLoadRecord As String
        

    strSqlTx = _
        "select " & _
        "  salesRepAssociate.assoc_id as sales_rep_assoc_id " & _
        " ,salesRepAssociate.assoc_Nbr as sales_rep_assoc_nbr " & _
        " ,serviceRepAssociate.assoc_id as service_rep_assoc_id " & _
        " ,serviceRepAssociate.assoc_Nbr as service_rep_assoc_nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,cust_item.initial_purchase_dt " & _
        " ,cust_item.cust_item_nbr " & _
        " ,cust_item.cust_item_price " & _
        " ,cust_group.cust_grp_id " & _
        " ,cust_group.cust_grp_nbr " & _
        " ,cust_group.cust_grp_nm "
    strSqlTx = strSqlTx & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,assembly_order.assem_ord_id " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,sales_invoice_item.item_price as si_item_price " & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct "
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.assem_cost " & _
        " ,item.commodity_cd " & _
        " ,lot.lot_id " & _
        " ,lot.unit_cost as lot_unit_cost " & _
        " ,inventory_adj.adj_qty " & _
        " ,inventory_adj.inv_loc_id " & _
        " ,an.nm as sales_rep_assoc_nm " & _
        " ,sn.nm as service_rep_assoc_nm " & _
        " ,cn.nm as cust_nm "
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
        "left join associate salesRepAssociate on salesRepAssociate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "left join associate serviceRepAssociate on serviceRepAssociate.assoc_id = customer.cust_rep_assoc_id "
    strSqlTx = strSqlTx & _
        "left join cust_item on cust_item.cust_id = customer.cust_id " & _
        "  and cust_item.item_id = item.item_id "
    strSqlTx = strSqlTx & _
        "left join cust_group on cust_group.cust_grp_id = customer.cust_grp_id "
    strSqlTx = strSqlTx & _
        "left join name as an on an.obj_id = salesRepAssociate.assoc_id " & _
        "  and an.obj_nm = 'associate' " & _
        "  and an.nm_type_cd = 'STND' " & _
        "  and an.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as cn on cn.obj_id = customer.cust_id " & _
        "  and cn.obj_nm = 'customer' " & _
        "  and cn.nm_type_cd = 'STND' " & _
        "  and cn.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as sn on sn.obj_id = serviceRepAssociate.assoc_id " & _
        "  and sn.obj_nm = 'associate' " & _
        "  and sn.nm_type_cd = 'STND' " & _
        "  and sn.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'AS' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' " & _
        "and sales_invoice.status_cd <> 'CA' "
        
    r_objConn.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenStatic
    Set objRecordset.ActiveConnection = r_objConn
    Call objRecordset.Open(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        If IsNull(objRecordset("registration_dt")) = True Then
            dteCustRegDt = CDate("01/01/1901")
        Else
            dteCustRegDt = objRecordset("registration_dt")
        End If
        
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
                    
        dblSalesItemPrice = CDbl(objRecordset("si_item_price"))
        dblSalesItemCost = CDbl(objRecordset("lot_unit_cost"))
        dblOrigSalesItemCost = dblSalesItemCost
        If (isJLSItem(objRecordset("inv_loc_id"), 0, "", dteSalesDt)) Then
            dblSalesItemCost = dblSalesItemCost * getJSLIncreaseFactor()
        End If
            
                    'if customer is goodmn, cust_id = 1160, check
                    'the itemid to see if we need to add in additional cost
                    'for the component item.
        lngItemId = objRecordset("item_id")
        If objRecordset("cust_id") = 1160 Or _
           objRecordset("cust_id") = 2325 Or _
           objRecordset("cust_id") = 1020 Or _
           objRecordset("cust_id") = 2271 Then
            If lngItemId = 25792 Or _
               lngItemId = 25793 Or _
               lngItemId = 25794 Or _
               lngItemId = 25795 Or _
               lngItemId = 25796 Or _
               lngItemId = 25797 Or _
               lngItemId = 25798 Or _
               lngItemId = 25799 Then
                dblSalesItemCost = dblSalesItemCost + addInAdditionalCost(r_objError, r_objConn, lngItemId)
            End If
        End If
        
        dblSalesItemQty = CDbl(objRecordset("adj_qty") * -1)
        dblSalesMarkUpPct = getMarkUpPct(r_objError, dblSalesItemCost, dblSalesItemPrice)
        If dteSalesDt >= CDate("1/1/2012") Then
            dblCommMarkUpPct = dblSalesMarkUpPct
        Else
            dblCommMarkUpPct = getMarkUpPct(r_objError, dblOrigSalesItemCost, dblSalesItemPrice)
        End If
        dblSalesPrice = dblSalesItemPrice * dblSalesItemQty
        dblSalesCost = dblSalesItemCost * dblSalesItemQty
        dblSalesProfitAm = dblSalesPrice - dblSalesCost
        
        dblLotUnitCost = objRecordset("lot_unit_cost")
        dblItemPrice = objRecordset("item_item_price")
        dblItemCost = objRecordset("item_item_cost")
                    'if customer is Timco(1368) then set all the values to zero.
                    'this is because they use sales orders to timco to just transfer
                    'inventory from PCI location to Timco location.
        If objRecordset("cust_id") = 1368 Then
            dblSalesItemPrice = 0
            dblSalesItemCost = 0
            dblSalesMarkUpPct = 0
            dblSalesPrice = 0
            dblSalesCost = 0
            dblSalesProfitAm = 0
            dblLotUnitCost = 0
            dblItemPrice = 0
            dblItemCost = 0
        End If
        
        If IsNull(objRecordset("initial_purchase_dt")) = False Then
            dteInitialPurchaseDt = objRecordset("initial_purchase_dt")
        Else
            dteInitialPurchaseDt = CDate("12/31/2000")
        End If
        
        If dteSalesDt >= CDate("1/1/2012") Then
            dblSalesCommPct = getLatestCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), objRecordset("commodity_cd"))
        Else
            dblSalesCommPct = getCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), objRecordset("inv_loc_id"), 0, "", objRecordset("commodity_cd"))
        End If
        dblSalesCommAm = Round(dblSalesItemQty * dblSalesItemPrice * dblSalesCommPct, 2)
        
        strLoadRecord = ""
        strLoadRecord = strLoadRecord _
            & replaceNull2(objRecordset("sales_rep_assoc_id")) & "<col>" _
            & replaceNull(objRecordset("sales_rep_assoc_nbr")) & "<col>" _
            & replaceNull(objRecordset("sales_rep_assoc_nm")) & "<col>" _
            & replaceNull2(objRecordset("service_rep_assoc_id")) & "<col>" _
            & replaceNull(objRecordset("service_rep_assoc_nbr")) & "<col>" _
            & replaceNull(objRecordset("service_rep_assoc_nm")) & "<col>" _
            & objRecordset("cust_id") & "<col>" _
            & objRecordset("cust_nbr") & "<col>" _
            & objRecordset("cust_nm") & "<col>" _
            & objRecordset("registration_dt") & "<col>"
        strLoadRecord = strLoadRecord _
            & objRecordset("cust_grp_id") & "<col>" _
            & objRecordset("cust_grp_nbr") & "<col>" _
            & objRecordset("cust_grp_nm") & "<col>" _
            & objRecordset("cust_item_nbr") & "<col>" _
            & objRecordset("cust_item_price") & "<col>" _
            & dblCustItemYrs & "<col>" _
            & objRecordset("sales_ord_id") & "<col>" _
            & objRecordset("ordered_dt") & "<col>" _
            & objRecordset("so_line_nbr") & "<col>" _
            & objRecordset("so_item_price") & "<col>" _
            & "<col>" _
            & "<col>" _
            & "<col>" _
            & objRecordset("ship_id") & "<col>" _
            & objRecordset("shipped_dt") & "<col>" _
            & objRecordset("drop_ship_fl") & "<col>" _
            & objRecordset("ship_type_cd") & "<col>" _
            & objRecordset("planned_ship_dt") & "<col>" _
            & objRecordset("planned_del_dt") & "<col>" _
            & objRecordset("sales_inv_id") & "<col>" _
            & objRecordset("discount_pct") & "<col>"
        strLoadRecord = strLoadRecord _
            & objRecordset("item_id") & "<col>" _
            & objRecordset("item_nbr") & "<col>" _
            & dblItemPrice & "<col>" _
            & dblItemCost & "<col>" _
            & objRecordset("assem_cost") & "<col>" _
            & objRecordset("item_type_cd") & "<col>" _
            & objRecordset("commodity_cd") & "<col>" _
            & objRecordset("item_desc") & "<col>" _
            & dteInitialPurchaseDt & "<col>" _
            & objRecordset("lot_id") & "<col>" _
            & dblLotUnitCost & "<col>" _
            & objRecordset("inv_loc_id") & "<col>" _
            & objRecordset("assem_ord_id") & "<col>" _
            & dblSalesMo & "<col>" _
            & dteSalesDt & "<col>" _
            & dblSalesItemPrice & "<col>" _
            & dblSalesItemCost & "<col>" _
            & dblSalesItemQty & "<col>" _
            & dblSalesMarkUpPct & "<col>" _
            & dblSalesCommPct & "<col>" _
            & dblSalesCommAm & "<col>" _
            & Round(dblSalesPrice, 2) & "<col>" _
            & Round(dblSalesCost, 2) & "<col>" _
            & Round(dblSalesProfitAm, 2) & "<row>"
        
        Print #v_lngFileId, strLoadRecord
        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub

Private Sub insertDropShipPurchasedItems( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            , ByVal v_lngFileId As Long _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertDropShipPurchasedItems"
          
    Dim objRecordset As adodb.Recordset, _
        strSqlTx As String, _
        dblSalesMarkUpPct As Double, _
        dblCommMarkUpPct As Double, _
        dblSalesCommPct As Double, _
        dblSalesCommAm As Double, _
        dteCustRegDt As Date, _
        dteInitialPurchaseDt As Date, _
        dteSalesDt As Date, _
        dblSalesItemPrice As Double, _
        dblSalesItemCost As Double, _
        dblOrigSalesItemCost As Double, _
        dblSalesItemQty As Double, _
        dblSalesPrice As Double, _
        dblSalesCost As Double, _
        dblSalesProfitAm As Double, _
        dblCustItemYrs As Double, _
        dblSalesMo As Double, _
        strLoadRecord As String, _
        dblLotUnitCost As Double, _
        dblItemCost As Double, _
        dblItemPrice As Double, _
        lngItemId As Long

    strSqlTx = _
        "select " & _
        "  salesRepAssociate.assoc_id as sales_rep_assoc_id " & _
        " ,salesRepAssociate.assoc_Nbr as sales_rep_assoc_nbr " & _
        " ,serviceRepAssociate.assoc_id as service_rep_assoc_id " & _
        " ,serviceRepAssociate.assoc_Nbr as service_rep_assoc_nbr " & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,customer.registration_dt " & _
        " ,cust_item.initial_purchase_dt " & _
        " ,cust_item.cust_item_nbr " & _
        " ,cust_item.cust_item_price " & _
        " ,cust_group.cust_grp_id " & _
        " ,cust_group.cust_grp_nbr " & _
        " ,cust_group.cust_grp_nm "
    strSqlTx = strSqlTx & _
        " ,sales_order.sales_ord_id " & _
        " ,sales_order.ordered_dt " & _
        " ,sales_order_item.line_nbr as so_line_nbr " & _
        " ,sales_order_item.item_price as so_item_price " & _
        " ,purchase_order.po_id " & _
        " ,purchase_order_item.line_nbr as po_line_nbr " & _
        " ,purchase_order_item.item_cost as po_item_cost " & _
        " ,shipment.ship_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,shipment_item.loaded_qty " & _
        " ,sales_invoice_item.item_price as si_item_price" & _
        " ,sales_invoice.sales_inv_id " & _
        " ,sales_invoice.discount_pct "
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd " & _
        " ,an.nm as sales_rep_assoc_nm " & _
        " ,sn.nm as service_rep_assoc_nm " & _
        " ,cn.nm as cust_nm "
    
    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join sales_order on sales_order.sales_ord_id = shipment.sales_ord_id " & _
        "inner join sales_order_item on sales_order_item.sales_ord_id = sales_order.sales_ord_id " & _
        "  and sales_order_item.line_nbr = shipment_item.line_nbr " & _
        "inner join item on item.item_id = shipment_item.item_id " & _
        "inner join purchase_order on purchase_order.po_id = sales_order.drop_ship_po_id " & _
        "inner join purchase_order_item on purchase_order_item.po_id = sales_order.drop_ship_po_id " & _
        "  and purchase_order_item.line_nbr = sales_order_item.drop_ship_po_line_nbr " & _
        "inner join customer on customer.cust_id = sales_order.cust_id " & _
        "inner join sales_invoice on sales_invoice.ship_id = shipment_item.ship_id " & _
        "inner join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
        "  and sales_invoice_item.line_nbr = shipment_item.line_nbr " & _
        "left join associate salesRepAssociate on salesRepAssociate.assoc_id = sales_order.sales_rep_assoc_id " & _
        "left join associate serviceRepAssociate on serviceRepAssociate.assoc_id = customer.cust_rep_assoc_id "
    strSqlTx = strSqlTx & _
        "left join cust_item on cust_item.cust_id = customer.cust_id " & _
        "  and cust_item.item_id = item.item_id "
    strSqlTx = strSqlTx & _
        "left join cust_group on cust_group.cust_grp_id = customer.cust_grp_id "
    strSqlTx = strSqlTx & _
        "left join name as an on an.obj_id = salesRepAssociate.assoc_id " & _
        "  and an.obj_nm = 'associate' " & _
        "  and an.nm_type_cd = 'STND' " & _
        "  and an.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as cn on cn.obj_id = customer.cust_id " & _
        "  and cn.obj_nm = 'customer' " & _
        "  and cn.nm_type_cd = 'STND' " & _
        "  and cn.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as sn on sn.obj_id = serviceRepAssociate.assoc_id " & _
        "  and sn.obj_nm = 'associate' " & _
        "  and sn.nm_type_cd = 'STND' " & _
        "  and sn.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'Y' " & _
        "and shipment.type_cd <> '1L' " & _
        "and shipment_item.loaded_qty > 0 " & _
        "and sales_invoice.status_cd <> 'CA' "
        
    r_objConn.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenStatic
    Set objRecordset.ActiveConnection = r_objConn
    Call objRecordset.Open(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        If IsNull(objRecordset("registration_dt")) = True Then
            dteCustRegDt = CDate("01/01/1901")
        Else
            dteCustRegDt = objRecordset("registration_dt")
        End If
    
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
        
        dblSalesItemPrice = CDbl(objRecordset("si_item_price"))
        dblSalesItemCost = CDbl(objRecordset("po_item_cost"))
        dblOrigSalesItemCost = dblSalesItemCost
        If (isJLSItem("", 0, "", dteSalesDt)) Then
            dblSalesItemCost = dblSalesItemCost * getJSLIncreaseFactor()
        End If
            
                    'if customer is goodmn, cust_id = 1160, check
                    'the itemid to see if we need to add in additional cost
                    'for the component item.
        lngItemId = objRecordset("item_id")
        If objRecordset("cust_id") = 1160 Or _
           objRecordset("cust_id") = 2325 Or _
           objRecordset("cust_id") = 1020 Or _
           objRecordset("cust_id") = 2271 Then
            If lngItemId = 25792 Or _
               lngItemId = 25793 Or _
               lngItemId = 25794 Or _
               lngItemId = 25795 Or _
               lngItemId = 25796 Or _
               lngItemId = 25797 Or _
               lngItemId = 25798 Or _
               lngItemId = 25799 Then
                dblSalesItemCost = dblSalesItemCost + addInAdditionalCost(r_objError, r_objConn, lngItemId)
            End If
        End If
        
        dblSalesItemQty = CDbl(objRecordset("loaded_qty"))
        dblSalesPrice = dblSalesItemPrice * dblSalesItemQty
        dblSalesCost = dblSalesItemCost * dblSalesItemQty
        dblSalesProfitAm = dblSalesPrice - dblSalesCost
        dblSalesMarkUpPct = getMarkUpPct(r_objError, dblSalesItemCost, dblSalesItemPrice)
        If dteSalesDt >= CDate("1/1/2012") Then
            dblCommMarkUpPct = dblSalesMarkUpPct
        Else
            dblCommMarkUpPct = getMarkUpPct(r_objError, dblOrigSalesItemCost, dblSalesItemPrice)
        End If
        
        dblItemPrice = objRecordset("item_item_price")
        dblItemCost = objRecordset("item_item_cost")
                    
                    'if customer is Timco(1368) then set all the values to zero.
                    'this is because they use sales orders to timco to just transfer
                    'inventory from PCI location to Timco location.
        If objRecordset("cust_id") = 1368 Then
            dblSalesItemPrice = 0
            dblSalesItemCost = 0
            dblSalesMarkUpPct = 0
            dblSalesPrice = 0
            dblSalesCost = 0
            dblSalesProfitAm = 0
            dblItemPrice = 0
            dblItemCost = 0
        End If
        
        
        If IsNull(objRecordset("initial_purchase_dt")) = False Then
            dteInitialPurchaseDt = objRecordset("initial_purchase_dt")
        Else
            dteInitialPurchaseDt = CDate("12/31/2000")
        End If
        
        If dteSalesDt >= CDate("1/1/2012") Then
            dblSalesCommPct = getLatestCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), objRecordset("commodity_cd"))
        Else
            dblSalesCommPct = getCommPct(r_objError, dblCommMarkUpPct, dteInitialPurchaseDt, dteSalesDt, dblCustItemYrs, objRecordset("cust_id"), objRecordset("sales_rep_assoc_id"), "", 0, "", objRecordset("commodity_cd"))
        End If
        dblSalesCommAm = Round(dblSalesItemQty * dblSalesItemPrice * dblSalesCommPct, 2)
        
        strLoadRecord = ""
        strLoadRecord = strLoadRecord _
            & replaceNull2(objRecordset("sales_rep_assoc_id")) & "<col>" _
            & replaceNull(objRecordset("sales_rep_assoc_nbr")) & "<col>" _
            & replaceNull(objRecordset("sales_rep_assoc_nm")) & "<col>" _
            & replaceNull2(objRecordset("service_rep_assoc_id")) & "<col>" _
            & replaceNull(objRecordset("service_rep_assoc_nbr")) & "<col>" _
            & replaceNull(objRecordset("service_rep_assoc_nm")) & "<col>" _
            & objRecordset("cust_id") & "<col>" _
            & objRecordset("cust_nbr") & "<col>" _
            & objRecordset("cust_nm") & "<col>" _
            & objRecordset("registration_dt") & "<col>"
        strLoadRecord = strLoadRecord _
            & objRecordset("cust_grp_id") & "<col>" _
            & objRecordset("cust_grp_nbr") & "<col>" _
            & objRecordset("cust_grp_nm") & "<col>" _
            & objRecordset("cust_item_nbr") & "<col>" _
            & objRecordset("cust_item_price") & "<col>" _
            & dblCustItemYrs & "<col>" _
            & objRecordset("sales_ord_id") & "<col>" _
            & objRecordset("ordered_dt") & "<col>" _
            & objRecordset("so_line_nbr") & "<col>" _
            & objRecordset("so_item_price") & "<col>" _
            & objRecordset("po_id") & "<col>" _
            & objRecordset("po_line_nbr") & "<col>" _
            & objRecordset("po_item_cost") & "<col>" _
            & objRecordset("ship_id") & "<col>" _
            & objRecordset("shipped_dt") & "<col>" _
            & objRecordset("drop_ship_fl") & "<col>" _
            & objRecordset("ship_type_cd") & "<col>" _
            & objRecordset("planned_ship_dt") & "<col>" _
            & objRecordset("planned_del_dt") & "<col>" _
            & objRecordset("sales_inv_id") & "<col>" _
            & objRecordset("discount_pct") & "<col>"
        strLoadRecord = strLoadRecord _
            & objRecordset("item_id") & "<col>" _
            & objRecordset("item_nbr") & "<col>" _
            & dblItemPrice & "<col>" _
            & dblItemCost & "<col>" _
            & "<col>" _
            & objRecordset("item_type_cd") & "<col>" _
            & objRecordset("commodity_cd") & "<col>" _
            & objRecordset("item_desc") & "<col>" _
            & dteInitialPurchaseDt & "<col>" _
            & "<col>" _
            & "<col>" _
            & "<col>" _
            & "<col>" _
            & dblSalesMo & "<col>" _
            & dteSalesDt & "<col>" _
            & dblSalesItemPrice & "<col>" _
            & dblSalesItemCost & "<col>" _
            & dblSalesItemQty & "<col>" _
            & dblSalesMarkUpPct & "<col>" _
            & dblSalesCommPct & "<col>" _
            & dblSalesCommAm & "<col>" _
            & Round(dblSalesPrice, 2) & "<col>" _
            & Round(dblSalesCost, 2) & "<col>" _
            & Round(dblSalesProfitAm, 2) & "<row>"
        
        Print #v_lngFileId, strLoadRecord

        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub



Private Function getMarkUpPct( _
              ByRef r_objError As Variant _
            , ByVal v_dblItemCost As Double _
            , ByVal v_dblItemPrice As Double _
            ) As Double

    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getMarkUpPct"
          
    If v_dblItemCost > 0 Then
        getMarkUpPct = Round((v_dblItemPrice - v_dblItemCost) / v_dblItemCost, 4)
    Else
        getMarkUpPct = 0.26
    End If
    
    Exit Function
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Function


Private Function isJLSItem( _
              ByRef v_strInvLocId As String _
            , ByRef v_strSupplierId As Integer _
            , ByRef v_strJLSInvLocId As String _
            , ByRef v_dteSalesDt As Date _
            ) As Boolean

    If (v_strInvLocId = "JLS") Or _
       (v_strInvLocId = "Timco" And v_strSupplierId = 2079) Or _
       (UCase(v_strJLSInvLocId) = "JLS") Then
        If (v_dteSalesDt >= CDate("7/1/2010")) Then
            isJLSItem = True
        Else
            isJLSItem = False
        End If
    Else
        isJLSItem = False
    End If
    
End Function

Private Function getCommPct( _
              ByRef r_objError As Variant _
            , ByVal v_dblMarkUpPct As Double _
            , ByVal v_dteInitialPurchaseDt As Date _
            , ByVal v_dteSalesDt As Date _
            , ByRef r_dblCustItemYrs As Double _
            , ByRef v_dblCustId As Double _
            , ByRef v_dblAssocId As Variant _
            , ByRef v_strInvLocId As String _
            , ByRef v_strSupplierId As Integer _
            , ByRef v_strJLSInvLocId As String _
            , ByRef v_strCommodityCode _
            ) As Double

    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getCommPct"
    
        'if commodity code is "SC" (Setup Charge) set commition pct to 0
    If (v_strCommodityCode = "SC") Then
        getCommPct = 0
        Exit Function
    End If
    
    If IsNull(v_dblAssocId) Then
        v_dblAssocId = 0
    End If
    r_dblCustItemYrs = Round(DateDiff("d", v_dteInitialPurchaseDt, v_dteSalesDt) / 365, 4)
    
        'if parts are from inventory location SherCon, commision schedule is lowered
        'Mark up must be an additional 7% to recieve the same commission rate.
        '
        'Parts shiped from timco with a supplier of JLS (2079) get this also
        'Parts shiped from timco with a supplier of Shercon (2084) get this also
    Dim x As Double
    
    If (isJLSItem(v_strInvLocId, v_strSupplierId, v_strJLSInvLocId, v_dteSalesDt)) Then
        x = 0.07
    Else
        x = 0
    End If
    
    Select Case True
    Case r_dblCustItemYrs <= 1
        Select Case True
        Case v_dblMarkUpPct >= (0.25 + x)
            getCommPct = 0.07
        Case v_dblMarkUpPct >= (0.22 + x) And v_dblMarkUpPct < (0.25 + x)
            getCommPct = 0.05
        Case v_dblMarkUpPct >= (0.2 + x) And v_dblMarkUpPct < (0.22 + x)
            getCommPct = 0.04
        Case v_dblMarkUpPct >= (0.15 + x) And v_dblMarkUpPct < (0.2 + x)
            getCommPct = 0.02
        Case v_dblMarkUpPct >= (0.13 + x) And v_dblMarkUpPct < (0.15 + x)
            getCommPct = 0.01
        Case Else
            getCommPct = 0
        End Select
    Case r_dblCustItemYrs > 1 And r_dblCustItemYrs <= 2
        Select Case True
        Case v_dblMarkUpPct >= (0.25 + x)
            getCommPct = 0.07
        Case v_dblMarkUpPct >= (0.22 + x) And v_dblMarkUpPct < (0.25 + x)
            getCommPct = 0.05
        Case v_dblMarkUpPct >= (0.2 + x) And v_dblMarkUpPct < (0.22 + x)
            getCommPct = 0.04
        Case v_dblMarkUpPct >= (0.15 + x) And v_dblMarkUpPct < (0.2 + x)
            getCommPct = 0.02
        Case v_dblMarkUpPct >= (0.13 + x) And v_dblMarkUpPct < (0.15 + x)
            getCommPct = 0.01
        Case Else
            getCommPct = 0
        End Select
    Case r_dblCustItemYrs > 2 And r_dblCustItemYrs <= 3
        Select Case True
        Case v_dblMarkUpPct >= (0.25 + x)
            getCommPct = 0.0525
        Case v_dblMarkUpPct >= (0.22 + x) And v_dblMarkUpPct < (0.25 + x)
            getCommPct = 0.0375
        Case v_dblMarkUpPct >= (0.2 + x) And v_dblMarkUpPct < (0.22 + x)
            getCommPct = 0.03
        Case v_dblMarkUpPct >= (0.15 + x) And v_dblMarkUpPct < (0.2 + x)
            getCommPct = 0.01
        Case Else
            getCommPct = 0
        End Select
    Case r_dblCustItemYrs > 3
        Select Case True
        Case v_dblMarkUpPct >= (0.25 + x)
            getCommPct = 0.035
        Case v_dblMarkUpPct >= (0.22 + x) And v_dblMarkUpPct < (0.25 + x)
            getCommPct = 0.025
        Case v_dblMarkUpPct >= (0.2 + x) And v_dblMarkUpPct < (0.22 + x)
            getCommPct = 0.02
        Case v_dblMarkUpPct >= (0.15 + x) And v_dblMarkUpPct < (0.2 + x)
            getCommPct = 0.01
        Case Else
            getCommPct = 0
        End Select
    End Select

                    'if the sales rep is JPB and the customer
                    'is car/tx, then cut the sales commision in 1/2.
                    'per john kuzmick.
    If v_dblCustId = 1072 And v_dblAssocId = 2022 Then
        getCommPct = getCommPct / 2
    End If

                    'if the sales rep is JPB and the customer
                    'is AMAN/T, then cut the sales commision in 1/2.
                    'per john kuzmick.
    If v_dblCustId = 1020 And v_dblAssocId = 2022 Then
        getCommPct = getCommPct / 2
    End If

    
    Exit Function
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Function

Private Function getLatestCommPct( _
              ByRef r_objError As Variant _
            , ByVal v_dblMarkUpPct As Double _
            , ByVal v_dteInitialPurchaseDt As Date _
            , ByVal v_dteSalesDt As Date _
            , ByRef r_dblCustItemYrs As Double _
            , ByRef v_dblCustId As Double _
            , ByRef v_dblAssocId As Variant _
            , ByRef v_strCommodityCode _
            ) As Double

    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getCommPct"
    
        'if commodity code is "SC" (Setup Charge) set commition pct to 0
    If (v_strCommodityCode = "SC") Then
        getLatestCommPct = 0
        Exit Function
    End If
    
    If IsNull(v_dblAssocId) Then
        v_dblAssocId = 0
    End If
    r_dblCustItemYrs = Round(DateDiff("d", v_dteInitialPurchaseDt, v_dteSalesDt) / 365, 4)
    
    
    Select Case True
    Case r_dblCustItemYrs <= 1
        Select Case True
        Case v_dblMarkUpPct >= 0.25
            getLatestCommPct = 0.07
        Case v_dblMarkUpPct >= 0.22 And v_dblMarkUpPct < 0.25
            getLatestCommPct = 0.05
        Case v_dblMarkUpPct >= 0.2 And v_dblMarkUpPct < 0.22
            getLatestCommPct = 0.04
        Case v_dblMarkUpPct >= 0.15 And v_dblMarkUpPct < 0.2
            getLatestCommPct = 0.02
        Case v_dblMarkUpPct >= 0.13 And v_dblMarkUpPct < 0.15
            getLatestCommPct = 0.01
        Case Else
            getLatestCommPct = 0
        End Select
    Case r_dblCustItemYrs > 1 And r_dblCustItemYrs <= 2
        Select Case True
        Case v_dblMarkUpPct >= 0.25
            getLatestCommPct = 0.07
        Case v_dblMarkUpPct >= 0.22 And v_dblMarkUpPct < 0.25
            getLatestCommPct = 0.05
        Case v_dblMarkUpPct >= 0.2 And v_dblMarkUpPct < 0.22
            getLatestCommPct = 0.04
        Case v_dblMarkUpPct >= 0.15 And v_dblMarkUpPct < 0.2
            getLatestCommPct = 0.02
        Case v_dblMarkUpPct >= 0.13 And v_dblMarkUpPct < 0.15
            getLatestCommPct = 0.01
        Case Else
            getLatestCommPct = 0
        End Select
    Case r_dblCustItemYrs > 2 And r_dblCustItemYrs <= 3
        Select Case True
        Case v_dblMarkUpPct >= 0.25
            getLatestCommPct = 0.0525
        Case v_dblMarkUpPct >= 0.22 And v_dblMarkUpPct < 0.25
            getLatestCommPct = 0.0375
        Case v_dblMarkUpPct >= 0.2 And v_dblMarkUpPct < 0.22
            getLatestCommPct = 0.03
        Case v_dblMarkUpPct >= 0.15 And v_dblMarkUpPct < 0.2
            getLatestCommPct = 0.01
        Case Else
            getLatestCommPct = 0
        End Select
    Case r_dblCustItemYrs > 3
        Select Case True
        Case v_dblMarkUpPct >= 0.25
            getLatestCommPct = 0.035
        Case v_dblMarkUpPct >= 0.22 And v_dblMarkUpPct < 0.25
            getLatestCommPct = 0.025
        Case v_dblMarkUpPct >= 0.2 And v_dblMarkUpPct < 0.22
            getLatestCommPct = 0.02
        Case v_dblMarkUpPct >= 0.15 And v_dblMarkUpPct < 0.2
            getLatestCommPct = 0.01
        Case Else
            getLatestCommPct = 0
        End Select
    End Select

                    'if the sales rep is JPB and the customer
                    'is car/tx, then cut the sales commision in 1/2.
                    'per john kuzmick.
    If v_dblCustId = 1072 And v_dblAssocId = 2022 Then
        getLatestCommPct = getLatestCommPct / 2
    End If

                    'if the sales rep is JPB and the customer
                    'is AMAN/T, then cut the sales commision in 1/2.
                    'per john kuzmick.
    If v_dblCustId = 1020 And v_dblAssocId = 2022 Then
        getLatestCommPct = getLatestCommPct / 2
    End If

    
    Exit Function
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Function

Private Sub getSalesDt( _
              ByRef r_objError As Variant _
            , ByRef r_dteSalesDt As Date _
            , ByRef r_dblSalesMo As Double _
            , ByVal v_varShippedDt As Variant _
            , ByVal v_varPlannedShippedDt As Variant _
            , ByVal v_varPlannedDelDt As Variant _
            , ByVal v_strDropShipFl As String _
            )
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getSalesDt"
        
    If IsDate(v_varShippedDt) = False Then
        v_varShippedDt = CDate("01/01/1901")
    End If
    If IsDate(v_varPlannedShippedDt) = False Then
        v_varPlannedShippedDt = CDate("01/01/1901")
    End If
    If IsDate(v_varPlannedDelDt) = False Then
        v_varPlannedDelDt = CDate("01/01/1901")
    End If
    
    If v_strDropShipFl = "Y" Then
        r_dteSalesDt = v_varPlannedDelDt
    Else
        r_dteSalesDt = v_varPlannedShippedDt
    End If
    r_dblSalesMo = CDbl(Format(r_dteSalesDt, "yyyymm"))
    
    Exit Sub
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Private Function addInAdditionalCost( _
              ByRef r_objError As Variant _
            , ByRef r_objConn As adodb.Connection _
            , ByVal v_lngItemId As Long _
            ) As Double

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getSalesDt"
        
    Dim strSqlTx As String
    Dim dblItemCost As Double
    Dim lngItemId As Long
    
    Select Case v_lngItemId
        Case 25792
            lngItemId = 25779
        Case 25793
            lngItemId = 25780
        Case 25794
            lngItemId = 25781
        Case 25795
            lngItemId = 25782
        Case 25796
            lngItemId = 25783
        Case 25797
            lngItemId = 25785
        Case 25798
            lngItemId = 25784
        Case 25799
            lngItemId = 25788
        Case Else
            Call Err.Raise(1)
    End Select

    Dim objRecordset As adodb.Recordset
    Set objRecordset = New adodb.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenStatic
    Set objRecordset.ActiveConnection = r_objConn
    
    strSqlTx = "select item_cost from item where item_id = " & lngItemId
    
    Call objRecordset.Open(strSqlTx)

    dblItemCost = objRecordset("item_cost")
    
    addInAdditionalCost = dblItemCost
    
    Exit Function
errHandler:
    If r_objError.p_type_cd <> "E" And r_objError.p_type_cd <> "F" Then
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0100"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Function

Private Function replaceNull(ByVal inValue As Variant) As Variant
    If IsNull(inValue) Then
        replaceNull = ""
    Else
        replaceNull = inValue
    End If
End Function
Private Function replaceNull2(ByVal inValue As Variant) As Variant
    If IsNull(inValue) Then
        replaceNull2 = 0
    Else
        replaceNull2 = inValue
    End If
End Function

Private Function getJSLIncreaseFactor() As Double
    getJSLIncreaseFactor = 1.07
    If (Now >= CDate("2014-01-01")) Then
        getJSLIncreaseFactor = 1.1
    End If
End Function

