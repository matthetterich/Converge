VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchasingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const M_STR_CLASS_NM As String = "clsPurchasingDetail"

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
    Call insertDropShipPurchasedItems(r_objError, r_conConverge, lngFileId)
                                        'Insert shipment records
    Close #lngFileId
    
    strSqlTx = _
      "BULK INSERT dwh_purch_detail " _
    & "From '" & v_strUnloadFileNm & "' " _
    & "WITH ( " _
    & "  DATAFILETYPE = 'char' " _
    & ", FIELDTERMINATOR = '<col>' " _
    & ", ROWTERMINATOR = '<row>\n' " _
    & ", TABLOCK " _
    & ", KEEPNULLS " _
    & ") "
     
   strSqlTx = _
      "BULK INSERT dwh_purch_detail " _
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
    MsgBox (strRoutineNm & " " & Err.Description)
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
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "dropTable"
    
    Dim strSqlTx As String
    
                                        'drop the existing table
    strSqlTx = "drop table [dbo].[dwh_purch_detail]"
    On Error Resume Next
    Call r_objConn.Execute(strSqlTx)
    On Error GoTo 0
    
    Exit Sub
errHandler:
    MsgBox (strRoutineNm & " " & Err.Description)
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
        "CREATE TABLE [dbo].[dwh_purch_detail] ( " & _
        "[buyer_assoc_id]  [numeric](18, 0) NULL , " & _
        "[buyer_assoc_nbr] [nvarchar] (50)  NULL , " & _
        "[buyer_assoc_nm]  [nvarchar] (200) NULL , " & _
        "[supplier_id]     [numeric](18, 0) NULL , " & _
        "[supplier_nbr]    [nvarchar] (50)  NOT NULL , " & _
        "[supplier_nm]     [nvarchar] (200) NULL , " & _
        "[supplier_item_cost] [decimal](20, 5) NULL , " & _
        "[supplier_item_nbr] [nvarchar] (50)  NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[cust_id]         [numeric](18, 0) NULL , " & _
        "[cust_nbr]        [nvarchar] (50)  NULL , " & _
        "[cust_nm]         [nvarchar] (200)  NULL , " & _
        "[po_id]           [numeric](18, 0) NOT NULL , " & _
        "[po_drop_ship_fl] [char] (1)  NULL , " & _
        "[po_drop_ship_so_id] [numeric](18, 0) NULL , " & _
        "[po_type_cd]      [char] (4)  NULL , " & _
        "[poi_item_cost]   [decimal](20, 5) NOT NULL , " & _
        "[poi_item_desc]   [nvarchar] (200)  NULL , " & _
        "[poi_line_nbr]    [int] NULL , " & _
        "[poi_drop_ship_so_line_nbr] [int] NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[por_receipt_id]  [numeric](18, 0) NULL , " & _
        "[por_receipt_dt]  [datetime] NULL , " & _
        "[ship_id]         [numeric](18, 0) NULL , " & _
        "[ship_planned_del_dt] [datetime] NULL , " & _
        "[ship_loaded_qty] [numeric](18, 0) NULL , " & _
        "[item_id]         [numeric](18, 0) NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[item_nbr]        [nvarchar] (50)  NULL , " & _
        "[item_cost]       [decimal](20, 5) NULL , " & _
        "[item_commodity_cd] [char] (4)  NULL , " & _
        "[item_desc]       [nvarchar] (200)  NULL , " & _
        "[lot_id]          [numeric](18, 0) NULL , " & _
        "[lot_unit_cost]   [decimal](20, 5) NULL , " & _
        "[purch_mo]        [int] NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[purch_dt]        [datetime] NOT NULL , " & _
        "[purch_item_cost] [decimal](20, 5) NOT NULL ," & _
        "[purch_item_qty]  [numeric](18, 0) NOT NULL , " & _
        "[purch_cost]      [decimal](20, 2) NOT NULL  " & _
        ") ON [primary]" & _
        ""
    Call r_objConn.Execute(strSqlTx)
    
    Exit Sub
errHandler:
    MsgBox (strRoutineNm & " " & Err.Description)
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
        dblSalesCommPct As Double, _
        dblSalesCommAm As Double, _
        dteCustRegDt As Date, _
        dteInitialPurchaseDt As Date, _
        dteSalesDt As Date, _
        dblSalesItemPrice As Double, _
        dblSalesItemCost As Double, _
        dblSalesItemQty As Double, _
        dblSalesPrice As Double, _
        dblSalesCost As Double, _
        dblSalesProfitAm As Double, _
        dblCustItemYrs As Double, _
        dblSalesMo As Double, _
        strLoadRecord As String

                                        'create the sql statement
    strSqlTx = _
        "select " & _
        "  associate.assoc_id as buyer_assoc_id " & _
        " ,associate.assoc_nbr buyer_assoc_nbr " & _
        " ,an.nm buyer_assoc_nm " & _
        " ,supplier.supplier_id " & _
        " ,supplier.supplier_nbr " & _
        " ,sn.nm as supplier_nm " & _
        " ,supplier_item.supplier_item_cost " & _
        " ,supplier_item.supplier_item_nbr " & _
        ""
    strSqlTx = strSqlTx & _
        " ,purchase_order.po_id " & _
        " ,purchase_order.drop_ship_fl as po_drop_ship_fl " & _
        " ,purchase_order.type_cd as po_type_cd " & _
        " ,purchase_order_item.item_cost as poi_item_cost " & _
        " ,purchase_order_item.item_desc as poi_item_desc " & _
        " ,purchase_order_item.line_nbr as poi_line_nbr " & _
        " ,purchase_order_rec.po_receipt_id as por_receipt_id " & _
        " ,purchase_order_rec.receipt_dt as por_receipt_dt " & _
        ""
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_cost as item_cost " & _
        " ,item.commodity_cd as item_commodity_cd " & _
        " ,item.item_desc " & _
        " ,lot.lot_id " & _
        " ,lot.unit_cost as lot_unit_cost " & _
        " ,purchase_order_rec.receipt_dt as purch_dt " & _
        " ,purchase_order_item.item_cost as purch_item_cost " & _
        " ,inventory_adj.adj_qty as purch_item_qty " & _
        " ,purchase_order_item.item_cost * inventory_adj.adj_qty as purch_cost " & _
        ""
    strSqlTx = strSqlTx & _
        "From purchase_order_rec " & _
        "inner join purchase_order on purchase_order.po_id = purchase_order_rec.po_id " & _
        "inner join purchase_order_item on purchase_order_item.po_id = purchase_order_rec.po_id " & _
        "inner join inventory_adj on inventory_adj.po_receipt_id = purchase_order_rec.po_receipt_id " & _
        "  and inventory_adj.po_line_nbr = purchase_order_item.line_nbr " & _
        "inner join lot on lot.lot_id = inventory_adj.lot_id " & _
        "inner join item on item.item_id = purchase_order_item.item_id " & _
        "left join supplier on supplier.supplier_id = purchase_order.supplier_id " & _
        "left join associate on associate.assoc_id = purchase_order.buyer_assoc_id " & _
        ""
    strSqlTx = strSqlTx & _
        "left join supplier_item on supplier_item.supplier_id = supplier.supplier_id " & _
        "  and supplier_item.item_id = item.item_id " & _
        ""
    strSqlTx = strSqlTx & _
        "left join name as an on an.obj_id = associate.assoc_id " & _
        "  and an.obj_nm = 'associate' " & _
        "  and an.nm_type_cd = 'STND' " & _
        "  and an.def_nm_fl = 'Y' "
    strSqlTx = strSqlTx & _
        "left join name as sn on sn.obj_id = supplier.supplier_id " & _
        "  and sn.obj_nm = 'supplier' " & _
        "  and sn.nm_type_cd = 'STND' " & _
        "  and sn.def_nm_fl = 'Y' " & _
        ""
    strSqlTx = strSqlTx & _
        "where purchase_order.drop_ship_fl = 'N' "

                                        'execute the sql statement
    r_objConn.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    With objRecordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        Set .ActiveConnection = r_objConn
        Call .Open(strSqlTx)
    End With
                                        'loop thru recordset
    Do While objRecordset.EOF = False
        
        strLoadRecord = ""
        strLoadRecord = strLoadRecord _
          & objRecordset("buyer_assoc_id") & "<col>" _
          & objRecordset("buyer_assoc_nbr") & "<col>" _
          & objRecordset("buyer_assoc_nm") & "<col>" _
          & objRecordset("supplier_id") & "<col>" _
          & objRecordset("supplier_nbr") & "<col>" _
          & objRecordset("supplier_nm") & "<col>" _
          & objRecordset("supplier_item_cost") & "<col>" _
          & objRecordset("supplier_item_nbr") & "<col>" _
          & "<col>" _
          & "<col>" _
          & "<col>" _
          & objRecordset("po_id") & "<col>" _
          & objRecordset("po_drop_ship_fl") & "<col>" _
          & "<col>" _
          & objRecordset("po_type_cd") & "<col>" _
          & objRecordset("poi_item_cost") & "<col>" _
          & objRecordset("poi_item_desc") & "<col>" _
          & objRecordset("poi_line_nbr") & "<col>" _
          & "<col>" _
          & objRecordset("por_receipt_id") & "<col>" _
          & Format(objRecordset("por_receipt_dt"), "mm/dd/yyyy") & "<col>" _
          & "<col>" _
          & "<col>" _
          & "<col>"
        
        strLoadRecord = strLoadRecord _
          & objRecordset("item_id") & "<col>" _
          & objRecordset("item_nbr") & "<col>" _
          & objRecordset("item_cost") & "<col>" _
          & objRecordset("item_commodity_cd") & "<col>" _
          & objRecordset("item_desc") & "<col>" _
          & objRecordset("lot_id") & "<col>" _
          & objRecordset("lot_unit_cost") & "<col>" _
          & CDbl(Format(CDate(objRecordset("purch_dt")), "yyyymm")) & "<col>" _
          & Format(objRecordset("purch_dt"), "mm/dd/yyyy") & "<col>" _
          & objRecordset("purch_item_cost") & "<col>" _
          & objRecordset("purch_item_qty") & "<col>" _
          & objRecordset("purch_cost") & "<row>"

        Print #v_lngFileId, strLoadRecord
        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    MsgBox (strRoutineNm & " " & Err.Description)
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
        dblSalesCommPct As Double, _
        dblSalesCommAm As Double, _
        dteCustRegDt As Date, _
        dteInitialPurchaseDt As Date, _
        dteSalesDt As Date, _
        dblSalesItemPrice As Double, _
        dblSalesItemCost As Double, _
        dblSalesItemQty As Double, _
        dblSalesPrice As Double, _
        dblSalesCost As Double, _
        dblSalesProfitAm As Double, _
        dblCustItemYrs As Double, _
        dblSalesMo As Double, _
        strLoadRecord As String

    strSqlTx = _
        "select " & _
        "  associate.assoc_id as buyer_assoc_id " & _
        " ,associate.assoc_nbr buyer_assoc_nbr " & _
        " ,an.nm buyer_assoc_nm " & _
        " ,supplier.supplier_id " & _
        " ,supplier.supplier_nbr " & _
        " ,sn.nm as supplier_nm " & _
        " ,supplier_item.supplier_item_cost " & _
        " ,supplier_item.supplier_item_nbr " & _
        ""
    strSqlTx = strSqlTx & _
        " ,customer.cust_id " & _
        " ,customer.cust_nbr " & _
        " ,cn.nm as cust_nm " & _
        " ,purchase_order.po_id " & _
        " ,purchase_order.drop_ship_fl as po_drop_ship_fl " & _
        " ,sales_order.sales_ord_id as po_drop_ship_so_id " & _
        " ,purchase_order.type_cd as po_type_cd " & _
        " ,purchase_order_item.item_cost as poi_item_cost " & _
        " ,purchase_order_item.line_nbr as poi_line_nbr " & _
        " ,purchase_order_item.item_desc as poi_item_desc " & _
        " ,sales_order_item.line_nbr as poi_drop_ship_so_line_nbr " & _
        " ,shipment.ship_id " & _
        " ,shipment.planned_del_dt as ship_planned_del_dt " & _
        " ,shipment_item.loaded_qty as ship_loaded_qty "
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_cost as item_cost " & _
        " ,item.commodity_cd as item_commodity_cd " & _
        " ,item.item_desc " & _
        " ,shipment.planned_del_dt as purch_dt" & _
        " ,purchase_order_item.item_cost as purch_item_cost " & _
        " ,shipment_item.loaded_qty as purch_item_qty " & _
        " ,shipment_item.loaded_qty * purchase_order_item.item_cost as purch_cost "

    strSqlTx = strSqlTx & _
        " From shipment_item " & _
        " INNER JOIN shipment ON shipment.ship_id = shipment_item.Ship_id " & _
        " INNER JOIN sales_order ON sales_order.sales_ord_id = shipment.sales_ord_id " & _
        " INNER JOIN sales_order_item ON sales_order_item.sales_ord_id = sales_order.sales_ord_id AND sales_order_item.line_nbr = shipment_item.line_nbr " & _
        " INNER JOIN purchase_order ON purchase_order.po_id = sales_order.drop_ship_po_id " & _
        " INNER JOIN purchase_order_item ON purchase_order_item.po_id = sales_order.drop_ship_po_id AND purchase_order_item.line_nbr = sales_order_item.drop_ship_po_line_nbr " & _
        " LEFT OUTER JOIN item ON item.item_id = shipment_item.item_id " & _
        " LEFT OUTER JOIN customer ON customer.cust_id = sales_order.cust_id " & _
        " LEFT OUTER JOIN supplier ON supplier.supplier_id = purchase_order.supplier_id " & _
        " LEFT OUTER JOIN associate ON associate.assoc_id = purchase_order.buyer_assoc_id " & _
        " LEFT OUTER JOIN supplier_item ON supplier_item.supplier_id = supplier.supplier_id AND supplier_item.item_id = item.item_id " & _
        " LEFT OUTER JOIN name an ON an.obj_id = associate.assoc_id AND an.obj_nm = 'associate' AND an.nm_type_cd = 'STND' AND an.def_nm_fl = 'Y' " & _
        " LEFT OUTER JOIN name sn ON sn.obj_id = supplier.supplier_id AND sn.obj_nm = 'supplier' AND sn.nm_type_cd = 'STND' AND sn.def_nm_fl = 'Y' " & _
        " LEFT OUTER JOIN name cn ON cn.obj_id = customer.cust_id AND cn.obj_nm = 'customer' AND cn.nm_type_cd = 'STND' AND cn.def_nm_fl = 'Y' "
    
    strSqlTx = strSqlTx & _
        "where shipment.drop_ship_fl = 'Y' " & _
        "and shipment.type_cd <> '2L' " & _
        "and shipment_item.loaded_qty > 0 "
        
    r_objConn.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    With objRecordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        Set .ActiveConnection = r_objConn
        Call .Open(strSqlTx)
    End With
    
    Do While objRecordset.EOF = False
        
        strLoadRecord = ""
        strLoadRecord = strLoadRecord _
          & objRecordset("buyer_assoc_id") & "<col>" _
          & objRecordset("buyer_assoc_nbr") & "<col>" _
          & objRecordset("buyer_assoc_nm") & "<col>" _
          & objRecordset("supplier_id") & "<col>" _
          & objRecordset("supplier_nbr") & "<col>" _
          & objRecordset("supplier_nm") & "<col>" _
          & objRecordset("supplier_item_cost") & "<col>" _
          & objRecordset("supplier_item_nbr") & "<col>" _
          & objRecordset("cust_id") & "<col>" _
          & objRecordset("cust_nbr") & "<col>" _
          & objRecordset("cust_nm") & "<col>" _
          & objRecordset("po_id") & "<col>" _
          & objRecordset("po_drop_ship_fl") & "<col>" _
          & objRecordset("po_drop_ship_so_id") & "<col>" _
          & objRecordset("po_type_cd") & "<col>" _
          & objRecordset("poi_item_cost") & "<col>" _
          & objRecordset("poi_item_desc") & "<col>" _
          & objRecordset("poi_line_nbr") & "<col>" _
          & objRecordset("poi_drop_ship_so_line_nbr") & "<col>" _
          & "<col>" _
          & "<col>" _
          & objRecordset("ship_id") & "<col>" _
          & objRecordset("ship_planned_Del_dt") & "<col>" _
          & objRecordset("ship_loaded_qty") & "<col>"
        
        strLoadRecord = strLoadRecord _
          & objRecordset("item_id") & "<col>" _
          & objRecordset("item_nbr") & "<col>" _
          & objRecordset("item_cost") & "<col>" _
          & objRecordset("item_commodity_cd") & "<col>" _
          & objRecordset("item_desc") & "<col>" _
          & "<col>" _
          & "<col>" _
          & CDbl(Format(CDate(objRecordset("purch_dt")), "yyyymm")) & "<col>" _
          & Format(objRecordset("purch_dt"), "mm/dd/yyyy") & "<col>" _
          & objRecordset("purch_item_cost") & "<col>" _
          & objRecordset("purch_item_qty") & "<col>" _
          & objRecordset("purch_cost") & "<row>"

        Print #v_lngFileId, strLoadRecord

        Call objRecordset.MoveNext
    Loop
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    Exit Sub
errHandler:
    MsgBox (strRoutineNm & " " & Err.Description)
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



