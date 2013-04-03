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
Private Const M_STR_CLASS_NM As String = "clsSalesComm"

Public Sub extract( _
              ByRef r_objError As Variant _
            , ByVal v_strConnectStringTx As Variant _
            )

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "extract"
    
    Dim objConnection As ADODB.Connection, _
        dblDwhRowId As Double, _
        strConnectStringTx As String
    Set objConnection = New ADODB.Connection
    
    dblDwhRowId = 0
                                        'Open a the Connection
    Call objConnection.Open(v_strConnectStringTx)
                                        'Clean up previous warehouse table
    Call delete(r_objError, objConnection)
                                        'Insert shipment records
    Call insertNonDropShipPurchasedItems(r_objError, objConnection, dblDwhRowId)
                                        'Insert shipment records
    Call insertNonDropShipAssembledItems(r_objError, objConnection, dblDwhRowId)
                                        'Insert shipment records
    Call insertDropShipPurchasedItems(r_objError, objConnection, dblDwhRowId)
        
    Call buildAdditionalData(r_objError, objConnection)
    
    Call objConnection.Close
    Set objConnection = Nothing
    Exit Sub
    
errHandler:
    Set objConnection = Nothing
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


Private Sub delete( _
              ByRef r_objError As Variant _
            , ByRef r_objConnection As ADODB.Connection _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "delete"
    
    Dim strSqlTx As String
    
                                        'drop the existing table
    strSqlTx = "drop table [dbo].[dwh_sales_detail]"
    On Error Resume Next
    Call r_objConnection.Execute(strSqlTx)
    On Error GoTo 0
                                        'create a new table
    strSqlTx = _
        "CREATE TABLE [dbo].[dwh_sales_detail] ( " & _
        "[dwh_row_id] [numeric] (18, 0) not null , " & _
        "[sales_rep_assoc_id] [int] null , " & _
        "[sales_rep_assoc_nbr] [nvarchar] (50) null , " & _
        "[sales_rep_assoc_nm] [nvarchar] (200)  NULL , " & _
        "[cust_id] [int] null , " & _
        "[cust_nbr] [nvarchar] (50)  null , " & _
        "[cust_nm] [nvarchar] (200)  NULL , " & _
        "[cust_registration_dt] [datetime] NULL , " & _
        "[cust_grp_id] [int] NULL , " & _
        "[cust_grp_nbr] [nvarchar] (50) NULL , " & _
        "[cust_grp_nm] [nvarchar] (200) NULL , " & _
        "[cust_item_price] [decimal](20, 5) NULL , " & _
        "[sales_ord_id] [int] null , " & _
        "[sales_ord_ordered_dt] [datetime] null , " & _
        "[sales_ord_line_nbr] [int] null , " & _
        "[sales_ord_item_price] [decimal](20, 5) null , " & _
        "[po_id] [int] NULL , " & _
        "[po_line_nbr] [int] NULL , "
    strSqlTx = strSqlTx & _
        "[po_item_cost] [decimal](20, 5) NULL , " & _
        "[ship_id] [int] null , " & _
        "[ship_line_nbr] [int] null , " & _
        "[ship_shipped_dt] [datetime] null , " & _
        "[ship_drop_ship_fl] [char] (1)  null , " & _
        "[ship_type_cd] [nvarchar] (50)  null , " & _
        "[ship_planned_ship_dt] [datetime] NULL , " & _
        "[ship_planned_del_dt] [datetime] NULL , " & _
        "[sales_inv_id] [int] null , " & _
        "[sales_inv_discount_pct] [numeric](18, 4) NULL , " & _
        "[item_id] [int] null , "
    strSqlTx = strSqlTx & _
        "[item_nbr] [nvarchar] (50)  null , " & _
        "[item_price] [decimal](20, 5) null , " & _
        "[item_cost] [decimal](20, 5) null , " & _
        "[item_assem_cost] [decimal](20, 5) NULL , " & _
        "[item_type_cd] [nvarchar] (50)  null , " & _
        "[item_commodity_cd] [nvarchar] (50)  NULL , " & _
        "[item_desc] [nvarchar] (200)  NULL , " & _
        "[item_initial_purchase_dt] [datetime] null , " & _
        "[lot_id] [int] NULL , " & _
        "[lot_unit_cost] [numeric](18, 5) NULL , " & _
        "[assem_ord_id] [int] NULL , " & _
        "[sales_mo] [int] NULL , "
    strSqlTx = strSqlTx & _
        "[sales_dt] [datetime] null , " & _
        "[sales_item_price] [decimal](20, 5) null , " & _
        "[sales_item_cost] [decimal](20, 5) null ," & _
        "[sales_item_qty] [numeric](18, 0) null , " & _
        "[sales_margin_pct] [numeric](18, 4) null , " & _
        "[sales_comm_pct] [numeric](18, 4) null , " & _
        "[sales_comm_am] [decimal](20, 5) null , " & _
        "[sales_price] [decimal](20, 2) null , " & _
        "[sales_cost] [decimal](20, 2) null , " & _
        "[sales_profit_am] [decimal](20, 2) null  " & _
        ") ON [primary]"
    Call r_objConnection.Execute(strSqlTx)
    
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
            , ByRef r_objConnection As ADODB.Connection _
            , ByRef r_dblDwhRowId As Double _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipPurchasedItems"
          
    Dim objRecordset As ADODB.Recordset, _
        dteSalesDt As Date, _
        dblSalesMo As Double, _
        strSqlTx As String
                                        'create the sql statement
    strSqlTx = _
        "select " & _
        "  shipment.ship_id " & _
        " ,shipment.sales_ord_id " & _
        " ,shipment.cust_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,shipment_item.line_nbr as ship_line_nbr " & _
        " ,inventory_adj.adj_qty " & _
        " ,inventory_adj.lot_id " & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.assem_cost " & _
        " ,item.commodity_cd "

    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join inventory_adj on inventory_adj.ship_id = shipment_item.ship_id " & _
        "  and inventory_adj.so_line_nbr = shipment_item.line_nbr " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join item on item.item_id = shipment_item.item_id "
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' "

                                        'execute the sql statement
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
                                        'loop thru recordset
    Do While objRecordset.EOF = False
        
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
                    
        r_dblDwhRowId = r_dblDwhRowId + 1
        
        strSqlTx = _
            "insert into dwh_sales_detail (" & _
            "  dwh_row_id " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_planned_ship_dt " & _
            " ,ship_planned_del_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd "
        strSqlTx = strSqlTx & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price " & _
            " ,item_desc " & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,sales_item_qty " & _
            " ,sales_ord_id " & _
            " ,lot_id " & _
            " ,cust_id " & _
            " ,sales_dt " & _
            " ,sales_mo " & _
            " ,ship_line_nbr " & _
            " ) "
        
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(r_dblDwhRowId, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_ship_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_del_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_desc"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("adj_qty"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(dteSalesDt, "D") & _
            ", " & modUtilites.fInsertVariable(dblSalesMo, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_line_nbr"), "N") & _
            ")"

        Call r_objConnection.Execute(strSqlTx)
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
            , ByRef r_objConnection As ADODB.Connection _
            , ByRef r_dblDwhRowId As Double _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertNonDropShipAssembledItems"
          
    Dim objRecordset As ADODB.Recordset, _
        dteSalesDt As Date, _
        dblSalesMo As Double, _
        strSqlTx As String

    strSqlTx = _
        "select " & _
        "  shipment.ship_id " & _
        " ,shipment.sales_ord_id " & _
        " ,shipment.cust_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment_item.line_nbr as ship_line_nbr " & _
        " ,shipment.type_cd as ship_type_cd "
    strSqlTx = strSqlTx & _
        " ,inventory_adj.adj_qty " & _
        " ,inventory_adj.lot_id " & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.assem_cost " & _
        " ,item.commodity_cd "
    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join inventory_adj on inventory_adj.ship_id = shipment_item.ship_id " & _
        "  and inventory_adj.so_line_nbr = shipment_item.line_nbr " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join item on item.item_id = shipment_item.item_id "
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'AS' " & _
        "and shipment.drop_ship_fl = 'N' " & _
        "and shipment.type_cd <> '1L' "
        
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
                    
        r_dblDwhRowId = r_dblDwhRowId + 1
            
        strSqlTx = _
            "insert into dwh_sales_detail (" & _
            "  dwh_row_id " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_planned_ship_dt " & _
            " ,ship_planned_del_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd "
        strSqlTx = strSqlTx & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price " & _
            " ,item_desc " & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,sales_item_qty " & _
            " ,sales_ord_id " & _
            " ,lot_id " & _
            " ,cust_id " & _
            " ,sales_dt " & _
            " ,sales_mo " & _
            " ,ship_line_nbr " & _
            " ) "
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(r_dblDwhRowId, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_ship_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_del_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_desc"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("assem_cost"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("adj_qty"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("lot_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(dteSalesDt, "D") & _
            ", " & modUtilites.fInsertVariable(dblSalesMo, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_line_nbr"), "N") & _
            ")"
    
        Call r_objConnection.Execute(strSqlTx)
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
            , ByRef r_objConnection As ADODB.Connection _
            , ByRef r_dblDwhRowId As Double _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertDropShipPurchasedItems"
          
    Dim objRecordset As ADODB.Recordset, _
        dteSalesDt As Date, _
        dblSalesMo As Double, _
        strSqlTx As String

    strSqlTx = _
        "select " & _
        "  shipment.ship_id " & _
        " ,shipment.sales_ord_id " & _
        " ,shipment.cust_id " & _
        " ,shipment.shipped_dt " & _
        " ,shipment.planned_ship_dt " & _
        " ,shipment.planned_del_dt " & _
        " ,shipment.drop_ship_fl " & _
        " ,shipment.type_cd as ship_type_cd " & _
        " ,shipment_item.line_nbr as ship_line_nbr " & _
        " ,shipment_item.loaded_qty "
    strSqlTx = strSqlTx & _
        " ,item.item_id " & _
        " ,item.item_nbr " & _
        " ,item.item_price as item_item_price " & _
        " ,item.item_desc " & _
        " ,item.item_cost as item_item_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd "

    strSqlTx = strSqlTx & _
        "From shipment_item " & _
        "inner join shipment on shipment.ship_id = shipment_item.ship_id " & _
        "inner join item on item.item_id = shipment_item.item_id "
    strSqlTx = strSqlTx & _
        "where item.type_cd = 'PU' " & _
        "and shipment.drop_ship_fl = 'Y' " & _
        "and shipment.type_cd <> '1L' " & _
        "and shipment_item.loaded_qty > 0 "
        
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
    Do While objRecordset.EOF = False
        
        Call getSalesDt( _
                      r_objError _
                    , dteSalesDt _
                    , dblSalesMo _
                    , objRecordset("shipped_dt") _
                    , objRecordset("planned_ship_dt") _
                    , objRecordset("planned_del_dt") _
                    , objRecordset("drop_ship_fl") _
                    )
            
        r_dblDwhRowId = r_dblDwhRowId + 1
        
        strSqlTx = _
            "insert into dwh_sales_detail (" & _
            "  dwh_row_id " & _
            " ,ship_id " & _
            " ,ship_shipped_dt " & _
            " ,ship_planned_ship_dt " & _
            " ,ship_planned_del_dt " & _
            " ,ship_drop_ship_fl " & _
            " ,ship_type_cd "
        strSqlTx = strSqlTx & _
            " ,item_id " & _
            " ,item_nbr " & _
            " ,item_price " & _
            " ,item_desc " & _
            " ,item_cost " & _
            " ,item_assem_cost " & _
            " ,item_type_cd " & _
            " ,item_commodity_cd " & _
            " ,sales_item_qty " & _
            " ,sales_ord_id " & _
            " ,cust_id " & _
            " ,sales_dt " & _
            " ,sales_mo " & _
            " ,ship_line_nbr " & _
            " ) "
        strSqlTx = strSqlTx & _
            "values ( " & _
            "  " & modUtilites.fInsertVariable(r_dblDwhRowId, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("shipped_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_ship_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("planned_del_dt"), "D") & _
            ", " & modUtilites.fInsertVariable(objRecordset("drop_ship_fl"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_type_cd"), "S")
        strSqlTx = strSqlTx & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_nbr"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_price"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_desc"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_item_cost"), "N") & _
            ", " & modUtilites.fInsertVariable(Null, "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("item_type_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("commodity_cd"), "S") & _
            ", " & modUtilites.fInsertVariable(objRecordset("loaded_qty"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("sales_ord_id"), "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("cust_id"), "N") & _
            ", " & modUtilites.fInsertVariable(dteSalesDt, "D") & _
            ", " & modUtilites.fInsertVariable(dblSalesMo, "N") & _
            ", " & modUtilites.fInsertVariable(objRecordset("ship_line_nbr"), "N") & _
            ")"
    
        Call r_objConnection.Execute(strSqlTx)
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

Private Sub buildAdditionalData( _
              r_objError _
            , r_objConnection _
            )

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertDropShipPurchasedItems"
          
    Dim objRecordset As ADODB.Recordset, _
        rstCustomer As ADODB.Recordset, _
        rstSalesOrder As ADODB.Recordset, _
        rstSalesInvoice As ADODB.Recordset, _
        dteSalesDt As Date, _
        dblSalesMo As Double, _
        strSqlTx As String

    strSqlTx = _
        "select * " & _
        " from dwh_sales_detail "

                                        'execute the sql statement
    Set objRecordset = r_objConnection.Execute(strSqlTx)
    
                                        'loop thru recordset
    Do While objRecordset.EOF = False
        
        If objRecordset("ship_drop_ship_fl") = "N" Then
            strSqlTx = _
                "select " & _
                "  dwh_sales_detail.item_id " & _
                " ,dwh_sales_detail.item_nbr " & _
                " ,dwh_sales_detail.item_price as item_item_price " & _
                " ,dwh_sales_detail.item_desc " & _
                " ,dwh_sales_detail.item_cost as item_item_cost " & _
                " ,dwh_sales_detail.item_type_cd as item_type_cd " & _
                " ,dwh_sales_detail.item_commodity_cd " & _
                " ,dwh_sales_detail.lot_id " & _
                " ,dwh_sales_detail.sales_qty " & _
                " ,dwh_sales_detail.ship_id " & _
                " ,dwh_sales_detail.ship_shipped_dt " & _
                " ,dwh_sales_detail.ship_planned_ship_dt " & _
                " ,dwh_sales_detail.ship_planned_del_dt " & _
                " ,dwh_sales_detail.ship_drop_ship_fl " & _
                " ,dwh_sales_detail.ship_type_cd "
            strSqlTx = strSqlTx & _
                "  associate.assoc_id " & _
                " ,associate.assoc_Nbr " & _
                " ,customer.cust_id " & _
                " ,customer.cust_nbr " & _
                " ,customer.registration_dt " & _
                " ,cust_item.initial_purchase_dt " & _
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
                " ,assembly_order.assem_ord_id " & _
                " ,sales_invoice_item.item_price as si_item_price " & _
                " ,sales_invoice.sales_inv_id " & _
                " ,sales_invoice.discount_pct "
            strSqlTx = strSqlTx & _
                " ,lot.unit_cost as lot_unit_cost " & _
                " ,cust_item.cust_item_price " & _
                " ,an.nm as assoc_nm " & _
                " ,cn.nm as cust_nm "
        
            strSqlTx = strSqlTx & _
                "From dwh_sales_detail " & _
                "left join sales_order on sales_order.sales_ord_id = dwh_sales_detail.sales_ord_id " & _
                "left join sales_order_item on sales_order_item.sales_ord_id = dwh_sales_detail.sales_ord_id " & _
                "  and sales_order_item.line_nbr = dwh_sales_detail.ship_line_nbr " & _
                "left join lot on lot.lot_id = dwh_sales_detail.lot_id " & _
                "left join purchase_order on purchase_order.po_id = lot.po_id " & _
                "left join purchase_order_item on purchase_order_item.po_id = lot.po_id " & _
                "  and purchase_order_item.line_nbr = lot.po_line_nbr " & _
                "left join assembly_order on assembly_order.assem_ord_id = lot.assem_ord_id " & _
                "left join customer on customer.cust_id = dwh_sales_detail.cust_id " & _
                "left join sales_invoice on sales_invoice.ship_id = dwh_sales_detail.ship_id " & _
                "left join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
                "  and sales_invoice_item.line_nbr = dwh_sales_detail.ship_line_nbr " & _
                " join associate on associate.assoc_id = sales_order.sales_rep_assoc_id "
            strSqlTx = strSqlTx & _
                "left join cust_item on cust_item.cust_id = dwh_sales_detail.cust_id " & _
                "  and cust_item.item_id = dwh_sales_detail.item_id "
            strSqlTx = strSqlTx & _
                "left join cust_group on cust_group.cust_grp_id = customer.cust_grp_id "
            strSqlTx = strSqlTx & _
                "left join name as an on an.obj_id = associate.assoc_id " & _
                "  and an.obj_nm = 'associate' " & _
                "  and an.nm_type_cd = 'STND' " & _
                "  and an.def_nm_fl = 'Y' "
            strSqlTx = strSqlTx & _
                "left join name as cn on cn.obj_id = customer.cust_id " & _
                "  and cn.obj_nm = 'customer' " & _
                "  and cn.nm_type_cd = 'STND' " & _
                "  and cn.def_nm_fl = 'Y' "
            strSqlTx = strSqlTx & _
                "where dwh_row_id = " & objRecordset("dwh_row_id")
        Else
            
            strSqlTx = _
                "select " & _
                "  dwh_sales_detail.item_id " & _
                " ,dwh_sales_detail.item_nbr " & _
                " ,dwh_sales_detail.item_price as item_item_price " & _
                " ,dwh_sales_detail.item_desc " & _
                " ,dwh_sales_detail.item_cost as item_item_cost " & _
                " ,dwh_sales_detail.item_type_cd as item_type_cd " & _
                " ,dwh_sales_detail.item_commodity_cd " & _
                " ,dwh_sales_detail.lot_id " & _
                " ,dwh_sales_detail.sales_qty " & _
                " ,dwh_sales_detail.ship_id " & _
                " ,dwh_sales_detail.ship_shipped_dt " & _
                " ,dwh_sales_detail.ship_planned_ship_dt " & _
                " ,dwh_sales_detail.ship_planned_del_dt " & _
                " ,dwh_sales_detail.ship_drop_ship_fl " & _
                " ,dwh_sales_detail.ship_type_cd "
            strSqlTx = strSqlTx & _
                "  associate.assoc_id " & _
                " ,associate.assoc_Nbr " & _
                " ,customer.cust_id " & _
                " ,customer.cust_nbr " & _
                " ,customer.registration_dt " & _
                " ,cust_item.initial_purchase_dt " & _
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
                " ,sales_invoice_item.item_price as si_item_price" & _
                " ,sales_invoice.sales_inv_id " & _
                " ,sales_invoice.discount_pct "
            strSqlTx = strSqlTx & _
                " ,cust_item.cust_item_price " & _
                " ,an.nm as assoc_nm " & _
                " ,cn.nm as cust_nm "
            
            strSqlTx = strSqlTx & _
                "From dwh_sales_detail " & _
                "left join sales_order on sales_order.sales_ord_id = dwh_sales_detail.sales_ord_id " & _
                "left join sales_order_item on dwh_sales_detail.sales_ord_id = dwh_sales_detail.sales_ord_id " & _
                "  and sales_order_item.line_nbr = dwh_sales_detail.ship_line_nbr " & _
                "left join purchase_order on purchase_order.po_id = sales_order.drop_ship_po_id " & _
                "left join purchase_order_item on purchase_order_item.po_id = sales_order.drop_ship_po_id " & _
                "  and purchase_order_item.line_nbr = sales_order_item.drop_ship_po_line_nbr " & _
                "left join customer on customer.cust_id = dwh_sales_detail.cust_id " & _
                "left join sales_invoice on sales_invoice.ship_id = dwh_sales_detail.ship_id " & _
                "left join sales_invoice_item on sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id " & _
                "  and sales_invoice_item.line_nbr = dwh_sales_detail.ship_line_nbr " & _
                "left join associate on associate.assoc_id = sales_order.sales_rep_assoc_id "
            strSqlTx = strSqlTx & _
                "left join cust_item on cust_item.cust_id = customer.cust_id " & _
                "  and cust_item.item_id = item.item_id "
            strSqlTx = strSqlTx & _
                "left join cust_group on cust_group.cust_grp_id = customer.cust_grp_id "
            strSqlTx = strSqlTx & _
                "left join name as an on an.obj_id = associate.assoc_id " & _
                "  and an.obj_nm = 'associate' " & _
                "  and an.nm_type_cd = 'STND' " & _
                "  and an.def_nm_fl = 'Y' "
            strSqlTx = strSqlTx & _
                "left join name as cn on cn.obj_id = customer.cust_id " & _
                "  and cn.obj_nm = 'customer' " & _
                "  and cn.nm_type_cd = 'STND' " & _
                "  and cn.def_nm_fl = 'Y' "
            strSqlTx = strSqlTx & _
                "where dwh_row_id = " & objRecordset("dwh_row_id")
        End If
    
        Set objRecordset2 = r_objConnection.Execute(strSqlTx)
    
        strSqlTx = _
            "Update dwh_sales_detail set " & _
            "  " & fUpdateVariable("sales_rep_assoc_id", objRecordset2("sales_rep_assoc_id"), "N", "=") & _
            ", " & fUpdateVariable("sales_rep_assoc_nbr", objRecordset2("sales_rep_assoc_nbr"), "S", "=") & _
            ", " & fUpdateVariable("sales_rep_assoc_nm", objRecordset2(""), "S", "=") & _
            ", " & fUpdateVariable("cust_nbr", objRecordset2(""), "S", "=") & _
            ", " & fUpdateVariable("cust_nm", objRecordset2(""), "S", "=") & _
            ", " & fUpdateVariable("cust_registration_dt", objRecordset2(""), "D", "=") & _
            ", " & fUpdateVariable("cust_grp_id", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("cust_grp_nbr", objRecordset2(""), "S", "=") & _
            ", " & fUpdateVariable("cust_grp_nm", objRecordset2(""), "S", "=") & _
            ", " & fUpdateVariable("cust_item_price", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_ord_ordered_dt", objRecordset2(""), "D", "=")
        strSqlTx = strSqlTx & _
            ", " & fUpdateVariable("sales_ord_line_nbr", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_ord_item_price", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("po_id", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("po_line_nbr", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("po_item_cost", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_inv_id", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_inv_discount_pct", objRecordset2(""), "N", "=")
        strSqlTx = strSqlTx & _
            ", " & fUpdateVariable("item_initial_purchase_dt", objRecordset2(""), "D", "=") & _
            ", " & fUpdateVariable("lot_unit_cost", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("assem_ord_id", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_item_price", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_item_cost", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_item_qty", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_margin_pct", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_comm_pct", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_comm_am", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_price", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_cost", objRecordset2(""), "N", "=") & _
            ", " & fUpdateVariable("sales_profit_am", objRecordset2(""), "N", "=") & _
            " where " & _
            " " & fWhereVariable("dwh_row_id", objRecordset("dwh_row_id"), "N", "=")
            
            Call r_objConnection.Execute(strSqlTx)
            Call objRecordset2.Close
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



Private Function getMarginPct( _
              ByRef r_objError As Variant _
            , ByVal v_dblItemCost As Double _
            , ByVal v_dblItemPrice As Double _
            ) As Double

    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getMarginPct"
          
    If v_dblItemCost > 0 Then
        getMarginPct = Round((v_dblItemPrice - v_dblItemCost) / v_dblItemCost, 4)
    Else
        getMarginPct = 0.26
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

Private Function getCommPct( _
              ByRef r_objError As Variant _
            , ByVal v_dblMarginPct As Double _
            , ByVal v_dteInitialPurchaseDt As Date _
            , ByVal v_dteOrderedDt As Date _
            ) As Double

    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getCommPct"
          
    If DateDiff("yyyy", v_dteOrderedDt, v_dteInitialPurchaseDt) <= 1 Then
        If v_dblMarginPct >= 0.25 Then
            getCommPct = 0.07
        ElseIf v_dblMarginPct >= 0.22 Then
            getCommPct = 0.05
        ElseIf v_dblMarginPct >= 0.2 Then
            getCommPct = 0.04
        ElseIf v_dblMarginPct >= 0.15 Then
            getCommPct = 0.02
        ElseIf v_dblMarginPct >= 0.13 Then
            getCommPct = 0.01
        Else
            getCommPct = 0
        End If
    ElseIf DateDiff("yyyy", v_dteOrderedDt, v_dteInitialPurchaseDt) <= 2 Then
        If v_dblMarginPct >= 0.25 Then
            getCommPct = 0.07
        ElseIf v_dblMarginPct >= 0.22 Then
            getCommPct = 0.05
        ElseIf v_dblMarginPct >= 0.2 Then
            getCommPct = 0.04
        ElseIf v_dblMarginPct >= 0.15 Then
            getCommPct = 0.02
        ElseIf v_dblMarginPct >= 0.13 Then
            getCommPct = 0.01
        Else
            getCommPct = 0
        End If
    ElseIf DateDiff("yyyy", v_dteOrderedDt, v_dteInitialPurchaseDt) <= 3 Then
        If v_dblMarginPct >= 0.25 Then
            getCommPct = 0.0525
        ElseIf v_dblMarginPct >= 0.22 Then
            getCommPct = 0.0375
        ElseIf v_dblMarginPct >= 0.2 Then
            getCommPct = 0.03
        ElseIf v_dblMarginPct >= 0.15 Then
            getCommPct = 0.01
        Else
            getCommPct = 0
        End If
    Else
        If v_dblMarginPct >= 0.25 Then
            getCommPct = 0.035
        ElseIf v_dblMarginPct >= 0.22 Then
            getCommPct = 0.025
        ElseIf v_dblMarginPct >= 0.2 Then
            getCommPct = 0.02
        ElseIf v_dblMarginPct >= 0.15 Then
            getCommPct = 0.01
        Else
            getCommPct = 0
        End If
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
            , ByVal v_dteShippedDt As Date _
            , ByVal v_dtePlannedShippedDt As Date _
            , ByVal v_dtePlannedDelDt As Date _
            , ByVal v_strDropShipFl As String _
            )
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "getSalesDt"
        
    If v_strDropShipFl = "Y" Then
        r_dteSalesDt = v_dtePlannedDelDt
    Else
        r_dteSalesDt = v_dtePlannedShippedDt
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

