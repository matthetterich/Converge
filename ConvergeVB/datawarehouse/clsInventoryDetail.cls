VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventoryDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const M_STR_CLASS_NM As String = "clsInventoryDetail"

Public Sub extract( _
              ByRef r_objError As Variant _
            , ByRef r_conDwh As adodb.Connection _
            , ByRef r_conConverge As adodb.Connection _
            , ByVal v_strUnloadFileNm As String _
            )

    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "extract"
    
                                        'Clean up previous warehouse table
    Call dropTable(r_objError, r_conDwh)
                                        'Create New DWH Table
    Call createTable(r_objError, r_conDwh)
                                        'Insert shipment records
    Call insertRecords(r_objError, r_conDwh, r_conConverge, v_strUnloadFileNm)
                                        'Insert shipment records
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
            , ByRef r_objConnection As adodb.Connection _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "dropTable"
    
    Dim strSqlTx As String
    
                                        'drop the existing table
    strSqlTx = "drop table [dbo].[dwh_inventory_detail]"
    On Error Resume Next
    Call r_objConnection.Execute(strSqlTx)
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
            , ByRef r_objConnection As adodb.Connection _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "createTable"
    
    Dim strSqlTx As String
    
                                        'drop the existing table
                                        'create a new table
    
    strSqlTx = _
        "CREATE TABLE [dbo].[dwh_inventory_detail] ( " & _
        "[inv_adj_id] [int] NOT NULL , " & _
        "[inv_adj_qty] [int] NOT NULL , " & _
        "[inv_adj_ts] [datetime] NOT NULL , " & _
        "[item_id] [int] NOT NULL , " & _
        "[item_nbr] [nvarchar] (50)  NOT NULL , " & _
        "[item_price] [decimal](20, 5) NULL , " & _
        "[item_cost] [decimal](20, 5) NULL , " & _
        "[item_assem_cost] [decimal](20, 5) NULL , " & _
        "[item_type_cd] [nvarchar] (50)  NULL , " & _
        "[item_commodity_cd] [nvarchar] (50)  NULL , " & _
        "[item_desc] [nvarchar] (200)  NULL , " & _
        "[item_category_cd] [nvarchar] (50) NULL , " & _
        "[item_uom_cd] [nvarchar] (50) NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[lot_id] [int] NOT NULL , " & _
        "[lot_unit_cost] [numeric](18, 5) NULL , " & _
        "[lot_supplier_lot_id] [nvarchar] (50) NULL , " & _
        "[inv_bin_id] [nvarchar] (50) NULL , " & _
        "[inv_bin_desc] [nvarchar] (50) NULL , " & _
        "[inv_loc_id] [nvarchar] (50) NULL ," & _
        "[inv_loc_nm] [nvarchar] (50) NULL , " & _
        "[inv_loc_desc] [nvarchar] (50) NULL , " & _
        "[inv_loc_consignment_fl] [char] (1) NULL , " & _
        "[ship_id] [int] NULL , " & _
        "[ship_ret_id] [int] NULL , " & _
        "[sales_ord_id] [int] NULL , " & _
        "[so_line_nbr] [int] NULL , " & _
        ""
    strSqlTx = strSqlTx & _
        "[po_receipt_id] [int] NULL , " & _
        "[po_return_id] [int] NULL , " & _
        "[po_id] [int] NULL , " & _
        "[po_line_nbr] [int] NULL , " & _
        "[assem_ord_id] [int] NULL , " & _
        "[ac_line_nbr] [int] NULL , " & _
        "[phy_inv_id] [int] NULL , " & _
        "[prod_run_id] [int] NULL , " & _
        "[mat_req_id] [int] NULL , " & _
        "[cust_id] [int] NULL , " & _
        "[cust_nbr] [nvarchar] (50)  NULL , " & _
        "[supplier_id] [int] NULL , " & _
        "[supplier_nbr] [nvarchar] (50)  NULL , " & _
        "[assoc_id] [int] NULL , " & _
        "[assoc_nbr] [nvarchar] (50) NULL , " & _
        " "
    strSqlTx = strSqlTx & _
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

Private Sub insertRecords( _
              ByRef r_objError As Variant _
            , ByVal r_conDwh As adodb.Connection _
            , ByVal r_conConverge As adodb.Connection _
            , ByVal v_strUnloadFileNm As String _
            )
    
    On Error GoTo errHandler
    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_CLASS_NM & "." & "insertRecords"
          
    Dim objRecordset As adodb.Recordset, _
        strSqlTx As String, _
        strLoadRecord As String, _
        lngFileId As Long, _
        strOutRecord As String

    lngFileId = FreeFile
    Open v_strUnloadFileNm For Output As #lngFileId
    
                                        'create the sql statement
    strSqlTx = _
        "select " & _
        "  inventory_adj.inv_adj_id " & _
        " ,inventory_adj.item_id " & _
        " ,inventory_adj.lot_id " & _
        " ,inventory_adj.bin_id as inv_bin_id " & _
        " ,inventory_adj.inv_loc_id " & _
        " ,inventory_adj.adj_qty as inv_adj_qty" & _
        " ,inventory_adj.ship_id " & _
        " ,inventory_adj.ship_ret_id " & _
        " ,inventory_adj.sales_ord_id " & _
        " ,inventory_adj.so_line_nbr " & _
        " ,inventory_adj.po_receipt_id " & _
        " ,inventory_adj.po_return_id " & _
        " ,inventory_adj.po_id " & _
        " ,inventory_adj.po_line_nbr " & _
        " ,inventory_adj.assem_ord_id " & _
        " "
    strSqlTx = strSqlTx & _
        " ,inventory_adj.ac_line_nbr " & _
        " ,inventory_adj.phy_inv_id " & _
        " ,inventory_adj.prod_run_id " & _
        " ,inventory_adj.mat_req_id " & _
        " ,inventory_adj.entered_assoc_id " & _
        " ,inventory_adj.entered_ts inv_adj_ts " & _
        " "
    strSqlTx = strSqlTx & _
        " ,item.item_nbr " & _
        " ,item.item_cost " & _
        " ,item.item_price " & _
        " ,item.assem_cost as item_assem_cost " & _
        " ,item.type_cd as item_type_cd " & _
        " ,item.commodity_cd as item_commodity_cd " & _
        " ,item.category_cd as item_category_cd " & _
        " ,item.uom_cd item_uom_cd " & _
        " ,item.item_nbr " & _
        " ,item.item_desc " & _
        " "
    strSqlTx = strSqlTx & _
        " ,lot.supplier_lot_id as lot_supplier_lot_id " & _
        " ,lot.unit_cost as lot_unit_cost  " & _
        " ,inventory_bin.bin_desc as inv_bin_desc " & _
        " ,inventory_loc.inv_loc_nm " & _
        " ,inventory_loc.inv_loc_desc " & _
        " ,inventory_loc.consignment_fl inv_loc_consignment_fl " & _
        " ,sales_order.cust_id " & _
        " ,customer.cust_nbr " & _
        ", purchase_order.supplier_id " & _
        ", supplier.supplier_nbr " & _
        ", associate.assoc_id " & _
        ", associate.assoc_nbr " & _
        " "
    strSqlTx = strSqlTx & _
        " "
    strSqlTx = strSqlTx & _
        "From inventory_adj " & _
        "inner join item on item.item_id = inventory_adj.item_id " & _
        "inner join lot on lot.lot_id = inventory_adj.lot_id " & _
        "left join inventory_loc on inventory_loc.inv_loc_id = inventory_adj.inv_loc_id " & _
        "left join associate on associate.assoc_id = inventory_adj.entered_assoc_id " & _
        "left join inventory_bin on inventory_bin.bin_id = inventory_adj.bin_id " & _
        "left join sales_order on sales_order.sales_ord_id = inventory_adj.sales_ord_id " & _
        "left join customer on customer.cust_id = sales_order.cust_id " & _
        "left join purchase_order on purchase_order.po_id = inventory_adj.po_id " & _
        "left join supplier on supplier.supplier_id = purchase_order.supplier_id " & _
        " "
                                        'execute the sql statement
    r_conConverge.CommandTimeout = 7200
    Set objRecordset = New adodb.Recordset
    objRecordset.CursorLocation = adUseClient
    objRecordset.CursorType = adOpenStatic
    Set objRecordset.ActiveConnection = r_conConverge
    Call objRecordset.Open(strSqlTx)
                                        'loop thru recordset
    Do While objRecordset.EOF = False
        
        strOutRecord = ""
        strOutRecord = strOutRecord & _
            objRecordset("inv_adj_id") & "<col>" & _
            objRecordset("inv_adj_qty") & "<col>" & _
            objRecordset("inv_adj_ts") & "<col>" & _
            objRecordset("item_id") & "<col>" & _
            objRecordset("item_nbr") & "<col>" & _
            objRecordset("item_price") & "<col>" & _
            ""
        strOutRecord = strOutRecord & _
            objRecordset("item_cost") & "<col>" & _
            objRecordset("item_assem_cost") & "<col>" & _
            objRecordset("item_type_cd") & "<col>" & _
            objRecordset("item_commodity_cd") & "<col>" & _
            objRecordset("item_desc") & "<col>" & _
            objRecordset("item_category_cd") & "<col>" & _
            ""
        strOutRecord = strOutRecord & _
            objRecordset("item_uom_cd") & "<col>" & _
            objRecordset("lot_id") & "<col>" & _
            objRecordset("lot_unit_cost") & "<col>" & _
            objRecordset("lot_supplier_lot_id") & "<col>" & _
            ""
        strOutRecord = strOutRecord & _
            objRecordset("inv_bin_id") & "<col>" & _
            objRecordset("inv_bin_desc") & "<col>" & _
            objRecordset("inv_loc_id") & "<col>" & _
            objRecordset("inv_loc_nm") & "<col>" & _
            objRecordset("inv_loc_desc") & "<col>" & _
            objRecordset("inv_loc_consignment_fl") & "<col>" & _
            objRecordset("ship_id") & "<col>" & _
            objRecordset("ship_ret_id") & "<col>" & _
            ""
        strOutRecord = strOutRecord & _
            objRecordset("sales_ord_id") & "<col>" & _
            objRecordset("so_line_nbr") & "<col>" & _
            objRecordset("po_receipt_id") & "<col>" & _
            objRecordset("po_return_id") & "<col>" & _
            objRecordset("po_id") & "<col>" & _
            objRecordset("po_line_nbr") & "<col>" & _
            objRecordset("assem_ord_id") & "<col>" & _
            objRecordset("ac_line_nbr") & "<col>" & _
            objRecordset("phy_inv_id") & "<col>" & _
            ""
        strOutRecord = strOutRecord & _
            objRecordset("prod_run_id") & "<col>" & _
            objRecordset("mat_req_id") & "<col>" & _
            objRecordset("cust_id") & "<col>" & _
            objRecordset("cust_nbr") & "<col>" & _
            objRecordset("supplier_id") & "<col>" & _
            objRecordset("supplier_nbr") & "<col>" & _
            objRecordset("assoc_id") & "<col>" & _
            objRecordset("assoc_nbr") & "<row>"

        Print #lngFileId, strOutRecord
        Call objRecordset.MoveNext
    Loop
    Close #lngFileId
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    strSqlTx = _
      "BULK INSERT dwh_inventory_detail " _
    & "From '" & v_strUnloadFileNm & "' " _
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
    Call Err.Raise(1)
End Sub


