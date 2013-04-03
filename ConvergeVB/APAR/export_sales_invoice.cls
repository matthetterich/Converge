VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "export_sales_invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_dbconnection As Object
Private mvaro_supplier_nm As Object
Private mvaro_supplier_addr As Object
Private mvaro_bill_nm As Object
Private mvaro_bill_addr As Object
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As Object
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Bill Name Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_bill_nm(ByVal vData As Object)
    Set mvaro_bill_nm = vData
End Property
Public Property Get o_bill_nm() As Object
    Set o_bill_nm = mvaro_bill_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Bill Address Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_bill_addr(ByVal vData As Object)
    Set mvaro_bill_addr = vData
End Property
Public Property Get o_bill_addr() As Object
    Set o_bill_addr = mvaro_bill_addr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for supplier Name Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_supplier_nm(ByVal vData As Object)
    Set mvaro_supplier_nm = vData
End Property
Public Property Get o_supplier_nm() As Object
    Set o_supplier_nm = mvaro_supplier_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for supplier Address Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_supplier_addr(ByVal vData As Object)
    Set mvaro_supplier_addr = vData
End Property
Public Property Get o_supplier_addr() As Object
    Set o_supplier_addr = mvaro_supplier_addr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Transaction Data Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Error Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Database Connection Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Database Connection Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Export the Customer Information to QuickBooks
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub ExportToQuickBooks()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    ' Define
    Dim sRoutine_nm As String
    Dim strSql As String
    Dim strLine As String
    Dim o_record_set As ADODB.Recordset
    Dim objFileName As FileSystemObject
    Dim tsrLine As TextStream
    Dim strFileName As String
    Dim dblUpdateQty As Double
    Dim o_SalesOrder As Object
    Dim o_Customer As Object
    Dim o_Item As Object
    Dim oldSalesOrder As Long
    Dim ID As Long
    Dim LoopTimes As Long
    Dim objRecordset As ADODB.Recordset
    Dim item_nbr As String
    Dim billAddr As String
    Dim shipAddr As String
    Dim stndName As String
    Dim billName As String
    Dim shipName As String
    Dim intLength As Long
    Dim intCount As Long
    Dim dtDueDate As Date
    Dim strSalesRep As String
    Dim strShipVia As String
    Dim strItemDesc As String
                    
    Set o_SalesOrder = CreateObject("converge_sales.sales_order")
    Set o_Customer = CreateObject("converge_sales.customer")
    Set o_Item = CreateObject("converge_item.item")
           
    ' Initialize
    Set o_SalesOrder.o_trans_data = o_trans_data
    Set o_SalesOrder.o_error = o_error
    
    Set o_Customer.o_trans_data = o_trans_data
    Set o_Customer.o_error = o_error
    Set o_Item.o_trans_data = o_trans_data
    Set o_Item.o_error = o_error
    Set o_record_set = New ADODB.Recordset 'CreateObject("adodb.recordset")
    sRoutine_nm = "export_sales_invoice.cls.ExportToQuickBooks"
    Set o_supplier_nm = CreateObject("converge.name")
    Set o_supplier_addr = CreateObject("converge.address")
    Set o_supplier_nm.o_trans_data = o_trans_data
    Set o_supplier_nm.o_error = o_error
    Set o_supplier_addr.o_trans_data = o_trans_data
    Set o_supplier_addr.o_error = o_error
    Set objFileName = New FileSystemObject
   
    strSql = " SELECT sales_invoice.sales_inv_id, sales_invoice_item.item_qty, " & _
    " shipment.sales_ord_id, shipment.cust_id, shipment.carrier_nm," & _
    " ROUND(sales_invoice_item.item_price, 2) AS item_price, shipment.shipped_dt, " & _
    " Item.item_nbr, sales_invoice_item.cust_item_nbr, sales_invoice_item.item_desc " & _
    " FROM item RIGHT OUTER JOIN " & _
    " sales_invoice_item ON " & _
    " item.item_id = sales_invoice_item.item_id RIGHT OUTER JOIN " & _
    " sales_invoice ON " & _
    " sales_invoice_item.sales_inv_id = sales_invoice.sales_inv_id LEFT Outer JOIN " & _
    " shipment ON " & _
    " sales_invoice.ship_id = shipment.ship_id " & _
    " WHERE (sales_invoice.status_cd = 'OP') " & _
    " AND (sales_invoice_item.item_qty > 0) " & _
    " ORDER BY sales_invoice.sales_inv_id "

    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)
        
    If Err.Number = 0 Then
    Else
        Let o_error.p_type_cd = "F"
        Let o_error.p_err_cd = "0300"
        Let o_error.p_routine_nm = sRoutine_nm
        Let o_error.p_message_id = 0
        Let o_error.p_nbr = Err.Number
        Let o_error.p_desc = Err.Description
        Call Err.Clear
        Exit Sub
    End If

    strFileName = "c:\temp\invoice-" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "-" & Hour(Now) & "-" & Minute(Now) & ".iif"
    
    Set tsrLine = objFileName.CreateTextFile(strFileName, True)
    
    tsrLine.WriteLine "!TRNS,TRNSID,TRNSTYPE,DATE,ACCNT,NAME," & _
                      "CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,TOPRINT," & _
                      "ADDR1,ADDR2,ADDR3,ADDR4,ADDR5,DUEDATE," & _
                      "TERMS,PAID,SHIPVIA,SHIPDATE,REP,FOB,PONUM,SADDR1," & _
                      "SADDR2,SADDR3,SADDR4,SADDR5"
                      
    tsrLine.WriteLine "!SPL,SPLID,TRNSTYPE,DATE,ACCNT,NAME," & _
                      "CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,QNTY," & _
                      "PRICE,INVITEM,PAYMETH,TAXABLE,REIMBEXP,EXTRA"
                      
    tsrLine.WriteLine "!ENDTRNS"
    
    oldSalesOrder = 0
    ID = 1
    LoopTimes = 0
    
    Do Until o_record_set.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
                
        o_SalesOrder.p_sales_ord_id = o_record_set("sales_ord_id")
        Call o_SalesOrder.getSalesOrder
        o_Customer.p_cust_id = o_record_set("cust_id")
        Call o_Customer.getCustomer
        
        If oldSalesOrder = o_record_set("sales_inv_id") Then
            
            If Len(o_record_set("cust_item_nbr")) > 0 Then
                item_nbr = o_record_set("cust_item_nbr")
            Else
                item_nbr = o_record_set("item_nbr")
            End If
            
            If Len(item_nbr) > 25 Then
                item_nbr = Left(item_nbr, 25)
            End If
            
            strItemDesc = Replace(o_record_set("item_desc"), ",", " ")
            strItemDesc = Replace(strItemDesc, Chr(34), "in")
            
            strLine = "SPL," & ID & "," & "INVOICE," & Month(Now) & "/" & Day(Now) & "/" & _
                      Year(Now) & "," & "Sales," & ",," & "-" & o_record_set("item_price") * o_record_set("item_qty") & "," & "," & strItemDesc & "," & "N," & _
                      "-" & o_record_set("item_qty") & "," & o_record_set("item_price") & "," & item_nbr & "," & "," & "N," & "N"
        
            ' Write Line to file name
            tsrLine.WriteLine strLine
            ID = ID + 1
        Else
            strSql = "SELECT SUM(ROUND(sales_invoice_item.item_price, 2) * sales_invoice_item.item_qty) as tot_price "
            strSql = strSql & " From sales_invoice_item "
            strSql = strSql & " Where sales_inv_id = '" & o_record_set("sales_inv_id") & "'"
            
            Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
            If LoopTimes > 0 Then
                strLine = "ENDTRNS"
                ' Write Line to file name
                tsrLine.WriteLine strLine
            End If
                
            oldSalesOrder = o_record_set("sales_inv_id")
                 
            billAddr = ""
            shipAddr = ""
            stndName = ""
            billName = ""
            shipName = ""
            strShipVia = ""
            
            intLength = Len(o_SalesOrder.o_bill_address.p_Street_addr)
            
            For intCount = 1 To intLength
                If Mid(o_SalesOrder.o_bill_address.p_Street_addr, intCount, 1) = vbLf Or Mid(o_SalesOrder.o_bill_address.p_Street_addr, intCount, 1) = vbCr Or Mid(o_SalesOrder.o_bill_address.p_Street_addr, intCount, 1) = "," Then
                    billAddr = billAddr & " "
                Else
                    billAddr = billAddr & Mid(o_SalesOrder.o_bill_address.p_Street_addr, intCount, 1)
                End If
            
            Next intCount
            
            intLength = Len(o_SalesOrder.o_ship_address.p_Street_addr)
            
            For intCount = 1 To intLength
                If Mid(o_SalesOrder.o_ship_address.p_Street_addr, intCount, 1) = vbLf Or Mid(o_SalesOrder.o_ship_address.p_Street_addr, intCount, 1) = vbCr Or Mid(o_SalesOrder.o_ship_address.p_Street_addr, intCount, 1) = "," Then
                    shipAddr = shipAddr & " "
                Else
                    shipAddr = shipAddr & Mid(o_SalesOrder.o_ship_address.p_Street_addr, intCount, 1)
                End If
            
            Next intCount
            
            intLength = Len(o_Customer.o_stnd_name.p_nm)
            
            For intCount = 1 To intLength
                If Mid(o_Customer.o_stnd_name.p_nm, intCount, 1) = vbLf Or Mid(o_Customer.o_stnd_name.p_nm, intCount, 1) = vbCr Or Mid(o_Customer.o_stnd_name.p_nm, intCount, 1) = "," Then
                    stndName = stndName & " "
                Else
                    stndName = stndName & Mid(o_Customer.o_stnd_name.p_nm, intCount, 1)
                End If
            
            Next intCount
            
            For intCount = 1 To intLength
                If Mid(o_SalesOrder.o_ship_name.p_nm, intCount, 1) = vbLf Or Mid(o_SalesOrder.o_ship_name.p_nm, intCount, 1) = vbCr Or Mid(o_SalesOrder.o_ship_name.p_nm, intCount, 1) = "," Then
                    shipName = shipName & " "
                Else
                    shipName = shipName & Mid(o_SalesOrder.o_ship_name.p_nm, intCount, 1)
                End If
            
            Next intCount
            
            For intCount = 1 To intLength
                If Mid(o_SalesOrder.o_bill_name.p_nm, intCount, 1) = vbLf Or Mid(o_SalesOrder.o_bill_name.p_nm, intCount, 1) = vbCr Or Mid(o_SalesOrder.o_bill_name.p_nm, intCount, 1) = "," Then
                    billName = billName & " "
                Else
                    billName = billName & Mid(o_SalesOrder.o_bill_name.p_nm, intCount, 1)
                End If
            
            Next intCount
            
            If Len(o_record_set("carrier_nm")) > 15 Then
                strShipVia = Left(o_record_set("carrier_nm"), 15)
            Else
                If Len(o_record_set("carrier_nm")) > 0 Then
                   strShipVia = o_record_set("carrier_nm")
                Else
                  strShipVia = ""
               
                End If
            End If
            
            If o_SalesOrder.p_payment_terms_cd = "CIA" Or o_SalesOrder.p_payment_terms_cd = "COD" Then
                dtDueDate = o_record_set("shipped_dt")
            Else
                dtDueDate = DateAdd("d", 30, o_record_set("shipped_dt"))
            End If
            
            Select Case UCase(o_SalesOrder.p_sales_rep_assoc_nbr)
                Case "JEH"
                    strSalesRep = "Joe Hoffman"
                Case "BK"
                    strSalesRep = "Bob Kuzmick"
                Case "JK"
                    strSalesRep = "Jim Kuzmick"
                Case "JPB"
                    strSalesRep = "Joe Budd"
                Case Else
                    strSalesRep = "Jim Kuzmick"
            End Select
                        
            strLine = "TRNS," & ID & "," & "INVOICE," & o_record_set("shipped_dt") & "," & _
                      "Accounts Receivable," & o_Customer.p_cust_nbr & "," & _
                      "," & objRecordset("tot_price") & "," & o_record_set("sales_inv_id") & "," & "," & "N," & "Y," & _
                      billName & "," & billAddr & "," & o_SalesOrder.o_bill_address.p_City_nm & " " & o_SalesOrder.o_bill_address.p_state_nm & _
                      " " & o_SalesOrder.o_bill_address.p_Zip_cd & ",,,," & _
                      o_SalesOrder.p_payment_terms_desc & ",N," & strShipVia & "," & o_record_set("shipped_dt") & _
                      "," & strSalesRep & "," & o_SalesOrder.p_fob_terms_desc & "," & o_SalesOrder.p_cust_po_nbr & _
                      "," & shipName & "," & shipAddr & "," & o_SalesOrder.o_ship_address.p_City_nm & " " & _
                      o_SalesOrder.o_ship_address.p_state_nm & " " & o_SalesOrder.o_ship_address.p_Zip_cd & ",-----------------------------------,Our Order Number " & o_SalesOrder.p_sales_ord_id
        
            ' Write Line to file name
            tsrLine.WriteLine strLine
            ID = ID + 1
            
            ' Parse for the Comma
            strItemDesc = Replace(o_record_set("item_desc"), ",", " ")
            strItemDesc = Replace(strItemDesc, Chr(34), "in")
            
            
            If Len(o_record_set("cust_item_nbr")) > 0 Then
                item_nbr = o_record_set("cust_item_nbr")
            Else
                item_nbr = o_record_set("item_nbr")
            End If
            
            If Len(item_nbr) > 25 Then
                item_nbr = Left(item_nbr, 25)
            End If
            
            strLine = "SPL," & ID & "," & "INVOICE," & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & "," & _
                      "Sales," & ",," & "-" & o_record_set("item_price") * o_record_set("item_qty") & "," & "," & _
                      strItemDesc & "," & "N," & _
                      "-" & o_record_set("item_qty") & "," & o_record_set("item_price") & "," & item_nbr & "," & "," & "N," & "N"
        
            ' Write Line to file name
            tsrLine.WriteLine strLine
            ID = ID + 1
        End If
        
        o_trans_data.p_update_fl = True
        strSql = "UPDATE sales_invoice SET export_fl = 'Y', status_cd='SN' " & _
                 " WHERE " & fWhereVariable("sales_inv_id", o_record_set("sales_inv_id"), "N", "=")
        
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        
        If Err.Number = 0 Then
        Else
            Let o_error.p_type_cd = "F"
            Let o_error.p_err_cd = "0300"
            Let o_error.p_routine_nm = sRoutine_nm
            Let o_error.p_message_id = 0
            Let o_error.p_nbr = Err.Number
            Let o_error.p_desc = Err.Description
            Call Err.Clear
            Exit Sub
        End If
            
        o_record_set.MoveNext
        LoopTimes = LoopTimes + 1
    Loop
    
    strLine = "ENDTRNS"
    
    ' Write Line to file name
    tsrLine.WriteLine strLine
    Set o_record_set = Nothing
    Set objRecordset = Nothing
    tsrLine.Close
    
End Sub
