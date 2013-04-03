VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const M_STR_MODULE_NM = "clsInterface"
Private m_objQBSessionMgr As QBSessionManager
Private m_blnSessionOpenFl As Boolean
Private Const M_STR_QB_XML_COUNTRY_CD As String = "US"
Private Const M_INT_QB_XML_MAJOR_VERSION As Integer = 4
Private Const M_INT_QB_XML_MINOR_VERSION As Integer = 0

Public Enum enuQBTransType
   qtCustomer = 1
   qtSupplier = 2
   qtItem = 3
   qtInvoice = 4
End Enum


Public Function ProcessTrans( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByRef r_lngErrorCt As Long _
        , ByRef r_lngLoadCt As Long _
        ) As Long

    On Error GoTo ErrHandler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "ProcessTrans"
    
    Dim strSqlTx As String
    Dim dblUpdatedQty As Double
    Dim lngTransTypeCd As Long
    Dim lngCustId As Long
    Dim lngSupplierId As Long
    Dim lngItemId As Long
    Dim lngSalesInvId As Long
    Dim rstData As ADODB.Recordset
    Dim dblUpdateQty As Double
    Dim strErrorDesc As String
    
    ProcessTrans = 0
    strSqlTx = "select * from quickbooks_trans order by trans_id "
    Call r_objTransData.OpenRecordset(rstData, strSqlTx, r_objError)
                    'Create the Quickbooks session manager object
    
    If frmQBLoader.Check1.Value = 1 Then
        MsgBox ("Transaction table is empty: " & rstData.EOF)
        If rstData.EOF = False Then
            MsgBox (rstData.RecordCount & "tranactions exist in the transaction table")
        Else
            MsgBox (0 & "tranactions exist in the transaction table")
        End If
    End If
    
    
    If rstData.EOF = False Then
        If m_objQBSessionMgr Is Nothing Then
            If frmQBLoader.Check1.Value = 1 Then
                MsgBox ("session manager is not set")
                MsgBox ("connecting to the following database: " & GetSetting("CONVERGE", "QUICKBOOKS", "DBNAME", ""))
            End If
            Set m_objQBSessionMgr = New QBSessionManager
            With m_objQBSessionMgr
                Call .OpenConnection("Converge", "Converge")
                Call .BeginSession(GetSetting("CONVERGE", "QUICKBOOKS", "DBNAME", ""), omDontCare)
'               Call .BeginSession("", omDontCare)
            End With
            m_blnSessionOpenFl = True
            If frmQBLoader.Check1.Value = 1 Then
                MsgBox ("Successfully connected to quickbooks")
            End If
        Else
            If frmQBLoader.Check1.Value = 1 Then
                MsgBox ("session manager is already set")
            End If
        End If
    End If
    
    Do While rstData.EOF = False
        lngTransTypeCd = rstData("trans_type_cd")
        
        If frmQBLoader.Check1.Value = 1 Then
            MsgBox ("looping through transactions, trans id: " & rstData("trans_id") & " trans_type_cd " & rstData("trans_type_cd"))
        End If
        
        On Error Resume Next
        If lngTransTypeCd = qtCustomer Then
            lngCustId = rstData("cust_id")
            Call Me.syncCust(r_objError, r_objTransData, m_objQBSessionMgr, lngCustId)
        ElseIf lngTransTypeCd = qtSupplier Then
            lngSupplierId = rstData("supplier_id")
            Call Me.syncSupplier(r_objError, r_objTransData, m_objQBSessionMgr, lngSupplierId)
        ElseIf lngTransTypeCd = qtInvoice Then
            lngSalesInvId = rstData("sales_inv_id")
            Call Me.syncInvoice(r_objError, r_objTransData, m_objQBSessionMgr, lngSalesInvId)
        ElseIf lngTransTypeCd = qtItem Then
            lngItemId = rstData("item_id")
            Call Me.syncItem(r_objError, r_objTransData, m_objQBSessionMgr, lngItemId)
        End If
        If frmQBLoader.Check1.Value = 1 Then
            If Err.Number <> 0 Then
                MsgBox ("An error has occured, err description " & Err.Description)
            End If
        End If
        If Err.Number <> 0 Then
            Call Err.Clear
        End If
        On Error GoTo ErrHandler
        
        If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
            strErrorDesc = r_objError.p_desc
            strSqlTx = "update quickbooks_trans set " _
                    & " err_routine_nm = '" & r_objError.p_routine_nm & "' " _
                    & ",err_msg = '" & strErrorDesc & "' " _
                    & " where trans_id = " & rstData("trans_id") & ""
            Set r_objError = CreateObject("converge.error")
            Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
            r_lngErrorCt = r_lngErrorCt + 1
            
            If InStr(1, strErrorDesc, "BeginSession") > 0 Then
                m_objQBSessionMgr.EndSession
                m_objQBSessionMgr.CloseConnection
                Set m_objQBSessionMgr = Nothing
                m_blnSessionOpenFl = False
                Exit Function
            End If
        
        Else
            strSqlTx = "delete from quickbooks_trans where trans_id = " & rstData("trans_id") & " "
            Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
            r_lngLoadCt = r_lngLoadCt + 1
        End If
        Call rstData.MoveNext
    Loop
    
    Exit Function
ErrHandler:
    If m_blnSessionOpenFl = True Then
        m_objQBSessionMgr.EndSession
        m_objQBSessionMgr.CloseConnection
        Set m_objQBSessionMgr = Nothing
        m_blnSessionOpenFl = False
    End If
    Set m_objQBSessionMgr = Nothing
        
    If frmQBLoader.Check1.Value = 1 Then
       MsgBox ("An error has occured, err description " & Err.Description)
    End If
    
    
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Call Err.Clear
End Function

Public Sub syncInvoice( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_varSalesInvId As Variant _
                )
    
    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice"
    
    Dim objSalesInvoice As Object
    Dim objSalesOrder As Object
    Dim objShipment As Object
    Dim objSalesInvoiceItem As Object
    Dim objItem As Object
    Dim objAssociate As Object
    Dim varQBItemPrice As Variant
    Dim strSql As String
    Dim dblUpdateQty As Double
    Dim strItemNbr As String
    Dim rstInvoiceItems As ADODB.Recordset
    Dim objQBRequestMsgSet As IMsgSetRequest
    Dim objQBResponseMsgSet As IMsgSetResponse
    Dim objQBResponse As IResponse
    Dim objQBInvoiceAdd As IInvoiceAdd
    Dim objQBORInvoiceLineAdd As IORInvoiceLineAdd
    Dim objQBInvoiceLineAdd As IInvoiceLineAdd
    Dim objAddress As Object
    Dim blnItemExistFlag As Boolean
    
                    'Get the customer object for the customer number passed in.
    Set objSalesInvoice = CreateObject("converge_sales.sales_invoice")
    With objSalesInvoice
        .p_sales_inv_id = v_varSalesInvId
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .GetSalesInvoice
    End With
    
    If IsNull(objSalesInvoice.p_sales_inv_id) = True Then
        Call Err.Raise(1, , "sales invoice: " & v_varSalesInvId & " not found")
    End If
    
    
    Set objShipment = CreateObject("converge_sales.shipment")
    With objShipment
        .p_ship_id = objSalesInvoice.p_ship_id
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .GetShipment
    End With
    
    If IsNull(objShipment.p_ship_id) = True Then
        Call Err.Raise(1, , "Shipment: " & objSalesInvoice.p_ship_id & " from invoice: " & v_varSalesInvId & " not found")
    End If
    
    Set objSalesOrder = CreateObject("converge_sales.sales_order")
    With objSalesOrder
        .p_sales_ord_id = objShipment.p_sales_ord_id
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .getSalesOrder
    End With
        
    If IsNull(objSalesOrder.p_sales_ord_id) = True Then
        Call Err.Raise(1, , "sales order: " & objSalesInvoice.p_sales_ord_id & " from invoice: " & v_varSalesInvId & " not found")
    End If
    
    Call checkShipMethod(r_objError, r_objTransData, r_objQBSessionMgr, objShipment.p_carrier_nm)
                    'Prepare for quickbooks update or add
    Set objQBRequestMsgSet = Nothing
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
                    'Build Invoice QB Transaction
    Set objQBInvoiceAdd = objQBRequestMsgSet.AppendInvoiceAddRq
    With objQBInvoiceAdd
        Call .ARAccountRef.FullName.SetValue("Accounts Receivable")
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice15"
        Call .CustomerRef.FullName.SetValue(objSalesOrder.p_cust_nbr)
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice16"
        Call .FOB.SetValue(Mid(objSalesOrder.p_fob_terms_desc, 1, .FOB.GetMaxLength))
        Call .IsPending.SetValue(False)
        Call .IsToBePrinted.SetValue(True)
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice17"
        If objSalesOrder.p_drop_ship_fl = "N" Then
            Call .Memo.SetValue("P/L Number: " & objSalesInvoice.p_ship_id & "  S/O: " & objSalesOrder.p_sales_ord_id)
        Else
            Call .Memo.SetValue("S/O: " & objSalesOrder.p_sales_ord_id)
        End If
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice18"
        If IsNull(objSalesOrder.p_cust_po_nbr) Then
            objSalesOrder.p_cust_po_nbr = ""
        End If
        Call .PONumber.SetValue(Mid(objSalesOrder.p_cust_po_nbr, 1, .PONumber.GetMaxLength))
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice19"
        Call .RefNumber.SetValue(v_varSalesInvId)
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice205"
        If IsNull(objSalesOrder.p_sales_rep_assoc_nbr) = True Then
            Call Err.Raise(4, , "Sales Rep is missing from sales order: " & objSalesOrder.p_sales_ord_id)
        End If
        If Len(objSalesOrder.p_sales_rep_assoc_nbr) < 1 Then
            Call Err.Raise(4, , "Sales Rep is missing from sales order.")
        End If
        Call .SalesRepRef.FullName.SetValue(Mid(objSalesOrder.p_sales_rep_assoc_nbr, 1, .SalesRepRef.FullName.GetMaxLength))

        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice21"
        Call .ShipDate.SetValue(objShipment.p_shipped_dt)
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice22"
        Call .ShipMethodRef.FullName.SetValue(Mid(objShipment.p_carrier_nm, 1, .ShipMethodRef.FullName.GetMaxLength))
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice23"
        Call .TermsRef.FullName.SetValue(Mid(objSalesOrder.p_payment_terms_desc, 1, .TermsRef.FullName.GetMaxLength))
        strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncInvoice24"
        Call .TxnDate.SetValue(Date)
'        Call .ClassRef.FullName.SetValue ("")
'        Call .CustomerMsgRef.FullName.SetValue("10308")
'        Call .CustomerSalesTaxCodeRef.FullName.SetValue("N")
'        Call .DueDate.setvalue("")
'        Call .ItemSalesTaxRef.FullName.SetValue("N")
    End With
        
    With objQBInvoiceAdd.BillAddress
        Set objAddress = objSalesOrder.o_bill_address
        Call .Addr1.SetValue(Mid(objSalesOrder.o_bill_name.p_nm, 1, .Addr1.GetMaxLength))
        Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
        Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr3.GetMaxLength))
        Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
        Call .State.SetValue(objAddress.p_State_cd)
        Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
    End With
        
    Dim strShipNm As String
    With objQBInvoiceAdd.ShipAddress
        Set objAddress = objSalesOrder.o_ship_address
        strShipNm = objSalesOrder.o_ship_name.p_nm
        Call .Addr1.SetValue(Mid(objSalesOrder.o_ship_name.p_nm, 1, .Addr1.GetMaxLength))
        Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
        Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr3.GetMaxLength))
        Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
        Call .State.SetValue(objAddress.p_State_cd)
        Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
    End With
    
    Set objSalesInvoiceItem = CreateObject("converge_sales.sales_invoice_item")
    With objSalesInvoiceItem
        .p_sales_inv_id = v_varSalesInvId
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .getSalesInvoiceItems
    End With
    
    Call objSalesInvoiceItem.getSalesInvoiceItems
    Set rstInvoiceItems = objSalesInvoiceItem.o_record_set
    
    blnItemExistFlag = False
    Do While rstInvoiceItems.EOF = False
        If rstInvoiceItems("item_qty") > 0 Then
            blnItemExistFlag = True

            Set objQBORInvoiceLineAdd = objQBInvoiceAdd.ORInvoiceLineAddList.Append
            Set objQBInvoiceLineAdd = objQBORInvoiceLineAdd.InvoiceLineAdd
                
            With objQBInvoiceLineAdd
                Call .ORRatePriceLevel.Rate.SetValue(rstInvoiceItems("item_price"))
                Call .Quantity.SetValue(rstInvoiceItems("item_qty"))
                
                If Len(rstInvoiceItems("cust_item_nbr")) > 0 Then
                    strItemNbr = rstInvoiceItems("cust_item_nbr")
                Else
                    strItemNbr = rstInvoiceItems("item_nbr")
                End If
                Call checkItem(r_objError, r_objTransData, r_objQBSessionMgr, strItemNbr, rstInvoiceItems("item_desc"), rstInvoiceItems("item_price"))
                Call .ItemRef.FullName.SetValue(strItemNbr)
                Call .Desc.SetValue(Mid(rstInvoiceItems("item_desc"), 1, .Desc.GetMaxLength))
            End With
        End If
        Call rstInvoiceItems.MoveNext
    Loop
    
    Call rstInvoiceItems.Close
    
    If blnItemExistFlag = True Then
        Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
        Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
        If objQBResponse.StatusCode <> 0 Then
            Call Err.Raise(4, , objQBResponse.StatusMessage)
        End If
    End If
    
    Exit Sub
ErrHandler:
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub

Public Sub syncCust( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_strCustId As Variant _
                )
    
    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncCust"

    Dim objCustomer As Object
    Dim objQBRequestMsgSet As IMsgSetRequest
    Dim objQbCustomerQuery As ICustomerQuery
    Dim objQBResponseMsgSet As IMsgSetResponse
    Dim objQBResponse As IResponse
    Dim objQBCustomerList As ICustomerRetList
    Dim objQBCustomer As ICustomerRet
    Dim objQBCustMod As ICustomerMod
    Dim objQBCustAdd As ICustomerAdd
    Dim objAddress As Object
    Dim objName As Object
    
                    'Get the customer object for the customer number passed in.
    Set objCustomer = CreateObject("converge_sales.customer")
    With objCustomer
        .p_cust_nbr = Null
        .p_cust_id = v_strCustId
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .GetCustomer
    End With
    
    If IsNull(objCustomer.p_cust_id) = True Then
        Call Err.Raise(1, , "customer: " & v_strCustId & " not found")
    End If
    
                    'get the customer from quickbooks
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    Set objQbCustomerQuery = objQBRequestMsgSet.AppendCustomerQueryRq
    Call objQbCustomerQuery.ORCustomerListQuery.FullNameList.Add(CStr(objCustomer.p_cust_nbr))
    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    
    If (objQBResponse.StatusCode <> 0 And objQBResponse.StatusCode <> 500) Then
        Call Err.Raise(3, , objQBResponse.StatusMessage)
    End If
        
                    'Prepare for quickbooks update or add
    Set objQBRequestMsgSet = Nothing
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
                    'If status code = 0, then the customer already exist, update the customer
                    'If status code <> 0, then the customer does not exist, add the customer.
    If objQBResponse.StatusCode = 0 Then
        Set objQBCustomerList = objQBResponse.Detail
        Set objQBCustomer = objQBCustomerList.GetAt(0)
        Set objQBCustMod = objQBRequestMsgSet.AppendCustomerModRq
        With objQBCustMod
            Call .ListID.SetValue(objQBCustomer.ListID.GetValue)
            Call .Name.SetValue(objCustomer.p_cust_nbr)
            Call .CompanyName.SetValue(Mid(objCustomer.o_stnd_name.p_nm, 1, .CompanyName.GetMaxLength))
            Call .FirstName.SetEmpty
            Call .MiddleName.SetEmpty
            Call .LastName.SetEmpty
            If IsNull(objCustomer.p_phone_nbr) = False Then
                Call .Phone.SetValue(Mid(objCustomer.p_phone_nbr, 1, .Phone.GetMaxLength))
            Else
                Call .Phone.SetEmpty
            End If
            If IsNull(objCustomer.p_fax_nbr) = False Then
                Call .Fax.SetValue(Mid(objCustomer.p_fax_nbr, 1, .Fax.GetMaxLength))
            Else
                Call .Fax.SetEmpty
            End If
            If IsNull(objCustomer.p_email_addr) = False Then
                Call .Email.SetValue(Mid(objCustomer.p_email_addr, 1, .Email.GetMaxLength))
                
            Else
                Call .Email.SetEmpty
            End If
            Call .EditSequence.SetValue(objQBCustomer.EditSequence.GetValue)
        End With
        
        If IsNull(objCustomer.o_bill_name.p_nm_id) Then
            Set objName = objCustomer.o_stnd_name
        Else
            Set objName = objCustomer.o_bill_name
        End If
        If IsNull(objCustomer.o_bill_address.p_addr_id) Then
            Set objAddress = objCustomer.o_stnd_address
        Else
            Set objAddress = objCustomer.o_bill_address
        End If
        With objQBCustMod.BillAddress
            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
            Call .State.SetValue(objAddress.p_State_cd)
            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
        End With
        If IsNull(objCustomer.o_ship_name.p_nm_id) Then
            Set objName = objCustomer.o_stnd_name
        Else
            Set objName = objCustomer.o_ship_name
        End If
        If IsNull(objCustomer.o_ship_address.p_addr_id) Then
            Set objAddress = objCustomer.o_stnd_address
        Else
            Set objAddress = objCustomer.o_ship_address
        End If
        With objQBCustMod.ShipAddress
            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
            Call .State.SetValue(objAddress.p_State_cd)
            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
        End With
    
    Else
                    'add customer
        Set objQBCustAdd = objQBRequestMsgSet.AppendCustomerAddRq
        With objQBCustAdd
            Call .Name.SetValue(objCustomer.p_cust_nbr)
            Call .CompanyName.SetValue(Mid(objCustomer.o_stnd_name.p_nm, 1, .CompanyName.GetMaxLength))
            If IsNull(objCustomer.p_phone_nbr) = False Then
                Call .Phone.SetValue(Mid(objCustomer.p_phone_nbr, 1, .Phone.GetMaxLength))
            Else
                Call .Phone.SetEmpty
            End If
            If IsNull(objCustomer.p_fax_nbr) = False Then
                Call .Fax.SetValue(Mid(objCustomer.p_fax_nbr, 1, .Fax.GetMaxLength))
            Else
                Call .Fax.SetEmpty
            End If
            If IsNull(objCustomer.p_email_addr) = False Then
                Call .Email.SetValue(Mid(objCustomer.p_email_addr, 1, .Email.GetMaxLength))
            Else
                Call .Email.SetEmpty
            End If
        End With
        
        If IsNull(objCustomer.o_bill_name.p_nm_id) Then
            Set objName = objCustomer.o_stnd_name
        Else
            Set objName = objCustomer.o_bill_name
        End If
        If IsNull(objCustomer.o_bill_address.p_addr_id) Then
            Set objAddress = objCustomer.o_stnd_address
        Else
            Set objAddress = objCustomer.o_bill_address
        End If
        With objQBCustAdd.BillAddress
            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
            Call .State.SetValue(objAddress.p_State_cd)
            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
        End With
        
        If IsNull(objCustomer.o_ship_name.p_nm_id) Then
            Set objName = objCustomer.o_stnd_name
        Else
            Set objName = objCustomer.o_ship_name
        End If
        If IsNull(objCustomer.o_ship_address.p_addr_id) Then
            Set objAddress = objCustomer.o_stnd_address
        Else
            Set objAddress = objCustomer.o_ship_address
        End If
        With objQBCustAdd.ShipAddress
            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
            Call .State.SetValue(objAddress.p_State_cd)
            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
        End With
    End If
    
    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    If objQBResponse.StatusCode <> 0 Then
        Call Err.Raise(4, , objQBResponse.StatusMessage)
    End If
    
    Exit Sub
ErrHandler:
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Public Sub syncSupplier( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_strSupplierId As Variant _
                )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncSupplier"

    Dim objSupplier As Object
    Dim objQBRequestMsgSet As IMsgSetRequest
    Dim objQBVendorQuery As IVendorQuery
    Dim objQBResponseMsgSet As IMsgSetResponse
    Dim objQBResponse As IResponse
    Dim objQBVendorList As IVendorRetList
    Dim objQBVendor As IVendorRet
    Dim objQBVendorMod As IVendorMod
    Dim objQBVendorAdd As IVendorAdd
    Dim objAddress As Object
    Dim objName As Object
    
                    'Get the customer object for the customer number passed in.
    Set objSupplier = CreateObject("converge_purchase.supplier")
    With objSupplier
        .p_supplier_id = v_strSupplierId
        .p_supplier_nbr = Null
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .GetSupplier
    End With
    
    If IsNull(objSupplier.p_supplier_id) = True Then
        Call Err.Raise(1, , "supplier: " & v_strSupplierId & " not found")
    End If
    
                    'get the customer from quickbooks
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    Set objQBVendorQuery = objQBRequestMsgSet.AppendVendorQueryRq
    Call objQBVendorQuery.ORVendorListQuery.FullNameList.Add(CStr(objSupplier.p_supplier_nbr))
    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    
    If (objQBResponse.StatusCode <> 0 And objQBResponse.StatusCode <> 500) Then
        Call Err.Raise(3, , objQBResponse.StatusMessage)
    End If
        
                    'Prepare for quickbooks update or add
    Set objQBRequestMsgSet = Nothing
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
                    'If status code = 0, then the vendor already exist, update the vendor
                    'If status code <> 0, then the vendor does not exist, add the vendor.
    If objQBResponse.StatusCode = 0 Then
                        'Currently, Timco does not want to overlay supplier information quickbooks
                        'so, only send in adds
'''        Set objQBVendorList = objQBResponse.Detail
'''        Set objQBVendor = objQBVendorList.GetAt(0)
'''        Set objQBVendorMod = objQBRequestMsgSet.AppendVendorModRq
'''        With objQBVendorMod
'''            Call .ListID.SetValue(objQBVendor.ListID.GetValue)
'''            Call .Name.SetValue(objSupplier.p_supplier_nbr)
'''            Call .CompanyName.SetValue(objSupplier.o_stnd_name.p_nm)
'''            Call .FirstName.SetEmpty
'''            Call .MiddleName.SetEmpty
'''            Call .LastName.SetEmpty
'''            If IsNull(objSupplier.p_phone_nbr) = False Then
'''                Call .Phone.SetValue(objSupplier.p_phone_nbr)
'''            Else
'''                Call .Phone.SetEmpty
'''            End If
'''            If IsNull(objSupplier.p_fax_nbr) = False Then
'''                Call .Fax.SetValue(objSupplier.p_fax_nbr)
'''            Else
'''                Call .Fax.SetEmpty
'''            End If
'''            If IsNull(objSupplier.p_email_addr) = False Then
'''                Call .Email.SetValue(objSupplier.p_email_addr)
'''            Else
'''                Call .Email.SetEmpty
'''            End If
'''            Call .EditSequence.SetValue(objQBVendor.EditSequence.GetValue)
'''        End With
'''        Set objName = objSupplier.o_stnd_name
'''        Set objAddress = objSupplier.o_stnd_address
'''        With objQBVendorMod.VendorAddress
'''            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
'''            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
'''            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
'''            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
'''            Call .State.SetValue(objAddress.p_State_cd)
'''            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
'''        End With
    Else
                    'add customer
        Set objQBVendorAdd = objQBRequestMsgSet.AppendVendorAddRq
        With objQBVendorAdd
            Call .Name.SetValue(objSupplier.p_supplier_nbr)
            Call .CompanyName.SetValue(Mid(objSupplier.o_stnd_name.p_nm, 1, .CompanyName.GetMaxLength))
            Call .FirstName.SetEmpty
            Call .MiddleName.SetEmpty
            Call .LastName.SetEmpty
            If IsNull(objSupplier.p_phone_nbr) = False Then
                Call .Phone.SetValue(Mid(objSupplier.p_phone_nbr, 1, .Phone.GetMaxLength))
            Else
                Call .Phone.SetEmpty
            End If
            If IsNull(objSupplier.p_fax_nbr) = False Then
                Call .Fax.SetValue(Mid(objSupplier.p_fax_nbr, 1, .Fax.GetMaxLength))
            Else
                Call .Fax.SetEmpty
            End If
            If IsNull(objSupplier.p_email_addr) = False Then
                Call .Email.SetValue(Mid(objSupplier.p_email_addr, 1, .Email.GetMaxLength))
            Else
                Call .Email.SetEmpty
            End If
        End With
        
        Set objName = objSupplier.o_stnd_name
        Set objAddress = objSupplier.o_stnd_address
        With objQBVendorAdd.VendorAddress
            Call .Addr1.SetValue(Mid(objName.p_nm, 1, .Addr1.GetMaxLength))
            Call .Addr2.SetValue(Mid(getAddr(objAddress.p_Street_addr, 1), 1, .Addr2.GetMaxLength))
            Call .Addr3.SetValue(Mid(getAddr(objAddress.p_Street_addr, 2), 1, .Addr2.GetMaxLength))
            Call .City.SetValue(Mid(objAddress.p_City_nm, 1, .City.GetMaxLength))
            Call .State.SetValue(objAddress.p_State_cd)
            Call .PostalCode.SetValue(Mid(objAddress.p_Zip_cd, 1, .PostalCode.GetMaxLength))
        End With
                        'Currently, Timco does not want to overlay supplier information quickbooks
                        'so, only send in adds
        Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
        Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
        If objQBResponse.StatusCode <> 0 Then
            Call Err.Raise(4, , objQBResponse.StatusMessage)
        End If
    End If
    
    Exit Sub
ErrHandler:
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Public Sub syncItem( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_strItemId As Variant _
                )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "syncItem"

    Dim objItem As Object
    Dim objQBRequestMsgSet As IMsgSetRequest
    Dim objQBItemQuery As IItemQuery
    Dim objQBResponseMsgSet As IMsgSetResponse
    Dim objQBResponse As IResponse
    Dim objQBOrItemList As IORItemRetList
    Dim objQBOrItem As IORItemRet
    Dim objQBItemServiceMod As IItemServiceMod
    Dim objQBItemServiceAdd As IItemServiceAdd
    Dim objQBItemNonInventoryMod As IItemNonInventoryMod
    Dim objQBItemNonInventoryAdd As IItemNonInventoryAdd
    Dim objQBItemInventoryMod As IItemInventoryMod
    Dim objQBItemInventoryAdd As IItemInventoryAdd
    
                    'Get the customer object for the customer number passed in.
    Set objItem = CreateObject("converge_item.item")
    With objItem
        .p_Item_id = v_strItemId
        .p_Item_nbr = Null
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        Call .GetItem
    End With
    
    If IsNull(objItem.p_Item_id) = True Then
        Call Err.Raise(1, , "Item: " & v_strItemId & " not found")
    End If
                    'get the customer from quickbooks
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    Set objQBItemQuery = objQBRequestMsgSet.AppendItemQueryRq
    Call objQBItemQuery.ORListQuery.FullNameList.Add(CStr(objItem.p_Item_nbr))
    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    
    If (objQBResponse.StatusCode <> 0 And objQBResponse.StatusCode <> 500) Then
        Call Err.Raise(3, , objQBResponse.StatusMessage)
    End If
                    'Prepare for quickbooks update or add
    Set objQBRequestMsgSet = Nothing
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
                    'If status code = 0, then the vendor already exist, update the vendor
                    'If status code <> 0, then the vendor does not exist, add the vendor.
    If objQBResponse.StatusCode = 0 Then
'''        Set objQBOrItemList = objQBResponse.Detail
'''        Set objQBOrItem = objQBOrItemList.GetAt(0)
'''
'''        Set objQBItemServiceMod = objQBRequestMsgSet.AppendItemServiceModRq
'''        With objQBItemServiceMod
'''            Call .ListID.SetValue(objQBOrItem.ItemServiceRet.ListID.GetValue)
'''            Call .Name.SetValue(objItem.p_Item_nbr)
'''            Call .SalesTaxCodeRef.FullName.SetValue("Non")
'''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.Desc.SetValue(objItem.p_item_desc)
'''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.ORPrice.Price.SetValue(objItem.p_item_price)
'''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.AccountRef.FullName.SetValue("Sales")
'''            Call .EditSequence.SetValue(objQBOrItem.ItemServiceRet.EditSequence.GetValue)
'''        End With
'''
''''        Set objQBItemNonInventoryMod = objQBRequestMsgSet.AppendItemNonInventoryModRq
''''        With objQBItemNonInventoryMod
''''            Call .ListID.SetValue(objQBOrItem.ItemNonInventoryRet.ListID.GetValue)
''''            Call .Name.SetValue(objItem.p_item_nbr)
''''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.Desc.SetValue(objItem.p_item_desc)
''''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.ORPrice.Price.SetValue(objItem.p_item_price)
''''            Call .ORSalesPurchaseMod.SalesOrPurchaseMod.AccountRef.FullName.SetValue("Construction")
''''            Call .EditSequence.SetValue(objQBOrItem.ItemNonInventoryRet.EditSequence.GetValue)
''''        End With
'''
''''        Set objQBItemInventoryMod = objQBRequestMsgSet.AppendItemInventoryModRq
''''        With objQBItemInventoryMod
''''            Call .ListID.SetValue(objQBOrItem.ItemInventoryRet.ListID.GetValue)
''''            Call .Name.SetValue(objItem.p_item_nbr)
''''            Call .PrefVendorRef.FullName.SetValue("ABBEY")
''''            Call .PurchaseCost.SetValue(objItem.p_item_cost)
''''            Call .PurchaseDesc.SetValue(objItem.p_item_desc)
''''            Call .SalesPrice.SetValue(objItem.p_item_price)
''''            Call .SalesDesc.SetValue(objItem.p_item_desc)
''''            Call .SalesTaxCodeRef.FullName.SetValue("Non")
''''            Call .IncomeAccountRef.FullName.SetValue("Uncategorized Income")
''''            Call .COGSAccountRef.FullName.SetValue("Cost Of Goods Sold")
''''            Call .AssetAccountRef.FullName.SetValue("Inventory Asset")
''''            Call .EditSequence.SetValue(objQBOrItem.ItemInventoryRet.EditSequence.GetValue)
''''        End With
'''
    Else
                    'add Item
        Set objQBItemServiceAdd = objQBRequestMsgSet.AppendItemServiceAddRq
        With objQBItemServiceAdd
            Call .Name.SetValue(objItem.p_Item_nbr)
            Call .SalesTaxCodeRef.FullName.SetValue("Non")
            Call .ORSalesPurchase.SalesOrPurchase.Desc.SetValue(Mid(objItem.p_item_desc, 1, .ORSalesPurchase.SalesOrPurchase.Desc.GetMaxLength))
            Call .ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(objItem.p_item_price)
            Call .ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue("Sales")
        End With
    
        Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
        Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
        If objQBResponse.StatusCode <> 0 Then
            Call Err.Raise(4, , objQBResponse.StatusMessage)
        End If
'        Set objQBItemNonInventoryAdd = objQBRequestMsgSet.AppendItemNonInventoryAddRq
'        With objQBItemNonInventoryAdd
'            Call .Name.SetValue(objItem.p_item_nbr)
'            Call .ORSalesPurchase.SalesOrPurchase.Desc.SetValue(objItem.p_item_desc)
'            Call .ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(objItem.p_item_price)
'            Call .ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue("Construction")
'        End With

'        Set objQBItemInventoryAdd = objQBRequestMsgSet.AppendItemInventoryAddRq
'        With objQBItemInventoryAdd
'            Call .Name.SetValue(objItem.p_item_nbr)
'            Call .PrefVendorRef.FullName.SetValue("ABBEY")
'            Call .PurchaseCost.SetValue(objItem.p_item_cost)
'            Call .PurchaseDesc.SetValue(objItem.p_item_desc)
'            Call .SalesPrice.SetValue(objItem.p_item_price)
'            Call .SalesDesc.SetValue(objItem.p_item_desc)
'            Call .SalesTaxCodeRef.FullName.SetValue("Non")
'            Call .IncomeAccountRef.FullName.SetValue("Uncategorized Income")
'            Call .COGSAccountRef.FullName.SetValue("Cost Of Goods Sold")
'            Call .AssetAccountRef.FullName.SetValue("Inventory Asset")
'        End With
        
    End If
    
    Exit Sub
ErrHandler:
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Public Sub checkItem( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_strItemNbr As Variant _
                , ByVal v_varItemDesc As Variant _
                , ByVal v_varItemPrice As Variant _
                )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "checkItem"

    Dim objQBRequestMsgSet As IMsgSetRequest
    Dim objQBItemQuery As IItemQuery
    Dim objQBResponseMsgSet As IMsgSetResponse
    Dim objQBResponse As IResponse
    Dim objQBOrItemList As IORItemRetList
    Dim objQBOrItem As IORItemRet
    Dim objQBItemServiceMod As IItemServiceMod
    Dim objQBItemServiceAdd As IItemServiceAdd
    Dim objQBItemNonInventoryMod As IItemNonInventoryMod
    Dim objQBItemNonInventoryAdd As IItemNonInventoryAdd
    Dim objQBItemInventoryMod As IItemInventoryMod
    Dim objQBItemInventoryAdd As IItemInventoryAdd
    
                    'get the customer from quickbooks
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    Set objQBItemQuery = objQBRequestMsgSet.AppendItemQueryRq
    Call objQBItemQuery.ORListQuery.FullNameList.Add(v_strItemNbr)
    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    
    If (objQBResponse.StatusCode <> 0 And objQBResponse.StatusCode <> 500) Then
        Call Err.Raise(3, , objQBResponse.StatusMessage)
    End If
        
                    'Prepare for quickbooks update or add
    Set objQBRequestMsgSet = Nothing
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
                    'If status code = 0, then the vendor already exist, update the vendor
                    'If status code <> 0, then the vendor does not exist, add the vendor.
    If objQBResponse.StatusCode = 0 Then
    Else
        Set objQBItemServiceAdd = objQBRequestMsgSet.AppendItemServiceAddRq
        With objQBItemServiceAdd
            If Len(v_strItemNbr) > CLng(.Name.GetMaxLength) Then
                Debug.Print .Name.GetMaxLength & " " & Len(v_strItemNbr)
                Debug.Print .Name.GetMaxLength > Len(v_strItemNbr)
                Call Err.Raise(4, , "Item Nbr (" & v_strItemNbr & ") is greater than " & .Name.GetMaxLength & " characters long.  It can not be loaded")
            End If
            Call .Name.SetValue(v_strItemNbr)
            Call .SalesTaxCodeRef.FullName.SetValue("Non")
            Call .ORSalesPurchase.SalesOrPurchase.Desc.SetValue(Mid(v_varItemDesc, .ORSalesPurchase.SalesOrPurchase.Desc.GetMaxLength))
            Call .ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(v_varItemPrice)
            Call .ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue("Sales")
        End With
    
        Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
        Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
        If objQBResponse.StatusCode <> 0 Then
            Call Err.Raise(4, , objQBResponse.StatusMessage)
        End If
    End If
    
    Exit Sub
ErrHandler:
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub


Public Sub checkShipMethod( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByRef r_objQBSessionMgr As QBSessionManager _
                , ByVal v_strShipMethod As Variant _
                )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = G_STR_PROJECT_NM & "." & M_STR_MODULE_NM & "." & "checkShipMethod"

    Dim objQBRequestMsgSet As IMsgSetRequest _
        , objQBResponseMsgSet As IMsgSetResponse _
        , objQBResponse As IResponse _
        , objQBShipMethodAdd As IShipMethodAdd
    
    Set objQBRequestMsgSet = r_objQBSessionMgr.CreateMsgSetRequest(M_STR_QB_XML_COUNTRY_CD, M_INT_QB_XML_MAJOR_VERSION, M_INT_QB_XML_MINOR_VERSION)
    objQBRequestMsgSet.Attributes.OnError = roeContinue
    
    Set objQBShipMethodAdd = objQBRequestMsgSet.AppendShipMethodAddRq
    With objQBShipMethodAdd
        Call .IsActive.SetValue(True)
        Call .Name.SetValue(Mid(v_strShipMethod, 1, .Name.GetMaxLength))
    End With

    Set objQBResponseMsgSet = r_objQBSessionMgr.DoRequests(objQBRequestMsgSet)
    Set objQBResponse = objQBResponseMsgSet.ResponseList.GetAt(0)
    If objQBResponse.StatusCode <> 0 And objQBResponse.StatusCode <> 3100 Then
        Call Err.Raise(4, , objQBResponse.StatusMessage)
    End If
    
    Exit Sub
ErrHandler:
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
    Else
        With r_objError
          .p_type_cd = "F"
          .p_err_cd = "0600"
          .p_nbr = Err.Number
          .p_desc = Err.Description
          .p_routine_nm = strRoutineNm
          .p_message_id = 0
        End With
    End If
    Call Err.Raise(1)
End Sub



Private Function getAddr(ByVal v_varAddr As Variant, v_lngLineNo As Long) As Variant
  
    Dim lngCurrLineNo As Long
    Dim lngCurrPosNo As Long
    Dim strCharacter As String
    
    getAddr = ""
    
    If Len(v_varAddr) < 1 Then
       Exit Function
    End If

    lngCurrPosNo = 1
    lngCurrLineNo = 1
    Do Until lngCurrLineNo > v_lngLineNo Or lngCurrPosNo > Len(v_varAddr)
        strCharacter = Mid(v_varAddr, lngCurrPosNo, 2)
        Do Until strCharacter = vbCrLf Or lngCurrPosNo > Len(v_varAddr)
            If lngCurrLineNo = v_lngLineNo Then
                getAddr = getAddr & Mid(strCharacter, 1, 1)
            End If
            lngCurrPosNo = lngCurrPosNo + 1
            strCharacter = Mid(v_varAddr, lngCurrPosNo, 2)
        Loop
        lngCurrPosNo = lngCurrPosNo + 2
        lngCurrLineNo = lngCurrLineNo + 1
    Loop
  
End Function

Private Sub Class_Terminate()
    If m_blnSessionOpenFl = True Then
        m_objQBSessionMgr.EndSession
        m_objQBSessionMgr.CloseConnection
        m_blnSessionOpenFl = False
    End If
End Sub
