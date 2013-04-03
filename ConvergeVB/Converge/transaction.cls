VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_trans_object As Object
Private m_conLogging As ADODB.Connection
Private m_blnLoggingFl As Boolean

Public Sub setLoggingFl(ByVal v_blnLoggingFl As Boolean)
    m_blnLoggingFl = v_blnLoggingFl
End Sub


Public Property Set o_trans_object(ByVal vData As Object)
    Set mvaro_trans_object = vData
End Property
Public Property Get o_trans_object() As Object
    Set o_trans_object = mvaro_trans_object
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

Public Sub ProcessTrans(Optional ByRef v_objTransProps As Variant)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm   As String
    Dim strSql
    Dim rstData         As ADODB.Recordset
    Dim strStartTs As String
    strRoutine_nm = "transaction.cls.ProcessTrans"
    
    If m_blnLoggingFl = True Then
        strStartTs = getStartTimestamp()
    End If
    
    Call o_trans_data.OpenDatabase(o_error)
    
    If o_trans_data.update_fl = True Then
        Call o_trans_data.BeginTransDatabase(o_error)
    End If
    
    Set o_trans_object.o_error = o_error
    Set o_trans_object.o_trans_data = o_trans_data
    
    Select Case o_trans_data.p_Function_cd
     
       ' Get Section
         Case "GetItem"
             Call o_trans_object.GetItem
        Case "GetCustomer"
            Call o_trans_object.GetCustomer
        Case "GetAssemblyOrder"
            Call o_trans_object.GetAssemblyOrder
        Case "GetLot"
            Call o_trans_object.GetLot
        Case "GetShipment"
            Call o_trans_object.GetShipment
        Case "GetInventoryAdj"
            Call o_trans_object.GetInventoryAdj
        Case "GetPurchaseOrder"
            Call o_trans_object.GetPurchaseOrder
        Case "GetSupplier"
            Call o_trans_object.GetSupplier
    
    '' Production ''
      ' Assembly Order Information
        Case "AddAssemblyOrder"
            Call o_trans_object.AddAssemblyOrder
        Case "DeleteAssemblyOrder"
            Call o_trans_object.DeleteAssemblyOrder
        Case "ChangeAssemblyOrder"
            Call o_trans_object.ChangeAssemblyOrder
        Case "ChangeAssemblyOrderStatus"
            Call o_trans_object.ChangeStatus
        Case "GetAssemblyOrders"
            Call o_trans_object.getAssemblyOrders
        Case "GenerateAssemblyOrderPlan"
            Call o_trans_object.generatePlan
        Case "GetPastDueAssemblyOrders"
            Call o_trans_object.getPastDueAssemblyOrders
        Case "GetNonPastDueAssemblyOrders"
            Call o_trans_object.getNonPastDueAssemblyOrders
        Case "GetAssemblyOrderComps"
            Call o_trans_object.GetAssemblyOrderComps
        Case "GetAssemblyComponent"
            Call o_trans_object.GetAssemblyComponent
        Case "OpenOrder"
            Call o_trans_object.openOrder
        Case "GetMatReq"
            Call o_trans_object.GetMatReq
        Case "closeMatReq"
            Call o_trans_object.closeMatReqAsp
        Case "openMatReq"
            Call o_trans_object.openMatReqAsp
'        Case "addMatReq"
'            Call o_trans_object.addMatReq( _
'                              Me.o_error _
'                            , Me.o_trans_data _
'                            , v_objTransProps.getProp("strOrderTypeCd") _
'                            , v_objTransProps.getProp("lngSalesOrdId") _
'                            , v_objTransProps.getProp("lngSoLineNbr") _
'                            , v_objTransProps.getProp("lngAssemOrdId") _
'                            , v_objTransProps.getProp("lngAcLineNbr") _
'                            , v_objTransProps.getProp("lngItemId") _
'                            , v_objTransProps.getProp("strInvLocId") _
'                            , v_objTransProps.getProp("lngOrderedQty") _
'                            , v_objTransProps.getProp("dteReqDelDt") _
'                            , v_objTransProps.getProp("dblRequiredQtyVarPct") _
'                            , v_objTransProps.getProp("dblAllocPriorityNbr") _
'                            , v_objTransProps.getProp("lngSupplierId") _
'                            )
            
        Case "GetMatReqs"
            Call o_trans_object.GetMatReqs
        Case "GeneratePlan"
            Call o_trans_object.generatePlan
        Case "ModifySupplier"
            Call o_trans_object.ModifySupplier
        Case "GetProductionRuns"
            Call o_trans_object.GetProductionRuns
        Case "GetProductionRun"
            Call o_trans_object.GetProductionRun
        Case "AddProductionRun"
            Call o_trans_object.AddProductionRun
        Case "ChangeProductionRun"
            Call o_trans_object.ChangeProductionRun
        Case "DeleteProductionRun"
            Call o_trans_object.DeleteProductionRun
        Case "OpenProductionRun"
            Call o_trans_object.OpenProductionRun
        Case "ChangeProductionRunStatus"
            Call o_trans_object.ChangeStatus

            
    'Bill of Material
        Case "BOMExplosion_MultiLevel"
            Call o_trans_object.BOMExplosion_MultiLevel
        Case "BOMExplosion_SingleLevel"
            Call o_trans_object.BOMExplosion_SingleLevel
        Case "AddToBill"
            Call o_trans_object.AddToBill
        Case "ModifyBill"
            Call o_trans_object.ModifyBill
        Case "RemoveComponent"
            Call o_trans_object.RemoveComponent
        Case "DeleteComponents"
            Call o_trans_object.DeleteComponents
    
    ''Inventory ''
        ' Bin
        Case "GetBin"
            Call o_trans_object.GetInventoryBin
        Case "AddBin"
            Call o_trans_object.Add
        Case "DeleteBin"
            Call o_trans_object.Delete
        Case "ChangeBin"
            Call o_trans_object.Change
        Case "GetInventoryBins"
            Call o_trans_object.GetInventoryBins
            
        ' Lot
        
        Case "AddLot"
            Call o_trans_object.Add
        Case "DeleteLot"
            Call o_trans_object.Delete
        Case "ChangeLot"
            Call o_trans_object.Change
        Case "GetLots"
            Call o_trans_object.GetLots
        Case "GetLotInspects"
            Call o_trans_object.GetLotInspects
        Case "GetCompLots"
            Call o_trans_object.GetCompLots
        Case "GetLotInspectsForItem"
            Call o_trans_object.GetLotInspectsForItem
            
        ' Supplier Lot
        Case "GetSupplierLot"
            Call o_trans_object.GetSupplierLot
        Case "AddSupplierLot"
            Call o_trans_object.Add
        Case "DeleteSupplierLot"
            Call o_trans_object.Delete
        Case "ChangeSupplierLot"
            Call o_trans_object.Change
        Case "GetSupplierLots"
            Call o_trans_object.GetSupplierLots
            
        ' Lot Inspect
        Case "GetLotInspect"
            Call o_trans_object.GetLotInspect
        Case "AddLotInspect"
            Call o_trans_object.Add
        Case "DeleteLotInspect"
            Call o_trans_object.Delete
        Case "ChangeLotInspect"
            Call o_trans_object.Change
            
        ' inspection
        Case "AddLotInspection"
            Call o_trans_object.Add
        Case "ChangeLotInspection"
            Call o_trans_object.Change
        Case "DeleteLotInspection"
            Call o_trans_object.Delete
        Case "GetLotInspection"
            Call o_trans_object.GetInspect
            
            
        ' On Order Inventory Adj
        Case "GetOnOrdInvAdj"
            Call o_trans_object.GetOnOrdInvAdj
        Case "AddOnOrdInvAdj"
            Call o_trans_object.Add
        Case "DeleteOnOrdInvAdj"
            Call o_trans_object.Delete
        Case "ChangeOnOrdInvAdj"
            Call o_trans_object.Change
        Case "GetOnOrderInventoryAdjs"
            Call o_trans_object.GetOnOrderInventoryAdjs
        Case "ProcessOnOrdInvAdj"
            Call o_trans_object.ProcessOnOrdInvAdj
                                            
        ' Inventory Adj
        Case "AddInventoryAdj"
            Call o_trans_object.Add
        Case "DeleteInventoryAdj"
            Call o_trans_object.Delete
        Case "ChangeInventoryAdj"
            Call o_trans_object.Change
        Case "GetInventoryAdjs"
            Call o_trans_object.GetInventoryAdjs
        Case "GetTotalInventoryAdjs"
            Call o_trans_object.GetTotalInventoryAdjs
        Case "GetInventoryItems"
            Call o_trans_object.GetInventoryItems
        Case "GetProdRunInventoryAdjs"
            Call o_trans_object.GetProdRunInventoryAdjs
        Case "GetPORecInventoryAdjs"
            Call o_trans_object.GetPORecInventoryAdjs
        Case "AddPORecSupplierLot"
            Call o_trans_object.AddPORecSupplierLot
        Case "GetPickList"
            Call o_trans_object.GetPickList
        
        ' Inventory Allocation
        Case "GetInventoryAlloc"
            Call o_trans_object.GetInventoryAlloc
        Case "AllocateInventory"
            Call o_trans_object.Allocate
        Case "DeallocateInventory"
            Call o_trans_object.Deallocate
        Case "AdjustAllocation"
            Call o_trans_object.AdjustAllocation
        Case "GetInventoryAllocs"
            Call o_trans_object.GetInventoryAllocs
                
        ' On Order Inventory Allocation
        Case "GetOnOrderInventoryAlloc"
            Call o_trans_object.GetOnOrderInventoryAlloc
        Case "GetOnOrderInventoryAllocs"
            Call o_trans_object.GetOnOrderInventoryAllocs
                        
        Case "PriorityAllocateInventory"
            Call o_trans_object.ReallocateItem
        
        ' Location
        Case "GetDefaultInventoryLoc"
            Call o_trans_object.GetDefaultInventoryLoc
        Case "GetInventoryLoc"
            Call o_trans_object.GetInventoryLoc
        Case "AddInventoryLoc"
            Call o_trans_object.Add
        Case "DeleteInventoryLoc"
            Call o_trans_object.Delete
        Case "ChangeInventoryLoc"
            Call o_trans_object.Change
        Case "GetInventoryLocs"
            Call o_trans_object.GetInventoryLocs
    
        ' Physical Inventory
        Case "GetPhyInv"
            Call o_trans_object.GetPhyInv
        Case "GetPhyInvs"
            Call o_trans_object.GetPhyInvs
        Case "GetPhyInvAdjs"
            Call o_trans_object.GetPhyInvAdjs
        Case "GetTotalInventoryAdj"
            Call o_trans_object.GetQty
        Case "MoveInventoryAdj"
            Call o_trans_object.Move
        Case "AddPhyInv"
            Call o_trans_object.Add
        Case "DeletePhyInv"
            Call o_trans_object.Delete
        Case "ChangePhyInv"
            Call o_trans_object.Change
        Case "ChangePhyInvStatus"
            Call o_trans_object.ChangeStatus
    
    ''PURCHASING ''
        ' Supplier Item
        Case "GetSupplierItem"
            Call o_trans_object.GetSupplierItem
        Case "AddSupplierItem"
            Call o_trans_object.Add
        Case "DeleteSupplierItem"
            Call o_trans_object.Delete
        Case "ChangeSupplierItem"
            Call o_trans_object.Change
        Case "GetSupplierItems"
            Call o_trans_object.getSupplierItems
        Case "GetItemSuppliers"
            Call o_trans_object.getItemSuppliers
        
        
        ' Supplier
        Case "AddSupplier"
            Call o_trans_object.AddSupplier
        Case "DeleteSupplier"
            Call o_trans_object.DeleteSupplier
        Case "ChangeSupplier"
            Call o_trans_object.ChangeSupplier
        Case "GetSuppliers"
            Call o_trans_object.getSuppliers
            
        ' Supplier cost break
        Case "GetCostBreak"
            Call o_trans_object.GetCostBreak
        Case "AddCostBreak"
            Call o_trans_object.Add
        Case "ChangeCostBreak"
            Call o_trans_object.Change
        Case "DeleteCostBreak"
            Call o_trans_object.Delete
        Case "GetCostBreaks"
            Call o_trans_object.GetCostBreaks
        
        ' Purchase Order Rec Information
        Case "GetPurchaseOrderRec"
            Call o_trans_object.GetPurchaseOrderRec
        Case "AddPurchaseOrderRec"
            Call o_trans_object.Add
        Case "DeletePurchaseOrderRec"
            Call o_trans_object.Delete
        Case "ChangePurchaseOrderRec"
            Call o_trans_object.Change
        Case "GetPurchaseOrderRecs"
            Call o_trans_object.GetPurchaseOrderRecs
       
       ' Purchase Order Ret Information
        Case "GetPurchaseOrderRet"
            Call o_trans_object.GetPurchaseOrderRet
        Case "AddPurchaseOrderRet"
            Call o_trans_object.Add
        Case "DeletePurchaseOrderRet"
            Call o_trans_object.Delete
        Case "ChangePurchaseOrderRet"
            Call o_trans_object.Change
        Case "GetPurchaseOrderRets"
            Call o_trans_object.getPurchaseOrderRets
       
        ' Purchase Order History Information
        Case "GetPurchaseOrderHistories"
            Call o_trans_object.getPurchaseOrderHistories
        Case "GetPurchaseOrderHistory"
            Call o_trans_object.getPurchaseOrderHistory
        Case "GetPurchaseOrderItemHistories"
            Call o_trans_object.getPurchaseOrderItemHistories
        Case "GetPurchaseOrderItemHistory"
            Call o_trans_object.getPurchaseOrderItemHistory
       
        ' Purchase Order Information
        Case "AddPurchaseOrder"
            Call o_trans_object.Add
        Case "DeletePurchaseOrder"
            Call o_trans_object.Delete
        Case "ChangePurchaseOrder"
            Call o_trans_object.Change
        Case "ChangePurchaseOrderStatus"
            Call o_trans_object.ChangeStatus
        Case "GetSentPurchaseOrders"
            Call o_trans_object.getSentPurchaseOrders
        Case "GetPurchaseOrders"
            Call o_trans_object.getPurchaseOrders
            
        ' Purchase Order Line Information
        Case "GetPurchaseOrderItem"
            Call o_trans_object.GetPurchaseOrderItem
        Case "AddPurchaseOrderItem"
            Call o_trans_object.Add
        Case "DeletePurchaseOrderItem"
            Call o_trans_object.Delete
        Case "ChangePurchaseOrderItem"
            Call o_trans_object.Change
        Case "ChangePurchaseOrderItemStatus"
            Call o_trans_object.ChangeStatus
        Case "GetPurchaseOrderItems"
            Call o_trans_object.GetPurchaseOrderItems
        
    '' CONVERGE
        ' Message
        Case "GetMessage"
            Call o_trans_object.GetMessage
        Case "AddMessage"
            Call o_trans_object.Add
        Case "DeleteMessage"
            Call o_trans_object.Delete
        Case "ChangeMessage"
            Call o_trans_object.Change
        Case "GetMessages"
            Call o_trans_object.GetMessages
        
        'Code Description
        Case "GetCodeDesc"
            Call o_trans_object.GetCodeDesc
        Case "GetCodeDescs"
            Call o_trans_object.GetCodeDescs
        Case "AddCodeDesc"
            Call o_trans_object.Add
        Case "ChangeCodeDesc"
            Call o_trans_object.Change
        Case "DeleteCodeDesc"
            Call o_trans_object.Delete
       
    '' ITEM
        ' Item Information
         Case "AddItem"
             Call o_trans_object.Add
         Case "DeleteItem"
             Call o_trans_object.Delete
         Case "ChangeItem"
             Call o_trans_object.Change
         Case "GetItems"
             Call o_trans_object.getItems
         Case "GetItemCusts"
             Call o_trans_object.getItemCusts
             
         Case "GetItemInspect"
             Call o_trans_object.GetItem
         Case "AddItemInspect"
             Call o_trans_object.Add
         Case "DeleteItemInspect"
             Call o_trans_object.Delete
         Case "ChangeItemInspect"
             Call o_trans_object.Change
         Case "GetItemInspects"
             Call o_trans_object.GetItemInspects
             
         Case "GetInspection"
             Call o_trans_object.GetInspection
         Case "AddInspection"
             Call o_trans_object.Add
         Case "ChangeInspection"
             Call o_trans_object.Change
         Case "DeleteInspection"
             Call o_trans_object.Delete
         Case "GetInspections"
             Call o_trans_object.GetInspections
             
    '' SALES
       ' Sales Invoice Information
        Case "GetSalesInvoice"
            Call o_trans_object.GetSalesInvoice
        Case "AddSalesInvoice"
            Call o_trans_object.Add
        Case "DeleteSalesInvoice"
            Call o_trans_object.Delete
        Case "ChangeSalesInvoice"
            Call o_trans_object.Change
        Case "GetSalesInvoices"
            Call o_trans_object.getSalesInvoices
            
        ' Sales Invoice Line Information
        Case "GetSalesInvoiceItem"
            Call o_trans_object.GetSalesInvoiceItem
        Case "AddSalesInvoiceItem"
            Call o_trans_object.Add
        Case "DeleteSalesInvoiceItem"
            Call o_trans_object.Delete
        Case "ChangeSalesInvoiceItem"
            Call o_trans_object.Change
        
        ' Customer price break
        Case "GetPriceBreak"
            Call o_trans_object.GetPriceBreak
        Case "AddPriceBreak"
            Call o_trans_object.Add
        Case "ChangePriceBreak"
            Call o_trans_object.Change
        Case "DeletePriceBreak"
            Call o_trans_object.Delete
        Case "GetPriceBreaks"
            Call o_trans_object.GetPriceBreaks
            
        ' Sales Invoice Line Information
        Case "GetSalesInvoiceItems"
            Call o_trans_object.getSalesInvoiceItems
            
        ' Sales Order Information
        Case "GetSalesOrder"
            Call o_trans_object.GetSalesOrder
        Case "AddSalesOrder"
            Call o_trans_object.Add
        Case "AddSubSalesOrder"
            Call o_trans_object.AddSubSalesOrder
        Case "DeleteSalesOrder"
            Call o_trans_object.Delete
        Case "ChangeSalesOrder"
            Call o_trans_object.Change
        Case "ChangeSalesOrderStatus"
            Call o_trans_object.ChangeStatus
        Case "GetCustSalesOrders"
            Call o_trans_object.getCustSalesOrders
        Case "getDropShipSalesOrdersReadytoPrint"
            Call o_trans_object.getDropShipSalesOrdersReadytoPrint
        Case "AddDropShipPo"
            Call o_trans_object.AddDropShipPo
        Case "GenerateSalesOrderPlan"
            Call o_trans_object.generatePlan
            
        ' Sales Order Line Information
        Case "GetSalesOrderLineItem"
            Call o_trans_object.GetSalesOrderLineItem
        Case "AddSalesOrderLineItem"
            Call o_trans_object.Add
        Case "DeleteSalesOrderLineItem"
            Call o_trans_object.Delete
        Case "ChangeSalesOrderLineItem"
            Call o_trans_object.Change
        Case "ChangeSalesOrderLineItemStatus"
            Call o_trans_object.ChangeStatus
            
        ' Sales Order Line Information
        Case "GetSalesOrderLineItems"
            Call o_trans_object.getSalesOrderLineItems
        Case "AddSalesOrderLineItems"
            Call o_trans_object.Add
        Case "DeleteSalesOrderLineItems"
            Call o_trans_object.Delete
        Case "ChangeSalesOrderLineItems"
            Call o_trans_object.Change
        Case "GetSalesOrderItems"
            Call o_trans_object.getSalesOrderItems

        ' Sales Quote Information
        Case "GetSalesQuote"
            Call o_trans_object.GetSalesQuote
        Case "AddSalesQuote"
            Call o_trans_object.Add
        Case "DeleteSalesQuote"
            Call o_trans_object.Delete
        Case "ChangeSalesQuote"
            Call o_trans_object.Change
        Case "GetCustSalesQuotes"
            Call o_trans_object.getCustSalesQuotes
            
        ' Sales Quote Line Information
        Case "GetSalesQuoteItem"
            Call o_trans_object.GetSalesQuoteItem
        Case "AddSalesQuoteItem"
            Call o_trans_object.Add
        Case "DeleteSalesQuoteItem"
            Call o_trans_object.Delete
        Case "ChangeSalesQuoteItem"
            Call o_trans_object.Change
        Case "DeleteSalesQuoteLineItems"
            Call o_trans_object.DeleteLineItems
            
        ' Sales Quote Line Information
        Case "GetSalesQuoteItems"
            Call o_trans_object.getSalesQuoteItems
            
        ' Sales Quote Information
        
        Case "AddShipment"
            Call o_trans_object.Add
        Case "DeleteShipment"
            Call o_trans_object.Delete
        Case "ChangeShipment"
            Call o_trans_object.Change
        Case "ChangeShipmentStatus"
            Call o_trans_object.ChangeStatus
        Case "GetShipments"
            Call o_trans_object.getShipments
        Case "GetShipmentItem"
            Call o_trans_object.GetShipmentItem
        Case "GetShipmentItems"
            Call o_trans_object.GetShipmentItems
        Case "ChangeShipmentItem"
            Call o_trans_object.Update
            
        ' Customer
        Case "AddCustomer"
            Call o_trans_object.AddCustomer
        Case "DeleteCustomer"
            Call o_trans_object.DeleteCustomer
        Case "ChangeCustomer"
            Call o_trans_object.ChangeCustomer
        Case "GetContacts"
            Call o_trans_object.getContacts
        Case "GetContact"
            Call o_trans_object.GetContact
        Case "AddContact"
            Call o_trans_object.AddContact
        Case "DeleteContact"
            Call o_trans_object.DeleteContact
        Case "ChangeContact"
            Call o_trans_object.ChangeContact
        Case "GetAssociate"
            Call o_trans_object.GetAssociate
        Case "AddAssociate"
            Call o_trans_object.AddAssociate
        Case "DeleteAssociate"
            Call o_trans_object.DeleteAssociate
        Case "ChangeAssociate"
            Call o_trans_object.ChangeAssociate
        Case "GetAssociates"
            Call o_trans_object.getAssociates
        Case "GetComment"
            Call o_trans_object.GetComment
        Case "AddComment"
            Call o_trans_object.AddComment
        Case "ChangeComment"
            Call o_trans_object.ChangeComment
        Case "DeleteComment"
            Call o_trans_object.DeleteComment
        Case "GetComments"
            Call o_trans_object.GetComments
        Case "AddCustGrp"
            Call o_trans_object.Add
        Case "ChangeCustGrp"
            Call o_trans_object.Change
        Case "DeleteCustGrp"
            Call o_trans_object.Delete
        Case "GetCustGrp"
            Call o_trans_object.GetCustGrp
        Case "GetCustGrps"
            Call o_trans_object.getCustGrps
        Case "GetCustGrpCustomers"
            Call o_trans_object.getCustGrpCustomers
        Case "GetCustomers"
            Call o_trans_object.getCustomers
        Case "GetCustItem"
            Call o_trans_object.GetCustItem
        Case "AddCustItem"
            Call o_trans_object.Add
        Case "ChangeCustItem"
            Call o_trans_object.Change
        Case "DeleteCustItem"
            Call o_trans_object.Delete
        Case "GetCustItems"
            Call o_trans_object.getCustItems
        'HR
        Case "ChangePassword"
            Call o_trans_object.ChangePassword
        
        ' Security Authorization
        Case "GetSecurityAuth"
            Call o_trans_object.GetSecurityAuth
        Case "AddSecurityAuth"
            Call o_trans_object.Add
        Case "DeleteSecurityAuth"
            Call o_trans_object.Delete
        Case "ChangeSecurityAuth"
            Call o_trans_object.Change
        Case "GetSecurityAuths"
            Call o_trans_object.GetSecurityAuths
        
        ' Security Authorization
        Case "GetSecurityProf"
            Call o_trans_object.GetSecurityProf
        Case "AddSecurityProf"
            Call o_trans_object.Add
        Case "DeleteSecurityProf"
            Call o_trans_object.Delete
        Case "ChangeSecurityProf"
            Call o_trans_object.Change
        Case "GetSecurityProfs"
            Call o_trans_object.GetSecurityProfs
        
        '' Apar
        Case "Export"
            Call o_trans_object.ExportToQuickBooks
        Case "chart"
            Call o_trans_object.Chart
        ' System Function
        Case "GetSystemFuncs"
            Call o_trans_object.GetSystemFuncs
        Case Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "0100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 0
            o_error.p_nbr = Empty
            o_error.p_desc = "Transaction not found in transaction class"
            o_error.p_sqlstate_cd = Empty
    End Select
    
    ' Validate associate with security profile.
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
    Else
        If CheckSecurity(o_trans_data.p_Function_cd) = False Then
            With o_error
               .p_type_cd = "F"
               .p_err_cd = "9999"
               .p_nbr = Err.Number
               .p_desc = Err.Description
               .p_routine_nm = strRoutine_nm
               .p_message_id = 99999
             End With
             Err.Clear
        End If
    End If
    
    If o_trans_data.update_fl = True Then
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Call o_trans_data.RollBackTransDatabase(o_error)
        Else
            Call o_trans_data.CommitTransDatabase(o_error)
        End If
        o_trans_data.update_fl = False
    End If
    
    If m_blnLoggingFl = True Then
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Call logTransaction(strStartTs, o_trans_data.p_assoc_id, o_trans_data.p_Function_cd, False)
        Else
            Call logTransaction(strStartTs, o_trans_data.p_assoc_id, o_trans_data.p_Function_cd, True)
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

Public Function CheckSecurity(ByVal strFunctioncd As String) As Boolean
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    Dim varSecurity As Variant
    Dim intCount As Integer
    
    strRoutine_nm = "transaction.cls.CheckSecurity"

    CheckSecurity = False
    
    varSecurity = o_trans_data.p_Security
    
    For intCount = 0 To UBound(varSecurity, 2)
        If LCase(Trim(varSecurity(0, intCount))) = LCase(Trim(strFunctioncd)) Then
            CheckSecurity = True
            Exit For
        End If
    Next intCount
    
    intCount = 1
    
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

Private Function getStartTimestamp() As String
    On Error GoTo error_handler
    Set m_conLogging = New ADODB.Connection
    Dim rstData As ADODB.Recordset
    Call m_conLogging.Open(Me.o_trans_data.p_ConnectString)
    Set rstData = m_conLogging.Execute("SELECT CONVERT(char, GETDATE(), 121) AS TS")
    getStartTimestamp = rstData("TS")
    Exit Function
error_handler:
    Err.Clear
End Function

Private Sub logTransaction( _
          ByVal v_strStartTs As String _
        , ByVal v_lngAssocId As Long _
        , ByVal v_strFunctionNm As String _
        , ByVal v_blnSuccessFl As Boolean _
        )

    On Error GoTo error_handler
    Dim strSqlTx
    Dim bit As Byte
    
    bit = v_blnSuccessFl
    
    strSqlTx = "" _
        & "insert into trans_log (" _
        & "  assoc_id " _
        & ", start_ts " _
        & ", end_ts " _
        & ", func_nm " _
        & ", success_fl " _
        & ", duration_milliseconds " _
        & ") values (" _
        & "  " & v_lngAssocId & " " _
        & ", CONVERT(datetime, '" & v_strStartTs & "', 121) " _
        & ", getDate() " _
        & ", '" & v_strFunctionNm & "' " _
        & ",  " & bit & "  " _
        & ",  DATEDIFF(ms, CONVERT(datetime, '" & v_strStartTs & "', 121), CONVERT(datetime, getDate(), 121)) " _
        & ")"
        Call m_conLogging.Execute(strSqlTx)
        Call m_conLogging.Close
        Exit Sub
error_handler:
    Err.Clear
    If m_conLogging Is Nothing Then
    Else
        Call m_conLogging.Close
    End If

End Sub
