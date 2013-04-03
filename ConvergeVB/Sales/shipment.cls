VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "shipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_old_shipment As Object
Private mvarp_ship_id As Variant
Private mvarp_cargo_value_am As Variant
Private mvarp_carrier_nm As Variant
Private mvarp_cust_id As Variant
Private mvarp_cust_nbr As Variant
Private mvarp_cust_ship_acct_nbr As Variant
Private mvarp_drop_ship_fl As Variant
Private mvarp_insur_comp_nm As Variant
Private mvarp_insur_pol_nbr As Variant
Private mvarp_insured_am As Variant
Private mvarp_on_hold_reason_cd As Variant
Private mvarp_on_hold_reason_desc As Variant
Private mvarp_pkg_qty As Variant
Private mvarp_pack_slip_stnd_tx As Variant
Private mvarp_planned_del_dt As Variant
Private mvarp_planned_ship_dt As Variant
Private mvarp_sales_ord_id As Variant
Private mvarp_status_cd As Variant
Private mvarp_status_desc As Variant
Private mvarp_shipped_dt As Variant
Private mvarp_tracking_nbr As Variant
Private mvarp_transportation_cd As Variant
Private mvarp_transportation_desc As Variant
Private mvarp_type_desc As Variant
Private mvarp_type_cd As Variant
Private mvarp_ship_tot_lbs As Variant
Private mvarp_load_instruction_cd As Variant
Private mvarp_load_instruction_desc As Variant
Private mvaro_record_set As adodb.Recordset
Private mvarp_start_dt As Variant
Private mvarp_end_dt As Variant
Private mvarp_bypass_desc_fl As Variant
Private mvarp_freight_price As Variant
Private mvarp_blnAddAllItemsFl As Boolean

Public Property Let p_blnAddAllItemsFl(ByVal vData As Boolean)
    mvarp_blnAddAllItemsFl = vData
End Property
Public Property Get p_blnAddAllItemsFl() As Boolean
   p_blnAddAllItemsFl = mvarp_blnAddAllItemsFl
End Property

Public Property Let p_bypass_desc_fl(ByVal vData As Variant)
    mvarp_bypass_desc_fl = vData
End Property
Public Property Get p_bypass_desc_fl() As Variant
        p_bypass_desc_fl = mvarp_bypass_desc_fl
End Property

Public Property Let p_start_dt(ByVal vData As Variant)
    mvarp_start_dt = vData
End Property
Public Property Get p_start_dt() As Variant
        p_start_dt = mvarp_start_dt
End Property
Public Property Let p_end_dt(ByVal vData As Variant)
    mvarp_end_dt = vData
End Property
Public Property Get p_end_dt() As Variant
        p_end_dt = mvarp_end_dt
End Property
Public Property Let p_freight_price(ByVal vData As Variant)
    mvarp_freight_price = vData
End Property
Public Property Get p_freight_price() As Variant
        p_freight_price = mvarp_freight_price
End Property

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
        Set o_record_set = mvaro_record_set
End Property

Public Property Set o_old_shipment(ByVal vData As Object)
    Set mvaro_old_shipment = vData
End Property
Public Property Get o_old_shipment() As Object
        Set o_old_shipment = mvaro_old_shipment
End Property

Public Property Let p_drop_ship_fl(ByVal vData As Variant)
    mvarp_drop_ship_fl = vData
End Property
Public Property Get p_drop_ship_fl() As Variant
        p_drop_ship_fl = mvarp_drop_ship_fl
End Property
Public Property Let p_insured_am(ByVal vData As Variant)
    mvarp_insured_am = vData
End Property
Public Property Get p_insured_am() As Variant
        p_insured_am = mvarp_insured_am
End Property
Public Property Let p_ship_tot_lbs(ByVal vData As Variant)
    mvarp_ship_tot_lbs = vData
End Property
Public Property Get p_ship_tot_lbs() As Variant
        p_ship_tot_lbs = mvarp_ship_tot_lbs
End Property
Public Property Let p_load_instruction_cd(ByVal vData As Variant)
    mvarp_load_instruction_cd = vData
End Property
Public Property Get p_load_instruction_cd() As Variant
        p_load_instruction_cd = mvarp_load_instruction_cd
End Property
Public Property Let p_load_instruction_desc(ByVal vData As Variant)
    mvarp_load_instruction_desc = vData
End Property
Public Property Get p_load_instruction_desc() As Variant
        p_load_instruction_desc = mvarp_load_instruction_desc
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    mvarp_type_cd = vData
End Property
Public Property Get p_type_cd() As Variant
        p_type_cd = mvarp_type_cd
End Property
Public Property Let p_type_desc(ByVal vData As Variant)
    mvarp_type_desc = vData
End Property
Public Property Get p_type_desc() As Variant
        p_type_desc = mvarp_type_desc
End Property
Public Property Let p_transportation_desc(ByVal vData As Variant)
    mvarp_transportation_desc = vData
End Property
Public Property Get p_transportation_desc() As Variant
        p_transportation_desc = mvarp_transportation_desc
End Property
Public Property Let p_transportation_cd(ByVal vData As Variant)
    mvarp_transportation_cd = vData
End Property
Public Property Get p_transportation_cd() As Variant
        p_transportation_cd = mvarp_transportation_cd
End Property
Public Property Let p_tracking_nbr(ByVal vData As Variant)
    mvarp_tracking_nbr = vData
End Property
Public Property Get p_tracking_nbr() As Variant
        p_tracking_nbr = mvarp_tracking_nbr
End Property
Public Property Let p_shipped_dt(ByVal vData As Variant)
    mvarp_shipped_dt = vData
End Property
Public Property Get p_shipped_dt() As Variant
        p_shipped_dt = mvarp_shipped_dt
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_sales_ord_id(ByVal vData As Variant)
    mvarp_sales_ord_id = vData
End Property
Public Property Get p_sales_ord_id() As Variant
        p_sales_ord_id = mvarp_sales_ord_id
End Property
Public Property Let p_pack_slip_stnd_tx(ByVal vData As Variant)
    mvarp_pack_slip_stnd_tx = vData
End Property
Public Property Get p_pack_slip_stnd_tx() As Variant
        p_pack_slip_stnd_tx = mvarp_pack_slip_stnd_tx
End Property
Public Property Let p_pkg_qty(ByVal vData As Variant)
    mvarp_pkg_qty = vData
End Property
Public Property Get p_pkg_qty() As Variant
        p_pkg_qty = mvarp_pkg_qty
End Property
Public Property Let p_on_hold_reason_desc(ByVal vData As Variant)
    mvarp_on_hold_reason_desc = vData
End Property
Public Property Get p_on_hold_reason_desc() As Variant
        p_on_hold_reason_desc = mvarp_on_hold_reason_desc
End Property
Public Property Let p_on_hold_reason_cd(ByVal vData As Variant)
    mvarp_on_hold_reason_cd = vData
End Property
Public Property Get p_on_hold_reason_cd() As Variant
        p_on_hold_reason_cd = mvarp_on_hold_reason_cd
End Property
Public Property Let p_insur_pol_nbr(ByVal vData As Variant)
    mvarp_insur_pol_nbr = vData
End Property
Public Property Get p_insur_pol_nbr() As Variant
        p_insur_pol_nbr = mvarp_insur_pol_nbr
End Property
Public Property Let p_insur_comp_nm(ByVal vData As Variant)
    mvarp_insur_comp_nm = vData
End Property
Public Property Get p_insur_comp_nm() As Variant
        p_insur_comp_nm = mvarp_insur_comp_nm
End Property
Public Property Let p_planned_del_dt(ByVal vData As Variant)
    mvarp_planned_del_dt = vData
End Property
Public Property Get p_planned_del_dt() As Variant
        p_planned_del_dt = mvarp_planned_del_dt
End Property
Public Property Let p_planned_ship_dt(ByVal vData As Variant)
    mvarp_planned_ship_dt = vData
End Property
Public Property Get p_planned_ship_dt() As Variant
        p_planned_ship_dt = mvarp_planned_ship_dt
End Property
Public Property Let p_cust_ship_acct_nbr(ByVal vData As Variant)
    mvarp_cust_ship_acct_nbr = vData
End Property
Public Property Get p_cust_ship_acct_nbr() As Variant
        p_cust_ship_acct_nbr = mvarp_cust_ship_acct_nbr
End Property
Public Property Let p_cust_nbr(ByVal vData As Variant)
    mvarp_cust_nbr = vData
End Property
Public Property Get p_cust_nbr() As Variant
        p_cust_nbr = mvarp_cust_nbr
End Property
Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
        p_cust_id = mvarp_cust_id
End Property
Public Property Let p_cargo_value_am(ByVal vData As Variant)
    mvarp_cargo_value_am = vData
End Property
Public Property Get p_cargo_value_am() As Variant
        p_cargo_value_am = mvarp_cargo_value_am
End Property
Public Property Let p_carrier_nm(ByVal vData As Variant)
    mvarp_carrier_nm = vData
End Property
Public Property Get p_carrier_nm() As Variant
        p_carrier_nm = mvarp_carrier_nm
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
Public Property Let p_ship_id(ByVal vData As Variant)
    mvarp_ship_id = vData
End Property
Public Property Get p_ship_id() As Variant
        p_ship_id = mvarp_ship_id
End Property

Public Sub GetShipment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Shipment.cls.GetShipment"
    
    Call SelectShipment
    
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
    strRoutine_nm = "Shipment.cls.Add"
    
    Call validate("add")
    Call IssueInsertData
    Call AddShipmentItems
    
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

Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Shipment.cls.Change"
    
    If o_old_shipment Is Nothing Then
        Set o_old_shipment = CreateObject("converge_sales.shipment")
        Set o_old_shipment.o_error = o_error
        Set o_old_shipment.o_trans_data = o_trans_data
        o_old_shipment.p_ship_id = p_ship_id
        o_old_shipment.p_bypass_desc_fl = "Y"
        o_old_shipment.GetShipment
    End If
    
    Call validate("change")
    Call IssueUpdateData
    
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
    strRoutine_nm = "Shipment.cls.Delete"
    
    If o_old_shipment Is Nothing Then
        Set o_old_shipment = CreateObject("converge_sales.shipment")
        Set o_old_shipment.o_error = o_error
        Set o_old_shipment.o_trans_data = o_trans_data
        o_old_shipment.p_ship_id = p_ship_id
        o_old_shipment.p_bypass_desc_fl = "Y"
        o_old_shipment.GetShipment
    End If
    
    Call validate("delete")
    Call IssueDeleteData
    Call IssueDeleteShipmentItems
    Call IssueDeleteName
    Call IssueDeleteAddr
    
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

Public Sub changeStatus()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdatedQty As Double
    Dim strSql As String
    Dim objSalesInvoice As sales_invoice
    
    strRoutine_nm = "shipment.cls.changeStatus"

    If p_status_cd = "CA" Or _
       p_status_cd = "OP" Or _
       p_status_cd = "CL" Or _
       p_status_cd = "IP" Or _
       p_status_cd = "IN" Or _
       p_status_cd = "OH" Then
    Else
        With o_error
            .p_type_cd = "E"
            .p_err_cd = "1000"
            .p_routine_nm = strRoutine_nm
            .p_message_id = 1017
            Exit Sub
        End With
    End If
    
    Set o_old_shipment = CreateObject("converge_sales.shipment")
    Set o_old_shipment.o_error = o_error
    Set o_old_shipment.o_trans_data = o_trans_data
    o_old_shipment.p_ship_id = p_ship_id
    o_old_shipment.p_bypass_desc_fl = "Y"
    o_old_shipment.GetShipment
    
    If o_old_shipment.p_status_cd = "PL" Then
        If p_status_cd = "OP" Or _
           p_status_cd = "CA" Or _
           p_status_cd = "OH" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "1200"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1017
                Exit Sub
            End With
        End If
    End If
    
    If o_old_shipment.p_status_cd = "OP" Then
        If p_status_cd = "OH" Or _
           p_status_cd = "OP" Or _
           p_status_cd = "CL" Or _
           p_status_cd = "CA" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "1400"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1017
                Exit Sub
            End With
        End If
    End If
    
    If o_old_shipment.p_status_cd = "IP" Then
        If p_status_cd = "CL" Or _
           p_status_cd = "OP" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "1600"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1017
                Exit Sub
            End With
        End If
    End If
    
    If o_old_shipment.p_status_cd = "CL" Then
        If p_status_cd = "IN" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "1800"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1017
                Exit Sub
            End With
        End If
    End If

    
    If o_old_shipment.p_status_cd = "CA" Then
        With o_error
            .p_type_cd = "E"
            .p_err_cd = "2000"
            .p_routine_nm = strRoutine_nm
            .p_message_id = 1017
            Exit Sub
        End With
    End If
    
    If o_old_shipment.p_status_cd = "OH" Then
        If p_status_cd = "OP" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "2200"
                .p_routine_nm = strRoutine_nm
                .p_message_id = 1017
                Exit Sub
            End With
        End If
    End If
    
    If o_old_shipment.p_status_cd = p_status_cd And p_status_cd <> "OP" Then
        Exit Sub
    End If
    
    If p_status_cd = "CL" Then
                    'If this is a first leg shipment, take the status directly to 'IN'.  An invoice
                    'does not need to be created for 1st leg shipments
        If Me.o_old_shipment.p_type_cd = "1L" Then
            p_status_cd = "IN"
        End If
        
        strSql = "update shipment set "
        strSql = strSql & fUpdateVariable("status_cd", p_status_cd, "S", "=")
        strSql = strSql & "," & fUpdateVariable("shipped_dt", "current_ts", "D", "=")
        strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
        strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
        strSql = strSql & " where"
        strSql = strSql & " " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    Else
        strSql = "update shipment set "
        strSql = strSql & fUpdateVariable("status_cd", p_status_cd, "S", "=")
        strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
        strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
        strSql = strSql & " where"
        strSql = strSql & " " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    
    End If
     
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedQty, o_error)
    
    If dblUpdatedQty = 0 Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "2400"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 1235
        Exit Sub
    End If
    
                    'if a drop shipment is being closed, check to see if
                    'is a shipment for goodmn cust_id = 1160
    If p_status_cd = "CL" And Me.o_old_shipment.p_drop_ship_fl = "Y" Then
        Dim objSalesOrder As sales_order
        Dim objRecordset As adodb.Recordset
        Set objSalesOrder = New sales_order
        With objSalesOrder
            .p_sales_ord_id = o_old_shipment.p_sales_ord_id
            Set .o_error = Me.o_error
            Set .o_trans_data = Me.o_trans_data
            .getSalesOrder
        End With
                    'custId 1160 = goodmn
                    'custId 2325 = goodmanFur
                    'custId 2271 = goodmanSP
                    'custId 1020 = aman/t which is now goodmanTN
        If objSalesOrder.p_cust_id = 1160 Or _
           objSalesOrder.p_cust_id = 2325 Or _
           objSalesOrder.p_cust_id = 1020 Or _
           objSalesOrder.p_cust_id = 2271 Then
            Call specialPhoenixProcess(Me.p_ship_id)
        End If
    End If
                    
                    ' do not invoice 1st Leg Shipments
    If p_status_cd = "IN" And Me.o_old_shipment.p_type_cd <> "1L" Then
        Set objSalesInvoice = New sales_invoice
        Set objSalesInvoice.o_error = Me.o_error
        Set objSalesInvoice.o_trans_data = Me.o_trans_data
                    
                    ' get the existing invoice for the shipment, if it is
                    ' in a sent status, cancel it.  If it has not been sent
                    ' delete it
        objSalesInvoice.p_ship_id = p_ship_id
        Call objSalesInvoice.GetSalesInvoice
        If IsNull(objSalesInvoice.p_sales_inv_id) = False Then
            If objSalesInvoice.p_status_cd = "SN" Then
                objSalesInvoice.p_status_cd = "CA"
                Call objSalesInvoice.Change
                With o_error
                  .p_type_cd = "I"
                  .p_err_cd = "9900"
                  .p_routine_nm = strRoutine_nm
                  .p_message_id = 0
                  .p_desc = objSalesInvoice.p_sales_inv_id
                End With
            Else
                Call objSalesInvoice.Delete
            End If
        End If
                    ' generate the new invoice
        objSalesInvoice.p_ship_id = p_ship_id
        objSalesInvoice.p_sales_ord_id = p_sales_ord_id
        objSalesInvoice.Generate
        objSalesInvoice.ExportToQB
        Set objSalesInvoice = Nothing
    End If
    
    If p_status_cd = "OP" Then
        Call Me.setStatus( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.p_ship_id _
                    )
    End If
    
    Exit Sub
    
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
End Sub


Public Sub setStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblShipId As Double _
            )

    On Error GoTo error_handler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment.cls.setStatus"
    
    Dim varLoadedQty As Variant _
        , strStatusCd As String _
        , strSql As String _
        , dblUpdatedQty As Double
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.p_ship_id = v_dblShipId
    Call Me.GetShipment
    
    If Me.p_status_cd = "OP" Or Me.p_status_cd = "IP" Then
    Else
        Exit Sub
    End If
    
    varLoadedQty = getLoadedQty(r_objError, r_objTransData, v_dblShipId)
    
    If varLoadedQty > 0 Then
        strStatusCd = "IP"
    Else
        strStatusCd = "OP"
    End If
    
    If Me.p_status_cd <> strStatusCd Then
        strSql = "update shipment set "
        strSql = strSql & fUpdateVariable("status_cd", strStatusCd, "S", "=")
        strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
        strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
        strSql = strSql & " where"
        strSql = strSql & " " & fWhereVariable("ship_id", v_dblShipId, "N", "=")
        
        Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedQty, r_objError)
    End If
    
    Exit Sub
error_handler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "2600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Function getLoadedQty( _
                  ByRef r_objError As Object _
                , ByRef r_objTransData As Object _
                , ByVal v_dblShipId As Double _
                )
    
    On Error GoTo error_handler
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "shipment.cls.getLoadedQty"
    
    Dim strSql As Variant
    Dim objRecordset As adodb.Recordset
    
    strSql = "SELECT sum(loaded_qty) "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE "
    strSql = strSql & fWhereVariable("ship_id", v_dblShipId, "N", "=")
    
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        getLoadedQty = 0
    Else
        getLoadedQty = objRecordset(0)
    End If
    
    Exit Function
error_handler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "2800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear

End Function
Private Sub validate(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    Dim objSalesOrder As Object
    
    strRoutine_nm = "Shipment.cls.Validate"
    Set objSalesOrder = CreateObject("converge_sales.sales_order")
    
    
    If sFunction_cd = "add" Then
        If IsDate(Me.p_planned_del_dt) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 3030
            Exit Sub
        End If
        If IsDate(Me.p_planned_ship_dt) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 3029
            Exit Sub
        End If
        
        If IsNumeric(Me.p_pkg_qty) = False Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3001"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 3431
            o_error.p_desc = "Package quantity is required"
            Exit Sub
        End If
        
        'p_sales_ord_id can be empty on a change but it can never be null
        If IsNull(p_sales_ord_id) = True Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3000"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1018
            Exit Sub
        End If
        
        Set objSalesOrder.o_error = o_error
        Set objSalesOrder.o_trans_data = o_trans_data
        
        If Len(p_sales_ord_id) > 0 Then
            objSalesOrder.p_sales_ord_id = p_sales_ord_id
        Else
            objSalesOrder.p_sales_ord_id = o_old_shipment.p_sales_ord_id
        End If
        
        objSalesOrder.p_bypass_desc_fl = "Y"
        objSalesOrder.p_bypass_nm_addr_fl = "Y"
        Call objSalesOrder.getSalesOrder
        
        If IsNull(objSalesOrder.p_sales_ord_id) = True Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3200"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1018
            Exit Sub
        End If
        
        If p_drop_ship_fl = "Y" Or p_drop_ship_fl = "N" Then
        Else
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "3400"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1275
            Exit Sub
        End If
        
        'second leg consignment can't be a drop shipment
        If p_type_cd = "2L" And p_drop_ship_fl = "Y" Then
             o_error.p_type_cd = "F"
             o_error.p_err_cd = "3600"
             o_error.p_routine_nm = strRoutine_nm
             o_error.p_message_id = 1250
             Exit Sub
        End If
        
'==>if order is a consignment order, shipment must be either 1st or 2nd leg
'   if not, shipment must be regular
        If objSalesOrder.p_consignment_fl = "Y" Then
            If p_type_cd = "1L" Or p_type_cd = "2L" Then
            Else
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "3800"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1255
                Exit Sub
            End If
        Else
            If p_type_cd <> "RG" Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "4000"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1260
                Exit Sub
            End If
        End If
        
'==> if sales order is a drop shipment, shipment must be a drop
'    shipment unless it is 2nd leg consignment
'==> if sales order is not a drop shipment, shipment can't be a
'    drop shipment
        If objSalesOrder.p_drop_ship_fl = "Y" Then
            If p_drop_ship_fl = "N" Then
                If p_type_cd = "2L" Then
                Else
                    o_error.p_type_cd = "F"
                    o_error.p_err_cd = "4200"
                    o_error.p_routine_nm = strRoutine_nm
                    o_error.p_message_id = 1250
                    Exit Sub
                End If
            End If
        Else
            If p_drop_ship_fl = "Y" Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "4400"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1270
                Exit Sub
            End If
        End If
    
        If objSalesOrder.p_status_cd = "PL" Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "4600"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1280
            Exit Sub
        End If
    
        If fValuePresent(p_load_instruction_cd) = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "4750"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1021
            Exit Sub
        End If
    End If
    
    If sFunction_cd = "change" Then
        If o_old_shipment.p_status_cd = "PL" Or _
           o_old_shipment.p_status_cd = "OP" Or _
           o_old_shipment.p_status_cd = "OH" Or _
           o_old_shipment.p_status_cd = "IP" Or _
           o_old_shipment.p_status_cd = "CL" Then
        Else
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "4800"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1019
            Exit Sub
        End If
    
        If IsEmpty(Me.p_planned_del_dt) = False Then
            If IsDate(Me.p_planned_del_dt) = False Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "3000"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 3030
                Exit Sub
            End If
        End If
        If IsEmpty(Me.p_planned_ship_dt) = False Then
            If IsDate(Me.p_planned_ship_dt) = False Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "3000"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 3029
                Exit Sub
            End If
       End If
        If IsEmpty(Me.p_pkg_qty) = False Then
            If IsNumeric(Me.p_pkg_qty) = False Then
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "3031"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 3431
                o_error.p_desc = "Package quantity is required"
                Exit Sub
            End If
       End If
    End If
        
    If sFunction_cd = "delete" Then
        If o_old_shipment.p_status_cd = "PL" Then
        Else
             o_error.p_type_cd = "E"
             o_error.p_err_cd = "5000"
             o_error.p_routine_nm = strRoutine_nm
             o_error.p_message_id = 1285
             Exit Sub
        End If
    End If
   
    If sFunction_cd = "change" Or sFunction_cd = "add" Then
        If IsNull(p_load_instruction_cd) = True Then
            o_error.p_type_cd = "F"
            o_error.p_err_cd = "4850"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 1021
            Exit Sub
        End If
        If fValuePresent(p_load_instruction_cd) = True Then
            If Len(fGetDesc("load_instruction_cd", p_load_instruction_cd, o_error, o_trans_data)) > 0 Then
            Else
                o_error.p_type_cd = "F"
                o_error.p_err_cd = "4860"
                o_error.p_routine_nm = strRoutine_nm
                o_error.p_message_id = 1022
                Exit Sub
            End If
        End If
        
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5200"
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
    strRoutine_nm = "Shipment.cls.IssueDeleteData"

    strSql = "delete from Shipment "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteName()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "Shipment.cls.IssueDeleteName"
    
    strSql = "delete from name "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_ship_id, "N", "=")
    strSql = strSql & "and" & fWhereVariable("obj_nm", "Shipment", "S", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteAddr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "Shipment.cls.IssueDeleteAddr"
    
    strSql = "delete from address "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_id", p_ship_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "5800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueInsertData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "shipment.cls.IssueInsertData"
    
    p_ship_id = fGetId("ship_id", o_error, o_trans_data)
    If p_drop_ship_fl = "Y" Then
        p_status_cd = "OP"
    Else
        p_status_cd = "PL"
    End If
    
    strSql = "insert into shipment ( "
    strSql = strSql & "  ship_id"
    strSql = strSql & ", cargo_value_am"
    strSql = strSql & ", carrier_nm"
    strSql = strSql & ", cust_id"
    strSql = strSql & ", cust_ship_acct_nbr"
    strSql = strSql & ", drop_ship_fl"
    strSql = strSql & ", planned_del_dt"
    strSql = strSql & ", planned_ship_dt"
    strSql = strSql & ", insur_comp_nm"
    strSql = strSql & ", insur_pol_nbr"
    strSql = strSql & ", insured_am"
    strSql = strSql & ", on_hold_reason_cd"
    strSql = strSql & ", pkg_qty"
    strSql = strSql & ", pack_slip_stnd_tx"
    strSql = strSql & ", sales_ord_id"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", shipped_dt"
    strSql = strSql & ", tracking_nbr"
    strSql = strSql & ", transportation_cd"
    strSql = strSql & ", type_cd"
    strSql = strSql & ", ship_tot_lbs"
    strSql = strSql & ", load_instruction_cd"
    strSql = strSql & ", freight_price"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_ship_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cargo_value_am, "N")
    strSql = strSql & "," & fInsertVariable(p_carrier_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_cust_id, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_ship_acct_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_drop_ship_fl, "S")
    strSql = strSql & "," & fInsertVariable(p_planned_del_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_planned_ship_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_insur_comp_nm, "S")
    strSql = strSql & "," & fInsertVariable(p_insur_pol_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_insured_am, "S")
    strSql = strSql & "," & fInsertVariable(p_on_hold_reason_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_pkg_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_pack_slip_stnd_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_sales_ord_id, "N")
    strSql = strSql & "," & fInsertVariable(p_status_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_shipped_dt, "D")
    strSql = strSql & "," & fInsertVariable(p_tracking_nbr, "N")
    strSql = strSql & "," & fInsertVariable(p_transportation_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_ship_tot_lbs, "N")
    strSql = strSql & "," & fInsertVariable(p_load_instruction_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_freight_price, "N")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueUpdateData()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "shipment.cls.IssueUpdateData"
    
    strSql = "update shipment set "
    strSql = strSql & " " & fUpdateVariable("cargo_value_am", p_cargo_value_am, "N", "=")
    strSql = strSql & "," & fUpdateVariable("carrier_nm", p_carrier_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("cust_ship_acct_nbr", p_cust_ship_acct_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("planned_del_dt", p_planned_del_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("planned_ship_dt", p_planned_ship_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("insur_comp_nm", p_insur_comp_nm, "S", "=")
    strSql = strSql & "," & fUpdateVariable("insur_pol_nbr", p_insur_pol_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("insured_am", p_insured_am, "N", "=")
    strSql = strSql & "," & fUpdateVariable("on_hold_reason_cd", p_on_hold_reason_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("pkg_qty", p_pkg_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("pack_slip_stnd_tx", p_pack_slip_stnd_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("shipped_dt", p_shipped_dt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("tracking_nbr", p_tracking_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("transportation_cd", p_transportation_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ship_tot_lbs", p_ship_tot_lbs, "N", "=")
    strSql = strSql & "," & fUpdateVariable("load_instruction_cd", p_load_instruction_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("freight_price", p_freight_price, "N", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("ship_id", p_ship_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6200"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub SelectShipment()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "Shipment.cls.SelectShipment"

    If IsNull(p_ship_id) = True Then
       p_ship_id = Null
       Exit Sub
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM shipment "
        strSql = strSql & "WHERE "
        strSql = strSql & fWhereVariable("ship_id", p_ship_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_ship_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
    Else
        p_ship_id = objRecordset("ship_id")
        p_cargo_value_am = objRecordset("cargo_value_am")
        p_carrier_nm = objRecordset("carrier_nm")
        p_cust_id = objRecordset("cust_id")
        p_cust_ship_acct_nbr = objRecordset("cust_ship_acct_nbr")
        p_drop_ship_fl = objRecordset("drop_ship_fl")
        p_planned_del_dt = objRecordset("planned_del_dt")
        p_planned_ship_dt = objRecordset("planned_ship_dt")
        p_insur_comp_nm = objRecordset("insur_comp_nm")
        p_insur_pol_nbr = objRecordset("insur_pol_nbr")
        p_insured_am = objRecordset("insured_am")
        p_on_hold_reason_cd = objRecordset("on_hold_reason_cd")
        p_pkg_qty = objRecordset("pkg_qty")
        p_pack_slip_stnd_tx = objRecordset("pack_slip_stnd_tx")
        p_sales_ord_id = objRecordset("sales_ord_id")
        p_status_cd = objRecordset("status_cd")
        p_shipped_dt = objRecordset("shipped_dt")
        p_tracking_nbr = objRecordset("tracking_nbr")
        p_transportation_cd = objRecordset("transportation_cd")
        p_type_cd = objRecordset("type_cd")
        p_ship_tot_lbs = objRecordset("ship_tot_lbs")
        p_load_instruction_cd = objRecordset("load_instruction_cd")
        p_freight_price = objRecordset("freight_price")
        
        objRecordset.Close
        Set objRecordset = Nothing
        If p_bypass_desc_fl = "Y" Then
        Else
            p_load_instruction_desc = fGetDesc("load_instruction_cd", p_load_instruction_cd, o_error, o_trans_data)
            p_transportation_desc = fGetDesc("transportation_cd", p_transportation_cd, o_error, o_trans_data)
            p_cust_nbr = fGetCustNbr(p_cust_id, o_error, o_trans_data)
            p_on_hold_reason_desc = fGetDesc("on_hold_reason_cd", p_on_hold_reason_cd, o_error, o_trans_data)
            p_status_desc = fGetDesc("shipment_status_cd", p_status_cd, o_error, o_trans_data)
            p_type_desc = fGetDesc("shipment_type_cd", p_type_cd, o_error, o_trans_data)
        End If
    End If
     
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6400"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub AddShipmentItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    Dim objShipmentItem As Object
    Dim blnItemAddedFl As Boolean
    
    strRoutine_nm = "Shipment.cls.AddShipmentItems"

    strSql = "SELECT a.status_cd, b.mat_req_id, a.line_nbr, a.item_id, b.ready_qty, a.ordered_qty, a.planned_ship_qty, a.shipped_qty " & _
             "FROM sales_order_item a " & _
             "left join mat_req b " & _
             " on a.sales_ord_id = b.sales_ord_id and a.line_nbr = b.so_line_nbr " & _
             "WHERE " & _
             "   " & fWhereVariable("a.sales_ord_id", p_sales_ord_id, "N", "=")
    Set objShipmentItem = CreateObject("converge_sales.shipment_item")

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
            
    blnItemAddedFl = False

    Do Until objRecordset.EOF = True Or o_error.type_cd = "E" Or o_error.type_cd = "F"
        Debug.Print (objRecordset("line_nbr"))
        If Me.p_blnAddAllItemsFl = False And ( _
           objRecordset("status_cd") = "CA" Or objRecordset("status_cd") = "CP" Or _
           objRecordset("status_cd") = "OH" Or _
           objRecordset("status_cd") = "IC" Or _
           objRecordset("status_cd") = "CL") Then
        Else
            Set objShipmentItem.o_trans_data = o_trans_data
            Set objShipmentItem.o_error = o_error
            objShipmentItem.ship_id = p_ship_id
            objShipmentItem.line_nbr = objRecordset("line_nbr")
            objShipmentItem.ship_type_cd = p_type_cd
            objShipmentItem.sales_ord_id = p_sales_ord_id
            objShipmentItem.item_id = objRecordset("item_id")
            objShipmentItem.drop_ship_fl = p_drop_ship_fl
            If p_drop_ship_fl = "N" And IsNull(objRecordset("mat_req_id")) = False Then
                Select Case p_load_instruction_cd
                    Case "LA"
                        objShipmentItem.planned_qty = objRecordset("ready_qty") - objRecordset("planned_ship_qty") + objRecordset("shipped_qty")
                    Case "LP"
                        objShipmentItem.planned_qty = objRecordset("ready_qty") - objRecordset("planned_ship_qty") + objRecordset("shipped_qty")
                    Case "LC"
                        objShipmentItem.planned_qty = objRecordset("ordered_qty") - objRecordset("planned_ship_qty") + objRecordset("shipped_qty")
                    Case Else
                        objShipmentItem.planned_qty = objRecordset("ready_qty") - objRecordset("planned_ship_qty") + objRecordset("shipped_qty")
                End Select
                If objShipmentItem.planned_qty < 0 Then
                    objShipmentItem.planned_qty = 0
                End If
            Else
                objShipmentItem.planned_qty = 0
            End If
            If objRecordset("status_cd") = "CL" Or _
               objRecordset("status_cd") = "CP" Or _
               objRecordset("status_cd") = "OH" Or _
               objRecordset("status_cd") = "CA" Then
               objShipmentItem.planned_qty = 0
            End If
            objShipmentItem.loaded_qty = 0
            Call objShipmentItem.Add
            blnItemAddedFl = True
        End If
       objRecordset.MoveNext
    Loop
    objRecordset.Close
    Set objRecordset = Nothing
    
    If blnItemAddedFl = False Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "4800"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 3427
        Exit Sub
    End If
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub

Private Sub IssueDeleteShipmentItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm
    Dim dblUpdateQty As Double
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    Dim objShipmentItem As Object
    
    strRoutine_nm = "Shipment.cls.DeleteShipmentItems"

    strSql = "SELECT line_nbr "
    strSql = strSql & "FROM shipment_item "
    strSql = strSql & "WHERE "
    strSql = strSql & fWhereVariable("ship_id", p_ship_id, "N", "=")
    Set objShipmentItem = CreateObject("converge_sales.shipment_item")

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    Do Until objRecordset.EOF = True Or o_error.type_cd = "E" Or o_error.type_cd = "F"
        Set objShipmentItem.o_trans_data = o_trans_data
        Set objShipmentItem.o_error = o_error
        objShipmentItem.ship_id = p_ship_id
        objShipmentItem.line_nbr = objRecordset("line_nbr")
        objShipmentItem.p_sales_ord_id = p_sales_ord_id
        Call objShipmentItem.Delete
        objRecordset.MoveNext
    Loop
    objRecordset.Close
    Set objRecordset = Nothing

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "6800"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear

End Sub

Public Sub getShipments()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.GetShipments"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    strSql = "SELECT  " & _
                " a.ship_id, " & _
                " a.status_cd, b.cust_nbr," & _
                " a.sales_ord_id, " & _
                " a.planned_ship_dt, " & _
                " a.drop_ship_fl, " & _
                " a.type_cd, a.planned_del_dt, " & _
                " a.freight_price " & _
                " FROM shipment a, customer b " & _
                " WHERE a.cust_id = b.cust_id "

    If fValuePresent(p_cust_nbr) = True Then
       strSql = strSql & "  AND " & fWhereVariable("b.cust_nbr", p_cust_nbr, "S", "=")
    End If
    
    If fValuePresent(p_status_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.status_cd", p_status_cd, "S", "=")
    End If
    
    If fValuePresent(p_type_cd) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.type_cd", p_type_cd, "S", "=")
    End If
    
    If fValuePresent(p_sales_ord_id) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.sales_ord_id", p_sales_ord_id, "S", "=")
    End If
    
    If fValuePresent(p_start_dt) = True Then
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & " AND(" & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=") & ")"
        Else
            strSql = strSql & "  AND " & fWhereVariable("a.planned_ship_dt", p_start_dt, "D", "=")
        End If
    Else
        If fValuePresent(p_end_dt) = True Then
            strSql = strSql & " AND " & fWhereVariable("a.planned_ship_dt", p_end_dt, "D", "<=") & ")"
        End If
    End If
               
    If fValuePresent(p_drop_ship_fl) = True Then
       strSql = strSql & "  AND " & fWhereVariable("a.drop_ship_fl", p_drop_ship_fl, "S", "=")
    End If
               
    strSql = strSql & "  ORDER BY a.planned_ship_dt "
                
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
      .p_err_cd = "7000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub


Private Sub specialPhoenixProcess(ByVal v_lngShipId As Long)
    
                    'this routine will relieve inventory from the pheonix location
                    'when a drop shipment is made from pheonix to goodmn.  The item
                    'being releived is a component item that timco sent to pheonix
                    'to be used in an assembly process
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Orders.cls.specialPhoenixProcess"
    Dim objRecordset1 As adodb.Recordset
    Dim strSql As String
    
            'loop through all the items on the shipment looking for items that
            'require special processing
    strSql = "select * from shipment_item where ship_id = " & v_lngShipId
    Call o_trans_data.OpenRecordset(objRecordset1, strSql, o_error)
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Do While objRecordset1.EOF = False And Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
        Dim lngItemId As Long
        Dim lngAdjQty As Long
        Dim lngTotAdjQty As Long
        lngItemId = objRecordset1("item_id")
        lngAdjQty = objRecordset1("loaded_qty")
        If lngItemId = 25792 Or _
            lngItemId = 25793 Or _
            lngItemId = 25794 Or _
            lngItemId = 25795 Or _
            lngItemId = 25796 Or _
            lngItemId = 25797 Or _
            lngItemId = 25798 Or _
            lngItemId = 25830 Or _
            lngItemId = 25799 Then
            
            Select Case lngItemId
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
                Case 25830
                    lngItemId = 25787
            End Select
            
            
                    'Validate that there is enough inventory for the item being
                    'consumed from the phoenix location.
            Dim objRecordset As adodb.Recordset
            strSql = "SELECT SUM(adj_qty) as inv_qty " _
                & "From inventory_adj " _
                & "WHERE inv_loc_id = 'Phoenix' AND item_id = " & lngItemId & " "
            
            Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
            If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
                Exit Sub
            End If
            
            If objRecordset.EOF = True Then
                Me.o_error.p_type_cd = "E"
                Me.o_error.p_err_cd = "4500"
                Me.o_error.err_desc = "Insufficient inventory for line number: " & objRecordset1("line_nbr")
                Me.o_error.p_routine_nm = strRoutine_nm
                Me.o_error.p_message_id = 3430
                Exit Sub
            Else
                If objRecordset("inv_qty") < lngAdjQty Then
                    Me.o_error.p_type_cd = "E"
                    Me.o_error.p_err_cd = "4501"
                    Me.o_error.err_desc = "Insufficient inventory for line number: " & objRecordset1("line_nbr")
                    Me.o_error.p_routine_nm = strRoutine_nm
                    Me.o_error.p_message_id = 3430
                    Exit Sub
                End If
            End If
                        
                        
            Dim objRecordset2 As adodb.Recordset
                    'get material requirement for the item
            strSql = "select * from mat_req " _
                & "where item_id = " & lngItemId & " " _
                & "and status_cd in ('OP', 'BO', 'RD') " _
                & "order by required_dt "
                     
            Call o_trans_data.OpenRecordset(objRecordset2, strSql, o_error)
            If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
                Exit Sub
            End If
                    'cycle through all the material requirements until the
                    'we relieve the quantity loaded on the shipment or
                    'run out of material requirements
            Do While lngTotAdjQty < lngAdjQty And _
               objRecordset2.EOF = False And _
               Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
                
                Dim lngMatReqReadyQty As Long
                Dim lntTotMRAdjQty As Long
                Dim objRecordset3 As adodb.Recordset
                Dim lngTotMRAdjQty As Long
                lngMatReqReadyQty = objRecordset2("ready_qty")
                lntTotMRAdjQty = 0
                
                        'get inventory for item
                strSql = "SELECT SUM(adj_qty) as inv_qty, lot_id, bin_id " _
                    & "From inventory_adj " _
                    & "WHERE inv_loc_id = 'Phoenix' AND item_id = " & lngItemId & " " _
                    & "GROUP BY lot_id, bin_id, item_id " _
                    & "Having Sum(adj_qty) > 0 "
                
                Call o_trans_data.OpenRecordset(objRecordset3, strSql, o_error)
                
                If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
                    Exit Sub
                End If
                    'cycle through inventory items until material requirement
                    'has been completely relieved
                Do While lngTotAdjQty < lngAdjQty And _
                    lngTotMRAdjQty < lngMatReqReadyQty And _
                    objRecordset3.EOF = False And _
                    Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
                
                    Dim lngCurrAdjQty As Long
                    Dim lngAvailInvQty As Long
                    Dim objInvAdj As Object
                    lngAvailInvQty = objRecordset3("inv_qty")
                    
                    If (lngAdjQty - lngTotAdjQty) > lngAvailInvQty Then
                        If (lngMatReqReadyQty - lngTotMRAdjQty) > lngAvailInvQty Then
                            lngCurrAdjQty = lngAvailInvQty
                        Else
                            lngCurrAdjQty = lngMatReqReadyQty - lngTotMRAdjQty
                        End If
                    Else
                        If (lngAdjQty - lngTotAdjQty) > (lngMatReqReadyQty - lngTotMRAdjQty) Then
                            lngCurrAdjQty = (lngMatReqReadyQty - lngTotMRAdjQty)
                        Else
                            lngCurrAdjQty = (lngAdjQty - lngTotAdjQty)
                        End If
                    End If
                    
                    Set objInvAdj = CreateObject("converge_inv.inventory_adj")
                    Set objInvAdj.o_error = Me.o_error
                    Set objInvAdj.o_trans_data = Me.o_trans_data
                    objInvAdj.p_item_id = lngItemId
                    objInvAdj.p_lot_id = objRecordset3("lot_id")
                    objInvAdj.p_adj_qty = lngCurrAdjQty * -1
                    objInvAdj.p_mat_req_id = objRecordset2("mat_Req_id")
                    objInvAdj.p_bin_id = objRecordset3("bin_id")
                    objInvAdj.p_phy_inv_id = 3110
                    objInvAdj.p_inv_loc_id = "Phoenix"
                    Call objInvAdj.Add
                        
                    lngTotAdjQty = lngTotAdjQty + lngCurrAdjQty
                    lngTotMRAdjQty = lngTotMRAdjQty + lngCurrAdjQty
                    objRecordset3.MoveNext
                Loop
                objRecordset2.MoveNext
            Loop
            
                'If the shipped quantity was not consumed against material requirements
                'consume the remainder of the inventory against the physical inventory order
            If lngTotAdjQty < lngAdjQty Then
                strSql = "SELECT SUM(adj_qty) as inv_qty, lot_id, bin_id " _
                    & "From inventory_adj " _
                    & "WHERE inv_loc_id = 'Phoenix' AND item_id = " & lngItemId & " " _
                    & "GROUP BY lot_id, bin_id, item_id " _
                    & "Having Sum(adj_qty) > 0 "
                
                Call o_trans_data.OpenRecordset(objRecordset3, strSql, o_error)
                
                If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
                    Exit Sub
                End If
                    'cycle through inventory items until material requirement
                    'has been completely relieved
                Do While lngTotAdjQty < lngAdjQty And _
                    objRecordset3.EOF = False And _
                    Me.o_error.p_type_cd <> "E" And Me.o_error.p_type_cd <> "F"
                    
                    lngAvailInvQty = objRecordset3("inv_qty")
                    
                    If (lngAdjQty - lngTotAdjQty) > lngAvailInvQty Then
                       lngCurrAdjQty = lngAvailInvQty
                    Else
                       lngCurrAdjQty = (lngAdjQty - lngTotAdjQty)
                    End If
                    
                    Set objInvAdj = CreateObject("converge_inv.inventory_adj")
                    Set objInvAdj.o_error = Me.o_error
                    Set objInvAdj.o_trans_data = Me.o_trans_data
                    objInvAdj.p_item_id = lngItemId
                    objInvAdj.p_lot_id = objRecordset3("lot_id")
                    objInvAdj.p_adj_qty = lngCurrAdjQty * -1
                    objInvAdj.p_bin_id = objRecordset3("bin_id")
                    objInvAdj.p_phy_inv_id = 3110
                    objInvAdj.p_inv_loc_id = "Phoenix"
                    Call objInvAdj.Add
                
                    lngTotAdjQty = lngTotAdjQty + lngCurrAdjQty
                    objRecordset3.MoveNext
                Loop
            End If
        End If
        Call objRecordset1.MoveNext
    Loop
    
    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "7000"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.clear
End Sub
