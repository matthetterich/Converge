VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_varItem_id As Variant
Private m_varItem_nbr As Variant
Private m_varItem_desc As Variant
Private m_varItem_price As Variant
Private m_varItem_cost As Variant
Private m_varLast_item_cost As Variant
Private m_varLast_item_price As Variant
Private m_varCost_cd As Variant
Private m_varCost_desc As Variant
Private m_varassem_cost As Variant
Private m_varRevision_Nbr As Variant
Private m_varType_cd As Variant
Private m_varType_desc As Variant
Private m_varCategory_cd As Variant
Private m_varCategory_desc As Variant
Private m_varUom_cd As Variant
Private m_varUom_desc As Variant
Private m_varStnd_carton_qty As Variant
Private m_varAbc_cd As Variant
Private m_varAbc_desc As Variant
Private m_varPo_increment_qty As Variant
Private m_varDurometer As Variant
Private m_varPo_min_qty As Variant
Private m_varPo_max_qty As Variant
Private m_varSafety_stock_qty As Variant
Private m_varLead_days As Variant
Private m_varReady_release_days As Variant
Private m_varFirm_plan_days As Variant
Private m_varCommodity_cd As Variant
Private m_varCommodity_desc As Variant
Private m_varInspection_days As Variant
Private m_varSupplier_lead_days As Variant
Private m_varPreferred_supplier_id As Variant
Private m_varItem_text As Variant
Private m_varSelling_instr_tx As Variant
Private m_varPurch_instr_tx As Variant
Private m_varContainerLbs As Variant
Private m_objPreferred_supplier As Object
Private m_objError As Object
Private m_objTrans_data As Object
Private m_varcust_revision_nbr As Variant
Private mvaro_record_set As adodb.Recordset
Private m_varMfgProcess_cd As Variant
Private m_varMfgProcess_desc As Variant
Private m_varPartsPerHour_qty As Variant
Private m_varPartsPerSkid As Variant
Private m_varPlanInclFl As Variant

'

Public Property Let p_cust_revision_nbr(ByVal vData As Variant)
    m_varcust_revision_nbr = vData
End Property
Public Property Get p_cust_revision_nbr() As Variant
    p_cust_revision_nbr = m_varcust_revision_nbr
End Property

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
End Property
Public Property Let p_durometer(ByVal vData As Variant)
    m_varDurometer = vData
End Property
Public Property Get p_durometer() As Variant
    p_durometer = m_varDurometer
End Property
Public Property Let p_revision_nbr(ByVal vData As Variant)
    m_varRevision_Nbr = vData
End Property
Public Property Get p_revision_nbr() As Variant
    p_revision_nbr = m_varRevision_Nbr
End Property
Public Property Let p_item_price(ByVal vData As Variant)
    m_varItem_price = vData
End Property
Public Property Get p_item_price() As Variant
    p_item_price = m_varItem_price
End Property
Public Property Let p_item_cost(ByVal vData As Variant)
    m_varItem_cost = vData
End Property
Public Property Get p_item_cost() As Variant
    p_item_cost = m_varItem_cost
End Property
Public Property Let p_last_item_cost(ByVal vData As Variant)
    m_varLast_item_cost = vData
End Property
Public Property Get p_last_item_cost() As Variant
    p_last_item_cost = m_varLast_item_cost
End Property
Public Property Let p_last_item_price(ByVal vData As Variant)
    m_varLast_item_price = vData
End Property
Public Property Get p_last_item_price() As Variant
    p_last_item_price = m_varLast_item_price
End Property
Public Property Let p_cost_cd(ByVal vData As Variant)
    m_varCost_cd = vData
End Property
Public Property Get p_cost_cd() As Variant
    p_cost_cd = m_varCost_cd
End Property
Public Property Let p_cost_desc(ByVal vData As Variant)
    m_varCost_desc = vData
End Property
Public Property Get p_cost_desc() As Variant
    p_cost_desc = m_varCost_desc
End Property
Public Property Let p_assem_cost(ByVal vData As Variant)
    m_varassem_cost = vData
End Property
Public Property Get p_assem_cost() As Variant
    p_assem_cost = m_varassem_cost
End Property

Public Property Let p_commodity_desc(ByVal vData As Variant)
    m_varCommodity_desc = vData
End Property
Public Property Get p_commodity_desc() As Variant
    p_commodity_desc = m_varCommodity_desc
End Property
Public Property Let p_category_desc(ByVal vData As Variant)
    m_varCategory_desc = vData
End Property
Public Property Get p_category_desc() As Variant
    p_category_desc = m_varCategory_desc
End Property
Public Property Let p_uom_desc(ByVal vData As Variant)
    m_varUom_desc = vData
End Property
Public Property Get p_uom_desc() As Variant
    p_uom_desc = m_varUom_desc
End Property
Public Property Let p_abc_desc(ByVal vData As Variant)
    m_varAbc_desc = vData
End Property
Public Property Get p_abc_desc() As Variant
    p_abc_desc = m_varAbc_desc
End Property
Public Property Let p_type_desc(ByVal vData As Variant)
    m_varType_desc = vData
End Property
Public Property Get p_type_desc() As Variant
    p_type_desc = m_varType_desc
End Property
Public Property Let p_preferred_supplier_id(ByVal vData As Variant)
    m_varPreferred_supplier_id = vData
End Property
Public Property Get p_preferred_supplier_id() As Variant
    p_preferred_supplier_id = m_varPreferred_supplier_id
End Property
Public Property Let p_supplier_lead_days(ByVal vData As Variant)
    m_varSupplier_lead_days = vData
End Property
Public Property Get p_supplier_lead_days() As Variant
    p_supplier_lead_days = m_varSupplier_lead_days
End Property
Public Property Let p_inspection_days(ByVal vData As Variant)
    m_varInspection_days = vData
End Property
Public Property Get p_inspection_days() As Variant
    p_inspection_days = m_varInspection_days
End Property
Public Property Let p_commodity_cd(ByVal vData As Variant)
    m_varCommodity_cd = vData
End Property
Public Property Get p_commodity_cd() As Variant
    p_commodity_cd = m_varCommodity_cd
End Property
Public Property Let p_firm_plan_days(ByVal vData As Variant)
    m_varFirm_plan_days = vData
End Property
Public Property Get p_firm_plan_days() As Variant
    p_firm_plan_days = m_varFirm_plan_days
End Property
Public Property Let p_ready_release_days(ByVal vData As Variant)
    m_varReady_release_days = vData
End Property
Public Property Get p_ready_release_days() As Variant
    p_ready_release_days = m_varReady_release_days
End Property
Public Property Let p_lead_days(ByVal vData As Variant)
    m_varLead_days = vData
End Property
Public Property Get p_lead_days() As Variant
    p_lead_days = m_varLead_days
End Property
Public Property Let p_safety_stock_qty(ByVal vData As Variant)
    m_varSafety_stock_qty = vData
End Property
Public Property Get p_safety_stock_qty() As Variant
    p_safety_stock_qty = m_varSafety_stock_qty
End Property
Public Property Let p_po_min_qty(ByVal vData As Variant)
    m_varPo_min_qty = vData
End Property
Public Property Get p_po_min_qty() As Variant
    p_po_min_qty = m_varPo_min_qty
End Property
Public Property Let p_po_max_qty(ByVal vData As Variant)
    m_varPo_max_qty = vData
End Property
Public Property Get p_po_max_qty() As Variant
    p_po_max_qty = m_varPo_max_qty
End Property
Public Property Let p_po_increment_qty(ByVal vData As Variant)
    m_varPo_increment_qty = vData
End Property
Public Property Get p_po_increment_qty() As Variant
    p_po_increment_qty = m_varPo_increment_qty
End Property
Public Property Let p_abc_cd(ByVal vData As Variant)
    m_varAbc_cd = vData
End Property
Public Property Get p_abc_cd() As Variant
    p_abc_cd = m_varAbc_cd
End Property
Public Property Let p_stnd_carton_qty(ByVal vData As Variant)
    m_varStnd_carton_qty = vData
End Property
Public Property Get p_stnd_carton_qty() As Variant
    p_stnd_carton_qty = m_varStnd_carton_qty
End Property
Public Property Let p_category_cd(ByVal vData As Variant)
    m_varCategory_cd = vData
End Property
Public Property Get p_category_cd() As Variant
    p_category_cd = m_varCategory_cd
End Property
Public Property Let p_uom_cd(ByVal vData As Variant)
    m_varUom_cd = vData
End Property
Public Property Get p_uom_cd() As Variant
    p_uom_cd = m_varUom_cd
End Property
Public Property Let p_type_cd(ByVal vData As Variant)
    m_varType_cd = vData
End Property
Public Property Get p_type_cd() As Variant
    p_type_cd = m_varType_cd
End Property
Public Property Let p_item_desc(ByVal vData As Variant)
    m_varItem_desc = vData
End Property
Public Property Get p_item_desc() As Variant
    p_item_desc = m_varItem_desc
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

Public Property Let p_item_text(ByVal vData As Variant)
    m_varItem_text = vData
End Property
Public Property Get p_item_text() As Variant
    p_item_text = m_varItem_text
End Property

Public Property Let p_selling_instr_tx(ByVal vData As Variant)
    m_varSelling_instr_tx = vData
End Property
Public Property Get p_selling_instr_tx() As Variant
    p_selling_instr_tx = m_varSelling_instr_tx
End Property

Public Property Let p_container_lbs(ByVal vData As Variant)
    m_varContainerLbs = vData
End Property
Public Property Get p_container_lbs() As Variant
    p_container_lbs = m_varContainerLbs
End Property

Public Property Let p_purch_instr_tx(ByVal vData As Variant)
    m_varPurch_instr_tx = vData
End Property
Public Property Get p_purch_instr_tx() As Variant
    p_purch_instr_tx = m_varPurch_instr_tx
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
Public Property Set o_preferred_supplier(ByVal vData As Object)
    Set m_objPreferred_supplier = vData
End Property
Public Property Get o_preferred_supplier() As Object
    Set o_preferred_supplier = m_objPreferred_supplier
End Property

Public Property Let p_mfg_process_cd(ByVal vData As Variant)
    m_varMfgProcess_cd = vData
End Property
Public Property Get p_mfg_process_cd() As Variant
    p_mfg_process_cd = m_varMfgProcess_cd
End Property

Public Property Let p_mfg_process_desc(ByVal vData As Variant)
    m_varMfgProcess_desc = vData
End Property
Public Property Get p_mfg_process_desc() As Variant
    p_mfg_process_desc = m_varMfgProcess_desc
End Property
Public Property Let p_parts_per_hour_qty(ByVal vData As Variant)
    m_varPartsPerHour_qty = vData
End Property
Public Property Get p_parts_per_hour_qty() As Variant
    p_parts_per_hour_qty = m_varPartsPerHour_qty
End Property

Public Property Let p_parts_per_skid(ByVal vData As Variant)
    m_varPartsPerSkid = vData
End Property
Public Property Get p_parts_per_skid() As Variant
    p_parts_per_skid = m_varPartsPerSkid
End Property
Public Property Let p_plan_incl_fl(ByVal vData As Variant)
    m_varPlanInclFl = vData
End Property
Public Property Get p_plan_incl_fl() As Variant
    p_plan_incl_fl = m_varPlanInclFl
End Property



Public Sub GetItem()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "Item.cls.GetItem"

    If IsNull(p_item_id) = True Then
        strSql = "SELECT * "
        strSql = strSql & "FROM Item "
        strSql = strSql & "WHERE "
        strSql = strSql & " " & fWhereVariable("item_nbr", p_item_nbr, "S", "=")
    Else
        strSql = "SELECT * "
        strSql = strSql & "FROM Item "
        strSql = strSql & "WHERE "
        strSql = strSql & " " & fWhereVariable("Item_id", p_item_id, "N", "=")
    End If
        
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_item_id = Null
        objRecordset.Close
        Set objRecordset = Nothing
        
    Else
        ' Load in properties
        p_item_id = objRecordset("Item_id")
        p_item_nbr = objRecordset("Item_nbr")
        p_item_price = objRecordset("Item_price")
        p_item_cost = objRecordset("item_cost")
        p_assem_cost = objRecordset("assem_cost")
        p_item_desc = objRecordset("item_desc")
        p_type_cd = objRecordset("type_cd")
        p_uom_cd = objRecordset("uom_cd")
        p_category_cd = objRecordset("category_cd")
        p_stnd_carton_qty = objRecordset("stnd_carton_qty")
        p_abc_cd = objRecordset("abc_cd")
        p_po_increment_qty = objRecordset("po_increment_qty")
        p_po_min_qty = objRecordset("po_min_qty")
        p_po_max_qty = objRecordset("po_max_qty")
        p_safety_stock_qty = objRecordset("safety_stock_qty")
        p_lead_days = objRecordset("lead_days")
        p_ready_release_days = objRecordset("ready_release_days")
        p_firm_plan_days = objRecordset("firm_plan_days")
        p_commodity_cd = objRecordset("commodity_cd")
        p_inspection_days = objRecordset("inspection_days")
        p_supplier_lead_days = objRecordset("supplier_lead_days")
        p_preferred_supplier_id = objRecordset("preferred_supplier_id")
        p_durometer = objRecordset("durometer")
        p_revision_nbr = objRecordset("revision_nbr")
        p_cust_revision_nbr = objRecordset("cust_revision_nbr")
        p_item_text = objRecordset("item_tx")
        p_selling_instr_tx = objRecordset("selling_instr_tx")
        p_purch_instr_tx = objRecordset("purch_instr_tx")
        p_container_lbs = objRecordset("container_lbs")
        p_last_item_cost = objRecordset("last_cost")
        p_last_item_price = objRecordset("last_price")
        p_cost_cd = objRecordset("cost_cd")
        p_parts_per_hour_qty = objRecordset("parts_per_hour_qty")
        p_mfg_process_cd = objRecordset("mfg_process_cd")
        p_parts_per_skid = objRecordset("parts_per_skid")
        p_plan_incl_fl = objRecordset("plan_incl_fl")
                
        objRecordset.Close
        Set objRecordset = Nothing
        
        ' Load in Descriptions
        p_uom_desc = fGetDesc("uom_cd", p_uom_cd, o_error, o_trans_data)
        p_category_desc = fGetDesc("item_category_cd", p_category_cd, o_error, o_trans_data)
        p_type_desc = fGetDesc("item_type_cd", p_type_cd, o_error, o_trans_data)
        p_abc_desc = fGetDesc("item_abc_cd", p_abc_cd, o_error, o_trans_data)
        p_commodity_desc = fGetDesc("item_commodity_cd", p_commodity_cd, o_error, o_trans_data)
        p_cost_desc = fGetDesc("cost_cd", p_cost_cd, o_error, o_trans_data)
        Me.p_mfg_process_desc = fGetDesc("item_mfg_proc_cd", p_mfg_process_cd, o_error, o_trans_data)
         
        Set o_preferred_supplier = CreateObject("converge_purchase.supplier")
        If Len(p_preferred_supplier_id) > 0 Then
            If p_preferred_supplier_id > 0 Then
                Set o_preferred_supplier.o_error = o_error
                Set o_preferred_supplier.o_trans_data = o_trans_data
                o_preferred_supplier.p_supplier_id = p_preferred_supplier_id
                Call o_preferred_supplier.GetSupplier
            End If
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

Public Sub Add()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Item.cls.AddItem"
    
    Call validate("add")
    p_item_id = fGetId("item_id", o_error, o_trans_data)
    Call IssueInsert
    
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_item_id)
    
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


Public Sub Change()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim aItem As item
    strRoutine_nm = "Item.cls.ChangeItem"
    
    Set aItem = New item
    aItem.p_item_id = Me.p_item_id
    aItem.p_item_nbr = Me.p_item_nbr
    Set aItem.o_error = Me.o_error
    Set aItem.o_trans_data = Me.o_trans_data
    
    aItem.GetItem
    
    Call validate("change")
    Call IssueUpdate(aItem.p_type_cd)
    If aItem.p_type_cd = "AS" Then
        'if the item is an assembled item, recalculate and update the assembled item cost
        Call updateAssembledItemCost(Me.o_error, Me.o_trans_data, Me.p_item_id, 0)
    Else
        'if the item is not an assembled item, update it's parent items assebled cost.  If it
        'is not an assebly component, nothing will happen
        Call updateParentAssembledItemCost(Me.o_error, Me.o_trans_data, Me.p_item_id, 0)
    End If
    
    
    Call InsertQBTrans(Me.o_error, Me.o_trans_data, Me.p_item_id)
    
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

Public Sub Delete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Item.cls.DeleteItem"
    Dim strSql As String
    
    Call IssueDelete
    
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


Public Function getPreferredSupplierId(ByVal v_varItemId As Variant, ByRef r_objError As Object, ByRef r_objTrans_data As Object)
    
    Set o_error = r_objError
    Set o_trans_data = r_objTrans_data

    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    strRoutine_nm = "Item.cls.getPreferredSupplierId"
    
    If v_varItemId > 0 Then
    Else
        getPreferredSupplierId = Null
        Exit Function
    End If
    
    strSql = "SELECT preferred_supplier_id "
    strSql = strSql & "FROM Item "
    strSql = strSql & "WHERE "
    strSql = strSql & " " & fWhereVariable("item_id", v_varItemId, "N", "=")

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If objRecordset.EOF = False Then
        getPreferredSupplierId = objRecordset("preferred_supplier_id")
    Else
        getPreferredSupplierId = Null
    End If
    
End Function
Private Sub validate(sFunction_cd)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Item.cls.validate"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    
    If fValuePresent(o_preferred_supplier.p_supplier_nbr) = True Then
        p_preferred_supplier_id = fGetSupplierId(o_preferred_supplier.p_supplier_nbr, o_error, o_trans_data)
        If IsNull(p_preferred_supplier_id) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0300"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3400
            Exit Sub
        End If
    End If
      
    If IsNull(o_preferred_supplier.p_supplier_nbr) = True Then
        p_preferred_supplier_id = Null
    End If
    
    If fValuePresent(p_commodity_cd) = True Then
        If IsNull(fGetDesc("commodity_cd", p_commodity_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "e"
            o_error.err_cd = "0310"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3401
            Exit Sub
        End If
    End If
      
    If fValuePresent(Me.p_mfg_process_cd) = True Then
        If IsNull(fGetDesc("item_mfg_proc_cd", p_mfg_process_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "e"
            o_error.err_cd = "0315"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            o_error.p_desc = "Mfg Processing Code is invalid: " & Me.p_mfg_process_cd
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_category_cd) = True Then
        If IsNull(fGetDesc("item_category_cd", p_category_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "e"
            o_error.err_cd = "0320"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3402
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_abc_cd) = True Then
       If IsNull(fGetDesc("abc_cd", p_abc_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "e"
            o_error.err_cd = "0330"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3403
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_uom_cd) = True Then
       If IsNull(fGetDesc("uom_cd", p_uom_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "e"
            o_error.err_cd = "0340"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3404
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_safety_stock_qty) = True Then
       If IsNumeric(p_safety_stock_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0350"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3405
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_po_max_qty) = True Then
       If IsNumeric(p_po_max_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0360"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3406
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_po_min_qty) = True Then
       If IsNumeric(p_po_min_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0370"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3407
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_po_increment_qty) = True Then
       If IsNumeric(p_po_increment_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0380"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3408
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_stnd_carton_qty) = True Then
       If IsNumeric(p_stnd_carton_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0390"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3409
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_parts_per_hour_qty) = True Then
       If IsNumeric(p_parts_per_hour_qty) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0392"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            o_error.p_desc = "parts per hour must be numeric"
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_parts_per_skid) = True Then
       If IsNumeric(p_parts_per_skid) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0393"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            o_error.p_desc = "parts per skid must be numeric"
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_plan_incl_fl) = True Then
       If Not (p_plan_incl_fl = "Y" Or p_plan_incl_fl = "N") Then
            o_error.type_cd = "F"
            o_error.err_cd = "0394"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            o_error.p_desc = "p_plan_incl_fl must equal 'Y' or 'N'"
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_container_lbs) = True Then
       If IsNumeric(p_container_lbs) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0391"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            o_error.p_desc = "Box Weight must be numeric"
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_supplier_lead_days) = True Then
       If IsNumeric(p_supplier_lead_days) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0400"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 34010
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_inspection_days) = True Then
       If IsNumeric(p_inspection_days) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0410"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3411
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_firm_plan_days) = True Then
       If IsNumeric(p_firm_plan_days) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0420"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3412
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_ready_release_days) = True Then
       If IsNumeric(p_ready_release_days) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0440"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3413
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_lead_days) = True Then
       If IsNumeric(p_lead_days) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0450"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3415
            Exit Sub
        End If
    End If
      
    If fValuePresent(p_item_price) = True Then
       If IsNumeric(p_item_price) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0460"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3416
            Exit Sub
        End If
    End If
    
    If fValuePresent(p_revision_nbr) = True Then
        If IsNumeric(p_revision_nbr) = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0460"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3420
            Exit Sub
        End If
    End If
      
    If sFunction_cd = "add" Then
        If fValuePresent(p_item_nbr) = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "0470"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3417
            Exit Sub
        End If
    
        If InStr(1, p_item_nbr, "&", vbBinaryCompare) > 0 Then
            o_error.type_cd = "E"
            o_error.err_cd = "0475"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 1078
            o_error.p_desc = "The '&' character is not valid in item number"
            Exit Sub
        End If
    
        strSql = "SELECT item_nbr "
        strSql = strSql & "FROM item "
        strSql = strSql & "WHERE         "
        strSql = strSql & "   " & fWhereVariable("item_nbr", p_item_nbr, "S", "=")
        
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            objRecordset.Close
            Set objRecordset = Nothing
            Exit Sub
        End If
        
        If objRecordset.EOF = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "0480"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3418
            objRecordset.Close
            Set objRecordset = Nothing
            Exit Sub
        End If
        
        objRecordset.Close
        Set objRecordset = Nothing
    End If
    
    If sFunction_cd = ("change") Then
        If IsNull(p_item_nbr) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0490"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3417
            Exit Sub
        End If
        
        If fValuePresent(p_item_nbr) = True Then
            strSql = "SELECT item_nbr "
            strSql = strSql & "FROM item "
            strSql = strSql & "WHERE         "
            strSql = strSql & "   " & fWhereVariable("item_nbr", p_item_nbr, "S", "=")
            strSql = strSql & "and " & fWhereVariable("item_id", p_item_id, "N", "<>")
        
            Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
            
            If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
                objRecordset.Close
                Set objRecordset = Nothing
                Exit Sub
            End If
            
            If objRecordset.EOF = False Then
                o_error.type_cd = "F"
                o_error.err_cd = "0500"
                o_error.routine_nm = strRoutine_nm
                o_error.message_id = 3418
                objRecordset.Close
                Set objRecordset = Nothing
                Exit Sub
            End If
            
            objRecordset.Close
            Set objRecordset = Nothing
        End If
    End If
    
    If sFunction_cd = ("change") Then
        If IsNull(p_item_desc) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0510"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3419
        End If
    End If
    
    If sFunction_cd = ("add") Then
        If fValuePresent(p_item_desc) = False Then
            o_error.type_cd = "E"
            o_error.err_cd = "0520"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3419
            Exit Sub
        End If
    End If

    If sFunction_cd = ("add") Then
        If IsNull(fGetDesc("item_type_cd", p_type_cd, o_error, o_trans_data)) = True Then
            o_error.type_cd = "E"
            o_error.err_cd = "0530"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 3410
            Exit Sub
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



Private Sub IssueInsert()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Item.cls.IssueInsert"

    strSql = "insert into Item ( "
    strSql = strSql & " Item_id "
    strSql = strSql & ",Item_nbr "
    strSql = strSql & ",item_price "
    strSql = strSql & ",item_cost "
    strSql = strSql & ",assem_cost "
    strSql = strSql & ",durometer "
    strSql = strSql & ",revision_nbr "
    strSql = strSql & ",item_desc "
    strSql = strSql & ",type_cd "
    strSql = strSql & ",category_cd "
    strSql = strSql & ",uom_cd "
    strSql = strSql & ",stnd_carton_qty "
    strSql = strSql & ",abc_cd "
    strSql = strSql & ",po_increment_qty "
    strSql = strSql & ",po_min_qty "
    strSql = strSql & ",po_max_qty "
    strSql = strSql & ",safety_stock_qty "
    strSql = strSql & ",lead_days "
    strSql = strSql & ",ready_release_days "
    strSql = strSql & ",firm_plan_days "
    strSql = strSql & ",commodity_cd "
    strSql = strSql & ",supplier_lead_days "
    strSql = strSql & ",inspection_days "
    strSql = strSql & ",preferred_supplier_id "
    strSql = strSql & ",item_tx "
    strSql = strSql & ",selling_instr_tx "
    strSql = strSql & ",purch_instr_tx "
    strSql = strSql & ",container_lbs "
    strSql = strSql & ",cust_revision_nbr "
    strSql = strSql & ",last_cost "
    strSql = strSql & ",cost_cd "
    strSql = strSql & ",last_price "
    strSql = strSql & ",parts_per_hour_qty "
    strSql = strSql & ",mfg_process_cd "
    strSql = strSql & ",parts_per_skid "
    strSql = strSql & ",plan_incl_fl "
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_item_price, "N")
    strSql = strSql & "," & fInsertVariable(p_item_cost, "N")
    strSql = strSql & "," & fInsertVariable(p_assem_cost, "N")
    strSql = strSql & "," & fInsertVariable(p_durometer, "S")
    strSql = strSql & ",1"
    strSql = strSql & "," & fInsertVariable(p_item_desc, "S")
    strSql = strSql & "," & fInsertVariable(p_type_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_category_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_uom_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_stnd_carton_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_abc_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_po_increment_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_po_min_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_po_max_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_safety_stock_qty, "N")
    strSql = strSql & "," & fInsertVariable(p_lead_days, "N")
    strSql = strSql & "," & fInsertVariable(p_ready_release_days, "N")
    strSql = strSql & "," & fInsertVariable(p_firm_plan_days, "N")
    strSql = strSql & "," & fInsertVariable(p_commodity_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_lead_days, "N")
    strSql = strSql & "," & fInsertVariable(p_inspection_days, "N")
    strSql = strSql & "," & fInsertVariable(p_preferred_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_text, "S")
    strSql = strSql & "," & fInsertVariable(p_selling_instr_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_purch_instr_tx, "S")
    strSql = strSql & "," & fInsertVariable(p_container_lbs, "N")
    strSql = strSql & "," & fInsertVariable(p_cust_revision_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_last_item_cost, "N")
    strSql = strSql & "," & fInsertVariable(p_cost_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_last_item_price, "N")
    strSql = strSql & "," & fInsertVariable(p_parts_per_hour_qty, "N")
    strSql = strSql & "," & fInsertVariable(Me.p_mfg_process_cd, "S")
    strSql = strSql & "," & fInsertVariable(p_parts_per_skid, "N")
    strSql = strSql & "," & fInsertVariable(p_plan_incl_fl, "S")
    strSql = strSql & ")"

    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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


Private Sub IssueUpdate(ByVal v_strItemTypeCd As String)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim dblUpdateQty As Double
    Dim strSql As String
    strRoutine_nm = "Item.cls.IssueUpdate"

    strSql = "update item Set "
    strSql = strSql & " " & fUpdateVariable("item_price", p_item_price, "N", "=")
        'only upate the cost on purchased items.  Assembled items are updated based on BOM
    If v_strItemTypeCd = "PU" Then
        strSql = strSql & "," & fUpdateVariable("item_cost", p_item_cost, "N", "=")
    End If
    strSql = strSql & "," & fUpdateVariable("assem_cost", p_assem_cost, "N", "=")
    strSql = strSql & "," & fUpdateVariable("durometer", p_durometer, "S", "=")
    strSql = strSql & ", revision_nbr = revision_nbr + 1 "
    strSql = strSql & "," & fUpdateVariable("item_nbr", p_item_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("item_desc", p_item_desc, "S", "=")
    strSql = strSql & "," & fUpdateVariable("uom_cd", p_uom_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("category_cd", p_category_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("stnd_carton_qty", p_stnd_carton_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("abc_cd", p_abc_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("po_increment_qty", p_po_increment_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("po_min_qty", p_po_min_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("po_max_qty", p_po_max_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("safety_stock_qty", p_safety_stock_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("lead_days", p_lead_days, "N", "=")
    strSql = strSql & "," & fUpdateVariable("ready_release_days", p_ready_release_days, "N", "=")
    strSql = strSql & "," & fUpdateVariable("firm_plan_days", p_firm_plan_days, "N", "=")
    strSql = strSql & "," & fUpdateVariable("commodity_cd", p_commodity_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("inspection_days", p_inspection_days, "N", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_lead_days", p_supplier_lead_days, "N", "=")
    strSql = strSql & "," & fUpdateVariable("preferred_supplier_id", p_preferred_supplier_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("item_tx", p_item_text, "S", "=")
    strSql = strSql & "," & fUpdateVariable("selling_instr_tx", p_selling_instr_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("purch_instr_tx", p_purch_instr_tx, "S", "=")
    strSql = strSql & "," & fUpdateVariable("container_lbs", p_container_lbs, "N", "=")
    strSql = strSql & "," & fUpdateVariable("cust_revision_nbr", p_cust_revision_nbr, "S", "=")
    strSql = strSql & "," & fUpdateVariable("last_cost", p_last_item_cost, "N", "=")
    strSql = strSql & "," & fUpdateVariable("last_price", p_last_item_price, "N", "=")
    strSql = strSql & "," & fUpdateVariable("cost_cd", p_cost_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("parts_per_hour_qty", p_parts_per_hour_qty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("parts_per_skid", p_parts_per_skid, "N", "=")
    strSql = strSql & "," & fUpdateVariable("mfg_process_cd", Me.p_mfg_process_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("plan_incl_fl", Me.p_plan_incl_fl, "S", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("Item_id", p_item_id, "N", "=")
                                                
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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


Private Sub IssueDelete()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    strRoutine_nm = "Item.cls.IssueDelete"

    strSql = "delete from Item "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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


Public Sub getItems()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Items.cls.GetItems"
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    Dim key As Double
    Dim lngCount As Long
    Dim strItemNbr As String
    Dim varSupplier_id As Variant
    
    strSql = "select * from item where 1 = 1"
    
    If IsNull(p_item_nbr) = True Or p_item_nbr = "" Then
    Else
       lngCount = InStr(1, p_item_nbr, "*", vbBinaryCompare)
       If lngCount > 0 Then
            strItemNbr = Left(p_item_nbr, lngCount - 1)
            strSql = strSql & " AND item_nbr LIKE '" & strItemNbr & "%'"
       Else
            strSql = strSql & "  AND " & fWhereVariable("item_nbr", p_item_nbr, "S", "=")
       End If
    End If
    
    If Len(p_type_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("type_cd", p_type_cd, "S", "=")
    End If

    If Len(p_category_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("category_cd", p_category_cd, "S", "=")
    End If

    If Len(p_abc_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("abc_cd", p_abc_cd, "S", "=")
    End If

    If Len(p_commodity_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("commodity_cd", p_commodity_cd, "S", "=")
    End If

    If Len(Me.p_mfg_process_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("mfg_process_cd", Me.p_mfg_process_cd, "S", "=")
    End If

    If Len(p_uom_cd) > 0 Then
        strSql = strSql & " and " & fWhereVariable("uom_cd", p_uom_cd, "S", "=")
    End If

    If Len(o_preferred_supplier.p_supplier_nbr) > 0 Then
        varSupplier_id = fGetSupplierId(o_preferred_supplier.p_supplier_nbr, o_error, o_trans_data)
        strSql = strSql & " and " & fWhereVariable("preferred_supplier_id", varSupplier_id, "N", "=")
    End If

    strSql = strSql & "Order By item_nbr "

    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Set o_record_set = objRecordset
    
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

Public Sub InsertQBTrans( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByVal v_varItemId As Variant _
        )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = "Converge_item" & "." & "item" & "." & "InsertQBTrans"
    
    Dim strSqlTx As String
    Dim dblUpdateQty As Double
    
    strSqlTx = "Insert Into quickbooks_trans ( "
    strSqlTx = strSqlTx & " trans_type_cd "
    strSqlTx = strSqlTx & ", item_id "
    strSqlTx = strSqlTx & ", insert_ts "
    strSqlTx = strSqlTx & " ) "
    strSqlTx = strSqlTx & "values ("
    strSqlTx = strSqlTx & "  3 "
    strSqlTx = strSqlTx & "," & v_varItemId & " "
    strSqlTx = strSqlTx & ", GetDate()  "
    strSqlTx = strSqlTx & ") "
   
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub updateAssembledItemCost( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByVal v_varItemId As Variant _
        , ByVal v_lngRecursionCt As Long _
        )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = "Converge_item" & "." & "item" & "." & "updateAssembledItemCost"
    Debug.Print strRoutineNm & v_varItemId
    Dim strSqlTx As String
    Dim dblUpdateQty As Double
    
    strSqlTx = "UPDATE item SET item_cost = ( "
    strSqlTx = strSqlTx & " SELECT SUM(material_cost) AS material_cost "
    strSqlTx = strSqlTx & " FROM ( "
    strSqlTx = strSqlTx & "    SELECT bill_of_material.item_qty * comp_item.item_cost + comp_item.assem_cost AS material_cost "
    strSqlTx = strSqlTx & "    From bill_of_material "
    strSqlTx = strSqlTx & "    INNER JOIN item comp_item ON bill_of_material.comp_item_id = comp_item.item_id "
    strSqlTx = strSqlTx & "    WHERE (" & fWhereVariable("bill_of_material.item_id", v_varItemId, "N", "=") & ")) x) "
    strSqlTx = strSqlTx & " Where " & fWhereVariable("item_id", v_varItemId, "N", "=") & " "
    Call r_objTransData.ExecuteDatabase(strSqlTx, dblUpdateQty, r_objError)
    Call updateParentAssembledItemCost(r_objError, r_objTransData, v_varItemId, v_lngRecursionCt)
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub updateParentAssembledItemCost( _
          ByRef r_objError As Object _
        , ByRef r_objTransData As Object _
        , ByVal v_varItemId As Variant _
        , ByVal v_lngRecursionCt As Long _
        )

    On Error GoTo ErrHandler
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutineNm As String
    strRoutineNm = "Converge_item" & "." & "item" & "." & "updateParentAssembledItemCost"
    Debug.Print strRoutineNm & v_varItemId
    Debug.Print v_lngRecursionCt
    If (v_lngRecursionCt > 15) Then
        r_objError.type_cd = "F"
        r_objError.err_cd = "0510"
        r_objError.routine_nm = strRoutineNm
        r_objError.message_id = 0
        r_objError.p_desc = "Max recursion for assembly cost update.  We are potentially in an infinite loop"
        Exit Sub
    End If
    
    Dim strSqlTx As String
    Dim lngItemid As Long
    Dim objRecordset As adodb.Recordset
    Dim lngRecursionCt As Long
    lngRecursionCt = v_lngRecursionCt + 1
    
    strSqlTx = " SELECT item_id " & _
               " From bill_of_material " & _
               " WHERE (" & fWhereVariable("comp_item_id", v_varItemId, "N", "=") & ") " & _
               " group by item_id "
    Call o_trans_data.OpenRecordset(objRecordset, strSqlTx, o_error)
    
    Do While objRecordset.EOF = False
        lngItemid = objRecordset("item_id")
        Call updateAssembledItemCost(r_objError, r_objTransData, lngItemid, lngRecursionCt)
        If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
           Exit Sub
        End If
        
        Call objRecordset.MoveNext
    Loop
    
    Exit Sub
ErrHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0600"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutineNm
      .p_message_id = 0
    End With
    Err.Clear
End Sub



