VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "assembly_ord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'/////////////////////////////////////////////////////////////////////////////
'(c) 2000 Gemini Software Solutions Inc. All Rights Reserved.
'
'This work is subject to U.S. and international copyright laws and treaties.
'
'No part of this work may be used, practiced, performed, copied, distributed,
'revised, modified, translated, abridged, condensed, expanded, collected,
'compiled, linked, recast, transformed or adapted without the prior written
'consent of GSS Inc.
'
'Any use or exploitation of this work without authorization could subject the
'perpetrator to criminal and civil liability.
'
'/////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'Description:   This class object is responsible for maintaining assembly
'               orders.
'//////////////////////////////////////////////////////////////////////////////

Private m_varAssemOrdID             As Variant, _
        m_varItemID                 As Variant, _
        m_varOrderedQty             As Variant, _
        m_varAdjQty                 As Variant, _
        m_varCompletedQty           As Variant, _
        m_varReleasedQty            As Variant, _
        m_varRemainingBalanceQty    As Variant, _
        m_varUpdatedAssocID         As Variant, _
        m_varEnteredAssocID         As Variant, _
        m_varTypeCD                 As Variant, _
        m_varTypeDesc               As Variant
        
Private m_varStatusCd               As Variant, _
        m_varStatusDesc             As Variant, _
        m_varLocID                  As Variant, _
        m_varItemNbr                As Variant, _
        m_dtmClosedDt               As Variant, _
        m_dtmDueDt                  As Variant, _
        m_dtmOpenedDt               As Variant, _
        m_dtmEnteredDt              As Variant, _
        m_dtmBeginOpenedDt          As Variant, _
        m_dtmEndOpenedDt            As Variant, _
        m_dtmBeginDueDt             As Variant, _
        m_dtmEndDueDt               As Variant, _
        m_dtmBeginEnteredDt         As Variant, _
        m_dtmEndEnteredDt           As Variant, _
        m_dtmBeginClosedDt          As Variant, _
        m_dtmEndClosedDt            As Variant, _
        m_objOld_assembly_order     As assembly_ord, _
        m_rstRecordSet              As ADODB.RecordSet, _
        m_objError                  As Object, _
        m_objTransData              As Object, _
        m_objItem                   As Object
                                        ' Error Object
Public Property Set o_item(ByVal v_objItem As Object)
    Set m_objItem = v_objItem
End Property
Public Property Get o_item() As Object
    Set o_item = m_objItem
End Property
Public Property Set o_old_assembly_order(ByVal vdata As assembly_ord)
    Set m_objOld_assembly_order = vdata
End Property
Public Property Get o_old_assembly_order() As assembly_ord
    Set o_old_assembly_order = m_objOld_assembly_order
End Property
                                        ' Transaction Object
Public Property Set o_error(ByVal v_objError As Object)
    Set m_objError = v_objError
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
                                        ' Transaction Object
Public Property Set o_trans_data(ByVal v_objTransData As Object)
    Set m_objTransData = v_objTransData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTransData
End Property
                                        ' The assembly order ID
Public Property Let AssemOrdID(ByVal v_intAssemOrdID As Variant)
    m_varAssemOrdID = v_intAssemOrdID
End Property

Public Property Get AssemOrdID() As Variant
    AssemOrdID = m_varAssemOrdID
End Property
                                        ' The location ID
Public Property Let LocID(ByVal v_strLocID As Variant)
    m_varLocID = v_strLocID
End Property

Public Property Get LocID() As Variant
    LocID = m_varLocID
End Property
                                        ' The item ID
Public Property Let ItemID(ByVal v_dblItemID As Variant)
    m_varItemID = v_dblItemID
End Property

Public Property Get ItemID() As Variant
    ItemID = m_varItemID
End Property
                                        ' The order qty
Public Property Let OrderedQty(ByVal v_dblOrderedQty As Variant)
    m_varOrderedQty = v_dblOrderedQty
End Property

Public Property Get OrderedQty() As Variant
    OrderedQty = m_varOrderedQty
End Property
                                        ' The Quantity from Inventory Adjustment
Public Property Let AdjQty(ByVal v_dblAdjQty As Variant)
    m_varAdjQty = v_dblAdjQty
End Property

Public Property Get AdjQty() As Variant
    AdjQty = m_varAdjQty
End Property
                                        
                                        ' The completed qty
Public Property Let CompletedQty(ByVal v_dblCompletedQty As Variant)
    m_varCompletedQty = v_dblCompletedQty
End Property

Public Property Get CompletedQty() As Variant
    CompletedQty = m_varCompletedQty
End Property
                                        ' The released qty
Public Property Let ReleasedQty(ByVal v_dblReleasedQty As Variant)
    m_varReleasedQty = v_dblReleasedQty
End Property

Public Property Get RemainingBalanceQty() As Variant
    RemainingBalanceQty = m_varRemainingBalanceQty
End Property
Public Property Let RemainingBalanceQty(ByVal v_dblRemainingBalanceQty As Variant)
    m_varRemainingBalanceQty = v_dblRemainingBalanceQty
End Property

Public Property Get ReleasedQty() As Variant
    ReleasedQty = m_varReleasedQty
End Property
                                        ' The Updating Associate ID
Public Property Let UpdatedAssocID(ByVal v_dblUpdatedAssocID As Variant)
    m_varUpdatedAssocID = v_dblUpdatedAssocID
End Property

Public Property Get UpdatedAssocID() As Variant
    UpdatedAssocID = m_varUpdatedAssocID
End Property
                                        ' The Associate ID of the first entry.
Public Property Let EnteredAssocID(ByVal v_dblEnteredAssocID As Variant)
    m_varEnteredAssocID = v_dblEnteredAssocID
End Property

Public Property Get EnteredAssocID() As Variant
    EnteredAssocID = m_varEnteredAssocID
End Property
                                        ' The status code.
Public Property Let StatusCd(ByVal v_strStatusCd As Variant)
    m_varStatusCd = v_strStatusCd
End Property

Public Property Get StatusCd() As Variant
    StatusCd = m_varStatusCd
End Property
                                        ' The status description.
Public Property Let StatusDesc(ByVal v_strStatusDesc As Variant)
    m_varStatusDesc = v_strStatusDesc
End Property

Public Property Get StatusDesc() As Variant
    StatusDesc = m_varStatusDesc
End Property
                                        ' The item number
Public Property Let ItemNbr(ByVal v_strItemNbr As Variant)
    m_varItemNbr = v_strItemNbr
End Property

Public Property Get ItemNbr() As Variant
    ItemNbr = m_varItemNbr
End Property
                                        ' The close date.
Public Property Let ClosedDt(ByVal v_dtmClosedDt As Variant)
    m_dtmClosedDt = v_dtmClosedDt
End Property

Public Property Get ClosedDt() As Variant
    ClosedDt = m_dtmClosedDt
End Property
                                        ' The due date.
Public Property Let DueDt(ByVal v_dtmDueDt As Variant)
    m_dtmDueDt = v_dtmDueDt
End Property

Public Property Get DueDt() As Variant
    DueDt = m_dtmDueDt
End Property
                                        ' The opened date.
Public Property Let OpenedDt(ByVal v_dtmOpenedDt As Variant)
    m_dtmOpenedDt = v_dtmOpenedDt
End Property

Public Property Get OpenedDt() As Variant
    OpenedDt = m_dtmOpenedDt
End Property
                                        ' The begin opened date for date query
Public Property Let BeginOpenedDt(ByVal v_dtmBeginOpenedDt As Variant)
    m_dtmBeginOpenedDt = v_dtmBeginOpenedDt
End Property

Public Property Get BeginOpenedDt() As Variant
    BeginOpenedDt = m_dtmBeginOpenedDt
End Property
                                        ' The end opened date for date query
Public Property Let EndOpenedDt(ByVal v_dtmEndOpenedDt As Variant)
    m_dtmEndOpenedDt = v_dtmEndOpenedDt
End Property

Public Property Get EndOpenedDt() As Variant
    EndOpenedDt = m_dtmEndOpenedDt
End Property
                                        ' The begin due date for date query
Public Property Let BeginDueDt(ByVal v_dtmBeginDueDt As Variant)
    m_dtmBeginDueDt = v_dtmBeginDueDt
End Property

Public Property Get BeginDueDt() As Variant
    BeginDueDt = m_dtmBeginDueDt
End Property
                                        ' The end due date for date query
Public Property Let EndDueDt(ByVal v_dtmEndDueDt As Variant)
    m_dtmEndDueDt = v_dtmEndDueDt
End Property

Public Property Get EndDueDt() As Variant
    EndDueDt = m_dtmEndDueDt
End Property
                                        ' The begin entered date for date query
Public Property Let BeginEnteredDt(ByVal v_dtmBeginEnteredDt As Variant)
    m_dtmBeginEnteredDt = v_dtmBeginEnteredDt
End Property

Public Property Get BeginEnteredDt() As Variant
    BeginEnteredDt = m_dtmBeginEnteredDt
End Property
                                        ' The end entered date for date query
Public Property Let EndEnteredDt(ByVal v_dtmEndEnteredDt As Variant)
    m_dtmEndEnteredDt = v_dtmEndEnteredDt
End Property

Public Property Get EndEnteredDt() As Variant
    EndEnteredDt = m_dtmEndEnteredDt
End Property
                                        ' The begin close date for date query
Public Property Let BeginClosedDt(ByVal v_dtmBeginClosedDt As Variant)
    m_dtmBeginClosedDt = v_dtmBeginClosedDt
End Property

Public Property Get BeginClosedDt() As Variant
    BeginClosedDt = m_dtmBeginClosedDt
End Property
                                        ' The end close date for date query
Public Property Let EndClosedDt(ByVal v_dtmEndClosedDt As Variant)
    m_dtmEndClosedDt = v_dtmEndClosedDt
End Property

Public Property Get EndClosedDt() As Variant
    EndClosedDt = m_dtmEndClosedDt
End Property
                                        ' The entered date.
Public Property Let EnteredDt(ByVal v_dtmEnteredDt As Variant)
    m_dtmEnteredDt = v_dtmEnteredDt
End Property

Public Property Get EnteredDt() As Variant
    EnteredDt = m_dtmEnteredDt
End Property
                                        ' Returning recordset
Public Property Set RecordSet(ByVal v_rstRecordSet As ADODB.RecordSet)
    Set m_rstRecordSet = v_rstRecordSet
End Property

Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = m_rstRecordSet
End Property
                                        ' The item type code.
Public Property Let TypeCd(ByVal v_strTypeCd As Variant)
    m_varTypeCD = v_strTypeCd
End Property

Public Property Get TypeCd() As Variant
    TypeCd = m_varTypeCD
End Property
                                        ' The item type description.
Public Property Let TypeDesc(ByVal v_strTypeDesc As Variant)
    m_varTypeDesc = v_strTypeDesc
End Property

Public Property Get TypeDesc() As Variant
    TypeDesc = m_varTypeDesc
End Property

Public Sub GetAssemblyOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm           As String
    Let strRoutine_nm = "assembly_ord.cls.GetAssemblyOrder"
    
    Dim strSql                  As String, _
        rstRecordSet            As ADODB.RecordSet
    
    If IsNull(Me.AssemOrdID) = True Then
        Exit Sub
    End If
    
    strSql = "SELECT a.assem_ord_id, a.inv_loc_id, " & _
             "       a.status_cd, a.closed_dt, " & _
             "       a.due_dt, a.item_id, " & _
             "       a.opened_dt, " & _
             "       a.entered_dt, a.ordered_qty, " & _
             "       a.released_qty, a.completed_qty, " & _
             "       b.type_cd " & _
             "FROM assembly_order a, item b " & _
             "WHERE a.item_id = b.item_id AND " & _
             fWhereVariable("assem_ord_id", AssemOrdID, "N", "=")
    
    Call Me.o_trans_data.OpenRecordset(rstRecordSet, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        rstRecordSet.Close
        Set rstRecordSet = Nothing
        Exit Sub
    End If
    
    If rstRecordSet.EOF = True Then
        Let Me.AssemOrdID = Null
    Else
        Me.AssemOrdID = rstRecordSet("assem_ord_id")
        Me.LocID = rstRecordSet("inv_loc_id")
        Me.StatusCd = rstRecordSet("status_cd")
        Me.StatusDesc = fGetDesc("assembly_order_status_cd", rstRecordSet("status_cd"), o_error, o_trans_data)
        Me.ClosedDt = rstRecordSet("closed_dt")
        Me.DueDt = rstRecordSet("due_dt")
        Me.ItemID = rstRecordSet("item_id")
        Me.ItemNbr = fGetItemNbr(rstRecordSet("item_id"), o_error, o_trans_data)
        Me.TypeCd = rstRecordSet("type_cd")
        Me.TypeDesc = fGetDesc("item_type_cd", rstRecordSet("type_cd"), o_error, o_trans_data)
        Me.OpenedDt = rstRecordSet("opened_dt")
        Me.EnteredDt = Format(rstRecordSet("entered_dt"), "mm/dd/yyyy")
        Me.OrderedQty = rstRecordSet("ordered_qty")
        Me.ReleasedQty = rstRecordSet("released_qty")
        If IsNull(rstRecordSet("completed_qty")) = True Then
            Me.CompletedQty = 0
        Else
            Me.CompletedQty = rstRecordSet("completed_qty")
        End If
        Me.RemainingBalanceQty = Me.OrderedQty - (Me.ReleasedQty + Me.CompletedQty)
        If Me.RemainingBalanceQty < 0 Then
            Me.RemainingBalanceQty = 0
        End If
    End If
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub GetAssemblyOrders()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    
    strRoutine_nm = "assembly_ord.cls.GetAssemblyOrders"
    
    strSql = "select a.assem_ord_id, a.inv_loc_id, " & _
             "       a.status_cd, a.closed_dt, " & _
             "       a.due_dt, a.item_id, " & _
             "       a.opened_dt, a.entered_dt, " & _
             "       a.ordered_qty, a.released_qty, " & _
             "       a.completed_qty, b.type_cd " & _
             "FROM assembly_order a, item b " & _
             "WHERE a.item_id = b.item_id "
    
    If Len(Me.AssemOrdID) > 0 Then
        strSql = strSql & " and assem_ord_id = " & Me.AssemOrdID
    End If
    
    If Len(Me.ItemNbr) > 0 Then
        strSql = strSql & " and a.item_id = " & modUtilites.fGetItemId(Me.ItemNbr, o_error, o_trans_data)
    End If
    
    If Len(Me.LocID) > 0 Then
        strSql = strSql & " and inv_loc_id = '" & Me.LocID & "' "
    End If
    
    If Len(Me.StatusCd) > 0 Then
        strSql = strSql & " and status_cd = '" & Me.StatusCd & "' "
    End If
    
    If IsDate(Me.BeginOpenedDt) And IsDate(Me.EndOpenedDt) Then
        strSql = strSql & " and " & fWhereVariable("opened_dt", Me.BeginOpenedDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("opened_dt", Me.EndOpenedDt, "D", "<=")
    End If
    
    If IsDate(Me.BeginDueDt) And IsDate(Me.EndDueDt) Then
        strSql = strSql & " and " & fWhereVariable("due_dt", Me.BeginDueDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("due_dt", Me.EndDueDt, "D", "<=")
    End If
    
    If IsDate(Me.BeginEnteredDt) And IsDate(Me.EndEnteredDt) Then
        strSql = strSql & " and " & fWhereVariable("entered_dt", Me.BeginEnteredDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("entered_dt", Me.EndEnteredDt, "D", "<=")
    End If
    
    If IsDate(Me.BeginClosedDt) And IsDate(Me.EndClosedDt) Then
        strSql = strSql & " and " & fWhereVariable("closed_dt", Me.BeginClosedDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("closed_dt", Me.EndClosedDt, "D", "<=")
    End If
    
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    Set Me.RecordSet = objRecordset

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub GetPastDueAssemblyOrders()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    
    strRoutine_nm = "assembly_ord.cls.GetPastDueAssemblyOrders"
    
    strSql = " SELECT assembly_order.assem_ord_id, "
    strSql = strSql & " assembly_order.status_cd, item.item_nbr, "
    strSql = strSql & " assembly_order.ordered_qty, assembly_order.released_qty,"
    strSql = strSql & " assembly_order.completed_qty, assembly_order.due_dt,"
    strSql = strSql & " Item.lead_days"
    strSql = strSql & " FROM assembly_order LEFT OUTER JOIN"
    strSql = strSql & " item ON assembly_order.item_id = item.item_id"
    strSql = strSql & " WHERE (assembly_order.status_cd = N'OP') OR"
    strSql = strSql & " (assembly_order.status_cd = N'PL')"
        
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    Set Me.RecordSet = objRecordset

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub GetNonPastDueAssemblyOrders()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    
    strRoutine_nm = "assembly_ord.cls.GetMatReqs"
    
    strSql = " SELECT assembly_order.assem_ord_id,"
        strSql = strSql & " assembly_order.status_cd, item.item_nbr,"
        strSql = strSql & " assembly_order.ordered_qty, assembly_order.released_qty,"
        strSql = strSql & " assembly_order.completed_qty, assembly_order.due_dt,"
        strSql = strSql & " Item.lead_days"
    strSql = strSql & " FROM assembly_order LEFT OUTER JOIN"
        strSql = strSql & " item ON assembly_order.item_id = item.item_id"
    strSql = strSql & " WHERE (assembly_order.status_cd = N'OP')"
      
    Call Me.o_trans_data.OpenRecordset(objRecordset, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    Set Me.RecordSet = objRecordset

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub AddAssemblyOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "Assembly_order.cls.Add"
    
    Call ValidateData("Add")
    Call IssueInsertData
    Call AddAssemblyComponents
        
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub ChangeAssemblyOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    strRoutine_nm = "Assembly_order.cls.Change"
    
    Dim strStatusCd As String
    Dim varCompRequiredDt As Variant
    
    Call GetOldAssemblyOrder
    Call ValidateData("Change")
    Call IssueUpdateData
    Call Me.GetAssemblyOrder
    varCompRequiredDt = ComputeCompRequiredDt(Me.ItemID, Me.DueDt)
    
    
                    'If the ordered quantity has changed,
                    'update the on order inventory, assembly component
                    'required quantity and set the status of the
                    'assembly order.
            'note: if required date is changed, assembly components
            'required date needs to be updated also
    'If Me.OrderedQty <> Me.o_old_assembly_order.OrderedQty Then
        
        strStatusCd = Me.StatusCd
        
        Call changeOnOrderInventory( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.AssemOrdID _
                    , Me.StatusCd _
                    , Me.OrderedQty _
                    , Me.CompletedQty _
                    , Me.LocID _
                    , Me.ItemID _
                    )
    
        Call updateAssemCompRequiredQty( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.AssemOrdID _
                    , Me.OrderedQty _
                    , varCompRequiredDt _
                    )
                    
        Call SetStatus( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.AssemOrdID _
                    )
        
        Call Me.GetAssemblyOrder
        
                    'if the assembly order was complete, but
                    'now it is in an Open status, open the
                    'assembly components.
        If strStatusCd = "CP" And _
           (Me.StatusCd = "OP" Or _
            Me.StatusCd = "RD" Or _
            Me.StatusCd = "SC") Then
            Call updateAssemCompStatus( _
                      o_error _
                    , o_trans_data _
                    , Me.AssemOrdID _
                    , "OP" _
                    )
        End If
                            
                    'if the assembly order was open, but
                    'now it is in an Complete status, Set the
                    'assembly components to complete.
        If Me.StatusCd = "CP" And _
           (strStatusCd = "OP" Or _
            strStatusCd = "RD" Or _
            strStatusCd = "SC") Then
            Call updateAssemCompStatus( _
                      o_error _
                    , o_trans_data _
                    , Me.AssemOrdID _
                    , "CP" _
                    )
        End If
    'End If
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub DeleteAssemblyOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    Let strRoutine_nm = "Assembly_order.cls.Delete"
    
    Call GetOldAssemblyOrder
                                        'only orders in a planned status can be
                                        'deleted
    If Me.o_old_assembly_order.StatusCd <> "PL" Then
        With Me.o_error
            .p_type_cd = "E"
            .p_err_cd = "0310"
            .p_message_id = 0
            .p_routine_nm = strRoutine_nm
            .p_desc = "Only assembly orders in a planned status can be deleted"
            .p_sqlstate_cd = ""
        End With
    End If
    
    Call IssueDeleteItem
    Call IssueDeleteData
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub



Private Function ValidateData(ByVal v_strFunction_cd As String) As Boolean

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm As String
    Let strRoutine_nm = "Assembly_order.cls.ValidateData"
    
    Dim strSql              As String, _
        rstInventoryLoc     As RecordSet
    
    If v_strFunction_cd = "Add" Then
        Set Me.o_item = CreateObject("converge_item.item")
        Set Me.o_item.o_error = Me.o_error
        Set Me.o_item.o_trans_data = Me.o_trans_data
        If Len(Me.ItemID) > 0 Then
            Me.o_item.p_item_id = Me.ItemID
            Me.o_item.p_item_nbr = Null
        Else
            Me.o_item.p_item_id = Null
            Me.o_item.p_item_nbr = Me.ItemNbr
        End If
        Call Me.o_item.getItem
        Me.ItemID = Me.o_item.p_item_id
        Me.ItemNbr = Me.o_item.p_item_nbr
        
        If fValuePresent(Me.LocID) = False Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0100"
            Let Me.o_error.p_message_id = 2117
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Inventory Location ID requires a value to be passed"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
        
        strSql = "SELECT count(*) as record_count" & _
                 " FROM inventory_loc " & _
                 "WHERE " & fWhereVariable("inv_loc_id", Me.LocID, "S", "=")
        Call Me.o_trans_data.OpenRecordset(rstInventoryLoc, strSql, Me.o_error)
        If rstInventoryLoc("record_count") = 0 Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0200"
            Let Me.o_error.p_message_id = 2123
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Inventory Location ID does not exist"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
                
    If v_strFunction_cd = "Add" Or v_strFunction_cd = "Change" Then
        If IsDate(Me.DueDt) = False Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0250"
            Let Me.o_error.p_message_id = 2119
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Due Date requires a value to be passed"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
    
    If v_strFunction_cd = "Add" Then
        If fValuePresent(Me.ItemID) = False Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0300"
            Let Me.o_error.p_message_id = 2120
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Item Number requires a value to be passed"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
        If Me.o_item.p_type_cd = "AS" Then
        Else
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0310"
            Let Me.o_error.p_message_id = 3302
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Item Must be an Assembled Item"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
    
    If v_strFunction_cd = "Add" Or v_strFunction_cd = "Change" Then
        If fValuePresent(Me.OrderedQty) = False Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0450"
            Let Me.o_error.p_message_id = 2122
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Ordered Quantity requires a value to be passed"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
        
    ValidateData = True
    Exit Function
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


Private Sub AddAssemblyComponents()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    Let strRoutine_nm = "Assembly_order.cls.AddAssemblyComponents"
    
    Dim objAssemblyComp As assembly_comp
    Dim varCompRequiredDt As Variant
    
    varCompRequiredDt = ComputeCompRequiredDt(Me.ItemID, Me.DueDt)
    
    
    Set objAssemblyComp = New assembly_comp
    Set objAssemblyComp.o_error = Me.o_error
    Set objAssemblyComp.o_trans_data = Me.o_trans_data
    objAssemblyComp.AssemOrdID = Me.AssemOrdID
    objAssemblyComp.ItemNbr = Me.ItemNbr
    objAssemblyComp.RequiredDt = varCompRequiredDt
    objAssemblyComp.RequiredQty = Me.OrderedQty
    Call objAssemblyComp.AddComponents
    
    Set objAssemblyComp = Nothing

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub openOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_order.cls.OpenOrder"
    
    StatusCd = "OP"
    Call ChangeStatus
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub ChangeStatus()

    On Error GoTo errHandler

    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_order.cls.changeStatus"
    
    Dim dblUpdatedqty As Double, _
        strSql As String
                    'The user can only manually Open, Close, Cancel or Hold the
                    'order.  All other statuses are system generated by the set
                    'status routine.
    If Me.StatusCd = "OP" Or _
       Me.StatusCd = "CL" Or _
       Me.StatusCd = "OH" Or _
       Me.StatusCd = "CA" Then
    Else
        Me.o_error.p_type_cd = "F"
        Me.o_error.p_err_cd = "1300"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 1
        Exit Sub
    End If

    Call GetOldAssemblyOrder
                    
                    'If the order is currently in a planned status, it can only
                    'be Opened or Cancelled.
    If Me.o_old_assembly_order.StatusCd = "PL" Then
        If Me.StatusCd = "OP" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1410"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1
            Exit Sub
        End If
    End If

                    'If the order is currently in a open status
                    'it can be closed cancelled or held.  Note that Ready
                    'scheduled Complete are also considered open statuses
    If Me.o_old_assembly_order.StatusCd = "OP" Or _
       Me.o_old_assembly_order.StatusCd = "RD" Or _
       Me.o_old_assembly_order.StatusCd = "SC" Or _
       Me.o_old_assembly_order.StatusCd = "CP" Then
        If Me.StatusCd = "OP" Or _
            Me.StatusCd = "CL" Or _
            Me.StatusCd = "CA" Or _
            Me.StatusCd = "OH" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1420"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1
            Exit Sub
        End If
    End If
                    'if order is already closed, it can only be closed again
                    'or opened
    If Me.o_old_assembly_order.StatusCd = "CL" Then
        If Me.StatusCd = "OP" Or Me.StatusCd = "CL" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1440"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1
            Exit Sub
        End If
    End If

                    'if order has been held, it can only be
                    'opened, cancelled or closed
    If Me.o_old_assembly_order.StatusCd = "OH" Then
        If Me.StatusCd = "OP" Or _
            Me.StatusCd = "CL" Or _
            Me.StatusCd = "CA" Then
        Else
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "1480"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1
            Exit Sub
        End If
    End If

                   'if order has been cancelled, the status can never be changed again.
    If Me.o_old_assembly_order.StatusCd = "CA" Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1450"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 1
        Exit Sub
    End If

                    'order can't be closed until it has at least 1 item completed
    If Me.StatusCd = "CL" And Me.o_old_assembly_order.CompletedQty < 1 Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1420"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 1
        Exit Sub
    End If
                    'order can only be cancelled if no items have been completed
    If Me.StatusCd = "CA" And Me.o_old_assembly_order.CompletedQty > 0 Then
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "1420"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 1
        Exit Sub
    End If
    
    Call IssueUpdateStatus( _
              Me.o_error _
            , Me.o_trans_data _
            , Me.StatusCd _
            , Me.AssemOrdID _
            )
                                        
    Call changeOnOrderInventory( _
              Me.o_error _
            , Me.o_trans_data _
            , Me.AssemOrdID _
            , Me.StatusCd _
            , Me.o_old_assembly_order.OrderedQty _
            , Me.o_old_assembly_order.CompletedQty _
            , Me.o_old_assembly_order.LocID _
            , Me.o_old_assembly_order.ItemID _
            )
    
    Call updateAssemCompStatus( _
              Me.o_error _
            , Me.o_trans_data _
            , Me.AssemOrdID _
            , Me.StatusCd _
            )
    
    If Me.StatusCd = "OP" Then
        Call SetStatus( _
              Me.o_error _
            , Me.o_trans_data _
            , Me.AssemOrdID _
            )
    
        Call Me.GetAssemblyOrder
        
                    'if the assembly order was open, but
                    'now it is in an Complete status, Set the
                    'assembly components to complete.
        If Me.StatusCd = "CP" Then
            Call updateAssemCompStatus( _
                      o_error _
                    , o_trans_data _
                    , Me.AssemOrdID _
                    , "CP" _
                    )
        End If
    
    End If
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub UpdateQuantities( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblAssemOrdId As Variant _
            )

    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_ord.cls.UpdateQuantities"

    Dim dblUpdateQty    As Double, _
        strSql          As String, _
        objRecordset    As ADODB.RecordSet, _
        varCompletedQty As Variant, _
        varReleasedQty   As Variant, _
        strStatusCd     As String
        
                                'get current assembly order information
    strSql = "Select sum(ordered_qty)as sum_ordered_qty " & _
             ", sum(completed_qty) as sum_completed_qty " & _
             " from prod_run " & _
             " where " & _
             "  " & fWhereVariable("assem_ord_id", v_dblAssemOrdId, "N", "=") & _
             "  AND " & fWhereVariable("status_cd", "CA", "S", "<>")
             
    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        varReleasedQty = 0
        varCompletedQty = 0
    Else
        varReleasedQty = objRecordset("sum_ordered_qty") - objRecordset("sum_completed_qty")
        varCompletedQty = objRecordset("sum_completed_qty")
    End If
    If varReleasedQty < 1 Then
        varReleasedQty = 0
    End If
    If Len(varReleasedQty) = 0 Then
        varReleasedQty = 0
    End If
    If Len(varCompletedQty) = 0 Then
        varCompletedQty = 0
    End If
    
    Call objRecordset.Close
    Set objRecordset = Nothing
    
    strSql = "update assembly_order set "
    strSql = strSql & " " & fUpdateVariable("completed_qty", varCompletedQty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("released_qty", varReleasedQty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "  " & fWhereVariable("assem_ord_id", v_dblAssemOrdId, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdateQty, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
       Exit Sub
    End If

    If dblUpdateQty = 0 Then
        With r_objError
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_routine_nm = strRoutine_nm
            .p_message_id = 1005
        End With
    End If
    
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.AssemOrdID = v_dblAssemOrdId
    Me.GetAssemblyOrder
    strStatusCd = Me.StatusCd
    
    Call changeOnOrderInventory( _
                  r_objError _
                , r_objTransData _
                , Me.AssemOrdID _
                , Me.StatusCd _
                , Me.OrderedQty _
                , Me.CompletedQty _
                , Me.LocID _
                , Me.ItemID _
                )
                
    Call SetStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblAssemOrdId _
                )
    
    Me.GetAssemblyOrder
                    'if the assembly order was complete, but
                    'now it is in an Open status, open the
                    'assembly components.
    If strStatusCd = "CP" And _
       (Me.StatusCd = "OP" Or _
        Me.StatusCd = "RD" Or _
        Me.StatusCd = "SC") Then
        Call updateAssemCompStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblAssemOrdId _
                , "OP" _
                )
    End If
                            
                    'if the assembly order was open, but
                    'now it is in an Complete status, Set the
                    'assembly components to complete.
    If Me.StatusCd = "CP" And _
       (strStatusCd = "OP" Or _
        strStatusCd = "RD" Or _
        strStatusCd = "SC") Then
        Call updateAssemCompStatus( _
                  r_objError _
                , r_objTransData _
                , v_dblAssemOrdId _
                , "CP" _
                )
    End If
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Public Sub SetStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblAssemOrdId As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.setstatus"
    
    Dim strStatusCd As String, _
        strSql As String, _
        objRecordset As ADODB.RecordSet
        
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.AssemOrdID = v_dblAssemOrdId
    Call Me.GetAssemblyOrder
    
    If Me.StatusCd = "CA" Or _
       Me.StatusCd = "PL" Or _
       Me.StatusCd = "CL" Or _
       Me.StatusCd = "OH" Then
        Exit Sub
    End If
        
    strStatusCd = "OP"
    
                'check all assembly components. if any assembly components
                'do not have a ready status, the assembly order will stay in an
                'open status
    strSql = "Select line_nbr " & _
             " from assembly_comp a, mat_req b " & _
             " where " & _
             "     " & fWhereVariable("a.assem_ord_id", Me.AssemOrdID, "N", "=") & _
             " and  a.assem_ord_id = b.assem_ord_id " & _
             " and  a.line_nbr = b.ac_line_nbr " & _
             " and  b.status_cd <> 'RD' " & _
             " and  b.status_cd <> 'CL' " & _
             " and  b.status_cd <> 'CA' "

    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
                'if the query above returns no records,
                'the the status is ready
    If objRecordset.EOF = True Then
        strStatusCd = "RD"
    End If
    
                'if the assembly order is ready
                'the the status is ready
    If strStatusCd = "RD" Then
        If Me.RemainingBalanceQty < 1 Then
            strStatusCd = "SC"
        End If
    End If
    
    If Me.CompletedQty >= Me.OrderedQty Then
        strStatusCd = "CP"
    End If
    
    If Me.StatusCd <> strStatusCd Then
        Call IssueUpdateStatus( _
                  r_objError _
                , r_objTransData _
                , strStatusCd _
                , v_dblAssemOrdId _
                )
        
    End If
                    
                    'do not call update component items from here
                    'this is called from the material requirement set
                    'status routine which is called from assembly_comp.changestatus.
                    'calling it from here will cause problems or maybe
                    'an infinite loop!!!!!
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Public Sub generatePlan()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.generatePlan"

    Dim objMatReq As mat_req
    Dim objRecordset As ADODB.RecordSet
    Dim dblPlannedQty As Double
    Dim dblUnplannedQty As Double
    Set objMatReq = New mat_req
    Set objMatReq.o_error = Me.o_error
    Set objMatReq.o_trans_data = Me.o_trans_data
    objMatReq.assem_ord_id = Me.AssemOrdID
    Call objMatReq.GetMatReqs
    Set objRecordset = objMatReq.RecordSet
    Set objMatReq = Nothing
    
    Call objRecordset.MoveFirst
    
    Do Until objRecordset.EOF = True Or _
        Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F"
        dblPlannedQty = objRecordset("back_ordered_qty") + objRecordset("ready_qty") + objRecordset("consumed_qty")
        dblUnplannedQty = objRecordset("required_qty") - dblPlannedQty
        If dblUnplannedQty > 0 Then
            Set objMatReq = New mat_req
            Set objMatReq.o_error = Me.o_error
            Set objMatReq.o_trans_data = Me.o_trans_data
            objMatReq.mat_req_id = objRecordset("mat_req_id")
            Call objMatReq.generatePlan
        End If
        Call objRecordset.MoveNext
    Loop

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub GetOldAssemblyOrder()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.GetOldAssemblyOrder"

    If Me.o_old_assembly_order Is Nothing Then
        Set Me.o_old_assembly_order = New assembly_ord
        Set Me.o_old_assembly_order.o_trans_data = Me.o_trans_data
        Set Me.o_old_assembly_order.o_error = Me.o_error
        Me.o_old_assembly_order.AssemOrdID = Me.AssemOrdID
        Me.o_old_assembly_order.GetAssemblyOrder
        If IsNull(Me.o_old_assembly_order.AssemOrdID) = True Then
            Me.o_error.p_type_cd = "E"
            Me.o_error.p_err_cd = "0800"
            Me.o_error.p_routine_nm = strRoutine_nm
            Me.o_error.p_message_id = 1
            Exit Sub
        End If
    End If

    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub IssueUpdateStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_varStatus_cd As Variant _
            , ByVal v_dblAssemOrdId As Variant _
            )
    
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.IssueUpdateStatus"
    
    Dim strSql As String
    Dim dblUpdatedqty As Double
    
    strSql = "update assembly_order set "
    strSql = strSql & fUpdateVariable("status_cd", v_varStatus_cd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", r_objTransData.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("assem_ord_id", v_dblAssemOrdId, "N", "=")
    
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
    
    If dblUpdatedqty = 0 Then
        With r_objError
            .type_cd = "F"
            .err_cd = "0100"
            .routine_nm = strRoutine_nm
            .message_id = 1005
        End With
        Exit Sub
    End If

    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub changeOnOrderInventory( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblAssemOrdId As Double _
            , ByVal v_strStatusCd As String _
            , ByVal v_dblOrderedQty As Double _
            , ByVal v_dblCompletedQty As Double _
            , ByVal v_strInvLocId As String _
            , ByVal v_dblItemID As Double _
            )
            
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.changeOnOrderInventory"
    
    Dim objOnOrdInvAdj As Object
    
    If v_strStatusCd = "PL" Then
        Exit Sub
    End If
    
    Set objOnOrdInvAdj = CreateObject("converge_inv.on_ord_inv_adj")
    With objOnOrdInvAdj
        Set .o_error = r_objError
        Set .o_trans_data = r_objTransData
        .p_po_id = Null
        .p_po_line_nbr = Null
        .p_assem_ord_id = v_dblAssemOrdId
        If v_strStatusCd = "CL" Or _
           v_strStatusCd = "CA" Then
            .p_on_ord_qty = 0
        Else
            .p_on_ord_qty = v_dblOrderedQty - v_dblCompletedQty
            If .p_on_ord_qty < 1 Then
                .p_on_ord_qty = 0
            End If
        End If
        .p_inv_loc_id = v_strInvLocId
        .p_item_id = v_dblItemID
        .Change
    End With
    Set objOnOrdInvAdj = Nothing
            
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub updateAssemCompStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblAssemOrdId As Double _
            , ByVal v_strStatusCd As String _
            )
            
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.updateAssemCompStatus"
    
                                        ' create assembly_comp object
    Dim objAssemblyComp As assembly_comp
    Dim objRecordset As ADODB.RecordSet
    Dim strSql As String
    Set objAssemblyComp = New assembly_comp
    
    strSql = "SELECT a.inv_loc_id, b.line_nbr, " & _
             "       b.status_cd, b.item_id, " & _
             "       b.required_qty, b.required_dt, " & _
             "       b.uom_cd " & _
             "FROM assembly_order a, assembly_comp b " & _
             "WHERE " & _
             " a.assem_ord_id = b.assem_ord_id AND " & _
             fWhereVariable("b.assem_ord_id", v_dblAssemOrdId, "N", "=")

    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Do Until objRecordset.EOF = True Or _
             r_objError.p_type_cd = "E" Or _
             r_objError.p_type_cd = "F"
             
        Set objAssemblyComp.o_error = r_objError
        Set objAssemblyComp.o_trans_data = r_objTransData
        objAssemblyComp.AssemOrdID = v_dblAssemOrdId
        objAssemblyComp.LineNbr = objRecordset("line_nbr")
        objAssemblyComp.StatusCd = v_strStatusCd
        Call objAssemblyComp.ChangeStatus
        objRecordset.MoveNext
    Loop
    
    objRecordset.Close
    Set objRecordset = Nothing
    Set objAssemblyComp = Nothing
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub updateAssemCompRequiredQty( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblAssemOrdId As Double _
            , ByVal v_dblOrderedQty As Double _
            , ByVal v_varRequiredDt As Variant _
            )
    On Error GoTo errHandler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
        
    Dim strRoutine_nm   As String
    strRoutine_nm = "assembly_ord.cls.updateAssemRequired"
    
                                        ' create assembly_comp object
    Dim objAssemblyComp As assembly_comp
    Dim objRecordset As ADODB.RecordSet
    Dim strSql As String
    Set objAssemblyComp = New assembly_comp
    
    strSql = "SELECT line_nbr, per_item_qty " & _
             "FROM assembly_comp b " & _
             "WHERE " & _
             fWhereVariable("b.assem_ord_id", v_dblAssemOrdId, "N", "=")

    Call r_objTransData.OpenRecordset(objRecordset, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    Do Until objRecordset.EOF = True Or _
             r_objError.p_type_cd = "E" Or _
             r_objError.p_type_cd = "F"
             
        Set objAssemblyComp.o_error = r_objError
        Set objAssemblyComp.o_trans_data = r_objTransData
        objAssemblyComp.AssemOrdID = v_dblAssemOrdId
        objAssemblyComp.LineNbr = objRecordset("line_nbr")
        objAssemblyComp.RequiredQty = v_dblOrderedQty * CDbl(objRecordset("per_item_qty"))
        objAssemblyComp.RequiredDt = v_varRequiredDt
        Call objAssemblyComp.Change
        objRecordset.MoveNext
    Loop
    
    objRecordset.Close
    Set objRecordset = Nothing
    Set objAssemblyComp = Nothing
    
    Exit Sub
errHandler:
    With r_objError
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub IssueInsertData()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    strRoutine_nm = "Assembly_order.cls.IssueInsertData"
    
    Dim strSql              As String, _
        dblUpdatedqty       As Double
    
    AssemOrdID = fGetId("assem_ord_id", Me.o_error, Me.o_trans_data)
                                        ' Set Item ID
    ItemID = fGetItemId(ItemNbr, Me.o_error, Me.o_trans_data)
                                        ' Build Insert Statement
    strSql = "insert into Assembly_order ( "
    strSql = strSql & "  assem_ord_id"
    strSql = strSql & ", inv_loc_id"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", closed_dt"
    strSql = strSql & ", due_dt"
    strSql = strSql & ", item_id"
    strSql = strSql & ", opened_dt"
    strSql = strSql & ", entered_dt"
    strSql = strSql & ", ordered_qty"
    strSql = strSql & ", completed_qty"
    strSql = strSql & ", released_qty"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(AssemOrdID, "N")
    strSql = strSql & "," & fInsertVariable(LocID, "S")
    strSql = strSql & "," & fInsertVariable("PL", "S")
    strSql = strSql & "," & fInsertVariable(ClosedDt, "D")
    strSql = strSql & "," & fInsertVariable(DueDt, "D")
    strSql = strSql & "," & fInsertVariable(ItemID, "N")
    strSql = strSql & "," & fInsertVariable(OpenedDt, "D")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(OrderedQty, "N")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, Me.o_error)
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub IssueUpdateData()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    Let strRoutine_nm = "Assembly_order.cls.IssueUpdateData"
    
    Dim strSql              As String, _
        dblUpdatedqty       As Double
    
    Let strSql = "update Assembly_order set "
    Let strSql = strSql & " " & fUpdateVariable("due_dt", Me.DueDt, "D", "=")
    Let strSql = strSql & "," & fUpdateVariable("ordered_qty", Me.OrderedQty, "N", "=")
    Let strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Me.o_trans_data.p_assoc_id, "N", "=")
    Let strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    Let strSql = strSql & " where"
    Let strSql = strSql & " " & fWhereVariable("assem_ord_id", Me.AssemOrdID, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, Me.o_error)
    
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Sub IssueDeleteData()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdateQty        As Double
        
    Let strRoutine_nm = "Assembly_order.cls.IssueDeleteData"

    Let strSql = "delete from assembly_order "
    Let strSql = strSql & "where "
    Let strSql = strSql & " " & fWhereVariable("assem_ord_id", Me.AssemOrdID, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, Me.o_error)
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


Private Sub IssueDeleteItem()

    On Error GoTo errHandler
    
    If Me.o_error.type_cd = "E" Or Me.o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdateQty        As Double
        
    Let strRoutine_nm = "Assembly_order.cls.IssueDeleteItem"

    Let strSql = "delete from assembly_comp "
    Let strSql = strSql & "where "
    Let strSql = strSql & " " & fWhereVariable("assem_ord_id", AssemOrdID, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, Me.o_error)
    Exit Sub
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub

Private Function ComputeCompRequiredDt(ByVal v_varItemId As Variant, ByVal v_varDueDt) As Variant
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        ComputeCompRequiredDt = Null
        Exit Function
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dtmRequiredDt       As Variant, _
        dblLeadDays         As Double, _
        rstItem             As ADODB.RecordSet
        
    strRoutine_nm = "assembly_ord.cls.ComputeCompRequiredDt"
                                        ' Get lead days for parent part
                                        ' supplier lead days serves as
                                        ' assembly lead days for assembly
                                        ' items
    strSql = "SELECT supplier_lead_days " & _
             "FROM item " & _
             "WHERE " & fWhereVariable("item_id", v_varItemId, "N", "=")
    Call Me.o_trans_data.OpenRecordset(rstItem, strSql, Me.o_error)
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        rstItem.Close
        Set rstItem = Nothing
        ComputeCompRequiredDt = CDate(v_varDueDt)
        Exit Function
    End If
    
    If rstItem.EOF Then
        ComputeCompRequiredDt = CDate(v_varDueDt)
        Exit Function
    End If
    
    dblLeadDays = CDbl(rstItem("supplier_lead_days"))
    rstItem.Close
    Set rstItem = Nothing
                                        ' Compute date
    ComputeCompRequiredDt = CVar(DateAdd("d", (dblLeadDays * (-1)), CDate(v_varDueDt)))
    
    Exit Function
errHandler:
    With Me.o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description & " Item Id:" & v_varItemId & " lead days:" & dblLeadDays & " required dt:" & v_varDueDt & " rstItem.eof " & rstItem.EOF
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Function


