VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "prod_run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'Description:   This class object is responsible for maintaining production
'               runs.
'//////////////////////////////////////////////////////////////////////////////

Private m_intProductionRunID        As Variant, _
        m_intAssemOrdID             As Variant, _
        m_dblItemID                 As Variant, _
        m_dblOrderedQty             As Variant, _
        m_dblCompletedQty           As Variant, _
        m_dblReleasedQty            As Variant, _
        m_dblUpdatedAssocID         As Variant, _
        m_dblEnteredAssocID         As Variant, _
        m_strStatusCd               As Variant, _
        m_strStatusDesc             As Variant, _
        m_strItemNbr                As Variant, _
        m_dtmPlannedStartDt         As Variant, _
        m_dtmDueDt                  As Variant, _
        m_dtmBeginPlannedStartDt    As Variant, _
        m_dtmEndPlannedStartDt      As Variant, _
        m_dtmBeginDueDt             As Variant, _
        m_dtmEndDueDt               As Variant, _
        m_rstRecordSet              As ADODB.RecordSet, _
        m_objError                  As Object, _
        m_objTransData              As Object, _
        m_objOldProductionRun       As prod_run
                                        ' Error Object
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
                                        
Public Property Set o_old_production_run(ByVal v_obj As prod_run)
    Set m_objOldProductionRun = v_obj
End Property
Public Property Get o_old_production_run() As prod_run
    Set o_old_production_run = m_objOldProductionRun
End Property
                                        
                                        ' The production run ID
Public Property Let ProductionRunID(ByVal v_intProductionRunID As Variant)
    m_intProductionRunID = v_intProductionRunID
End Property

Public Property Get ProductionRunID() As Variant
    ProductionRunID = m_intProductionRunID
End Property
                                        ' The assembly order ID
Public Property Let AssemOrdID(ByVal v_intAssemOrdID As Variant)
    m_intAssemOrdID = v_intAssemOrdID
End Property

Public Property Get AssemOrdID() As Variant
    AssemOrdID = m_intAssemOrdID
End Property
                                        ' The item ID
Public Property Let ItemID(ByVal v_dblItemID As Variant)
    m_dblItemID = v_dblItemID
End Property

Public Property Get ItemID() As Variant
    ItemID = m_dblItemID
End Property
                                        ' The order qty
Public Property Let OrderedQty(ByVal v_dblOrderedQty As Variant)
    m_dblOrderedQty = v_dblOrderedQty
End Property

Public Property Get OrderedQty() As Variant
    OrderedQty = m_dblOrderedQty
End Property
                                        ' The completed qty
Public Property Let CompletedQty(ByVal v_dblCompletedQty As Variant)
    m_dblCompletedQty = v_dblCompletedQty
End Property

Public Property Get CompletedQty() As Variant
    CompletedQty = m_dblCompletedQty
End Property
                                        ' The Updating Associate ID
Public Property Let UpdatedAssocID(ByVal v_dblUpdatedAssocID As Variant)
    m_dblUpdatedAssocID = v_dblUpdatedAssocID
End Property

Public Property Get UpdatedAssocID() As Variant
    UpdatedAssocID = m_dblUpdatedAssocID
End Property
                                        ' The Associate ID of the first entry.
Public Property Let EnteredAssocID(ByVal v_dblEnteredAssocID As Variant)
    m_dblEnteredAssocID = v_dblEnteredAssocID
End Property

Public Property Get EnteredAssocID() As Variant
    EnteredAssocID = m_dblEnteredAssocID
End Property
                                        ' The status code.
Public Property Let StatusCd(ByVal v_strStatusCd As Variant)
    m_strStatusCd = v_strStatusCd
End Property

Public Property Get StatusCd() As Variant
    StatusCd = m_strStatusCd
End Property
                                        ' The status description.
Public Property Let StatusDesc(ByVal v_strStatusDesc As Variant)
    m_strStatusDesc = v_strStatusDesc
End Property

Public Property Get StatusDesc() As Variant
    StatusDesc = m_strStatusDesc
End Property
                                        ' The item number
Public Property Let ItemNbr(ByVal v_strItemNbr As Variant)
    m_strItemNbr = v_strItemNbr
End Property

Public Property Get ItemNbr() As Variant
    ItemNbr = m_strItemNbr
End Property
                                        ' The planned start date
Public Property Let PlannedStartDt(ByVal v_dtmPlannedStartDate As Variant)
    m_dtmPlannedStartDt = v_dtmPlannedStartDate
End Property

Public Property Get PlannedStartDt() As Variant
    PlannedStartDt = m_dtmPlannedStartDt
End Property
                                        ' The begin planned start date
Public Property Let BeginPlannedStartDt(ByVal v_dtmBeginPlannedStartDate As Variant)
    m_dtmBeginPlannedStartDt = v_dtmBeginPlannedStartDate
End Property

Public Property Get BeginPlannedStartDt() As Variant
    BeginPlannedStartDt = m_dtmBeginPlannedStartDt
End Property
                                        ' The end planned start date
Public Property Let EndPlannedStartDt(ByVal v_dtmEndPlannedStartDate As Variant)
    m_dtmEndPlannedStartDt = v_dtmEndPlannedStartDate
End Property

Public Property Get EndPlannedStartDt() As Variant
    EndPlannedStartDt = m_dtmEndPlannedStartDt
End Property
                                        ' The due date
Public Property Let DueDt(ByVal v_dtmDueDate As Variant)
    m_dtmDueDt = v_dtmDueDate
End Property

Public Property Get DueDt() As Variant
    DueDt = m_dtmDueDt
End Property
                                        ' The begin due date
Public Property Let BeginDueDt(ByVal v_dtmBeginDueDate As Variant)
    m_dtmBeginDueDt = v_dtmBeginDueDate
End Property

Public Property Get BeginDueDt() As Variant
    BeginDueDt = m_dtmBeginDueDt
End Property
                                        ' The end due date
Public Property Let EndDueDt(ByVal v_dtmEndDueDate As Variant)
    m_dtmEndDueDt = v_dtmEndDueDate
End Property

Public Property Get EndDueDt() As Variant
    EndDueDt = m_dtmEndDueDt
End Property
                                        ' Returning recordset
Public Property Set RecordSet(ByVal v_rstRecordSet As ADODB.RecordSet)
    Set m_rstRecordSet = v_rstRecordSet
End Property

Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = m_rstRecordSet
End Property

Public Sub GetProductionRun()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm           As String, _
        strSql                  As String, _
        rstRecordSet            As ADODB.RecordSet
    
    Let strRoutine_nm = "prod_run.cls.GetProductionRun"
    
    If IsNull(ProductionRunID) = True Then
        Let ProductionRunID = Null
        Exit Sub
    Else
        strSql = "SELECT a.production_run_id, a.ordered_qty, " & _
                 "       a.completed_qty, b.item_id, " & _
                 "       a.assem_ord_id, a.status_cd, " & _
                 "       a.planned_start_dt, b.due_dt " & _
                 "FROM prod_run a, assembly_order b " & _
                 "WHERE " & _
                 "      a.assem_ord_id = b.assem_ord_id AND " & _
                    fWhereVariable("production_run_id", ProductionRunID, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(rstRecordSet, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstRecordSet.Close
        Set rstRecordSet = Nothing
        Exit Sub
    End If
    
    If rstRecordSet.EOF = True Then
        Let ProductionRunID = Null
    Else
        ProductionRunID = rstRecordSet("production_run_id")
        OrderedQty = rstRecordSet("ordered_qty")
        CompletedQty = rstRecordSet("completed_qty")
        ItemID = rstRecordSet("item_id")
        ItemNbr = fGetItemNbr(rstRecordSet("item_id"), _
                                o_error, o_trans_data)
        AssemOrdID = rstRecordSet("assem_ord_id")
        StatusCd = rstRecordSet("status_cd")
        PlannedStartDt = rstRecordSet("planned_start_dt")
        DueDt = rstRecordSet("due_dt")
        StatusDesc = fGetDesc("prod_run_status_cd", rstRecordSet("status_cd"), o_error, o_trans_data)
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


Public Sub GetProductionRuns()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        objRecordset        As ADODB.RecordSet
    
    strRoutine_nm = "prod_run.cls.GetProductionRuns"
    
    strSql = "select a.production_run_id, a.ordered_qty, " & _
             "       a.completed_qty, c.item_id, " & _
             "       a.assem_ord_id, a.status_cd, " & _
             "       a.planned_start_dt, b.type_cd, " & _
             "       c.due_dt " & _
             "FROM prod_run a, item b, assembly_order c " & _
             "WHERE a.assem_ord_id = c.assem_ord_id AND " & _
             "      c.item_id = b.item_id "
    
    If Not IsNull(ProductionRunID) Then
        strSql = strSql & " and a.production_run_id = " & ProductionRunID
    End If
    
    If Not IsNull(AssemOrdID) Then
        strSql = strSql & " and a.assem_ord_id = " & AssemOrdID
    End If
    
    If Not IsNull(ItemNbr) Then
        strSql = strSql & " and c.item_id = " & fGetItemId(ItemNbr, o_error, o_trans_data)
    End If
        
    If Not IsNull(StatusCd) Then
        strSql = strSql & " and a.status_cd = '" & StatusCd & "' "
    End If
    
    If Not IsNull(BeginPlannedStartDt) And Not IsNull(EndPlannedStartDt) Then
        strSql = strSql & " and " & fWhereVariable("a.planned_start_dt", BeginPlannedStartDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("a.planned_start_dt", EndPlannedStartDt, "D", "<=")
    End If
    
    If Not IsNull(BeginDueDt) And Not IsNull(EndDueDt) Then
        strSql = strSql & " and " & fWhereVariable("c.due_dt", BeginDueDt, "D", ">=")
        strSql = strSql & " and " & fWhereVariable("c.due_dt", EndDueDt, "D", "<=")
    End If
    
    strSql = strSql & " ORDER BY planned_start_dt asc, production_run_id asc"
        
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If

    Set RecordSet = objRecordset

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


Public Sub AddProductionRun()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    Let strRoutine_nm = "prod_run.cls.Add"
    
    If ValidateData("Add") = False Then
        Exit Sub
    End If
    
    Call IssueInsertData
        
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


Public Sub ChangeProductionRun()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    strRoutine_nm = "prod_run.cls.ChangeProductionRun"
        
    Dim objAssemblyOrder As assembly_ord
        
    Call GetOldProductionRun
    
    If ValidateData("Change") = False Then
        Exit Sub
    End If
    
    Call IssueUpdateData
        
    If Len(Me.OrderedQty) > 0 And _
       Me.OrderedQty <> Me.o_old_production_run.OrderedQty Then
        Call SetStatus( _
                      Me.o_error _
                    , Me.o_trans_data _
                    , Me.ProductionRunID _
                    )
                    
        Set objAssemblyOrder = New assembly_ord
        Call objAssemblyOrder.UpdateQuantities( _
                                  Me.o_error _
                                , Me.o_trans_data _
                                , Me.AssemOrdID _
                                )
        Set objAssemblyOrder = Nothing
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


Public Sub DeleteProductionRun()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    Let strRoutine_nm = "prod_run.cls.Delete"
    
    Call GetOldProductionRun
    
    If Me.o_old_production_run.StatusCd <> "PL" Then
        o_error.p_type_cd = "E"
        o_error.p_err_cd = "0550"
        o_error.p_message_id = 4001
        o_error.p_routine_nm = strRoutine_nm
        Exit Sub
    End If
    
    Call IssueDeleteData
    
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

Public Sub OpenProductionRun()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double
        
    strRoutine_nm = "Prod_run.cls.OpenProductionRun"
                                        ' Set Item ID Property
    Me.StatusCd = "OP"
    Call ChangeStatus
    
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

Public Sub ChangeStatus()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double, _
        objAssemblyOrder    As assembly_ord
    
    Let strRoutine_nm = "prod_run.cls.changestatus"
    
    Call GetOldProductionRun
    
    If Me.o_old_production_run.StatusCd = "PL" Then
        If Me.StatusCd = "OP" Or _
           Me.StatusCd = "CA" Then
        Else
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "0550"
                .p_message_id = 3306
                .p_routine_nm = strRoutine_nm
            End With
        End If
    End If
                                        
    If Me.o_old_production_run.StatusCd = "OP" Then
        If Me.StatusCd = "OP" Or _
           Me.StatusCd = "OH" Or _
           Me.StatusCd = "CL" Or _
           Me.StatusCd = "CA" Then
        Else
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "0550"
                .p_message_id = 3306
                .p_routine_nm = strRoutine_nm
            End With
        End If
    End If
                                        
    If Me.o_old_production_run.StatusCd = "CL" Then
        If Me.StatusCd <> "OP" Then
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "0550"
                .p_message_id = 3306
                .p_routine_nm = strRoutine_nm
            End With
        End If
    End If
                                        
    If Me.o_old_production_run.StatusCd = "OH" Then
        If StatusCd <> "OP" Then
            With Me.o_error
                .p_type_cd = "E"
                .p_err_cd = "0550"
                .p_message_id = 3306
                .p_routine_nm = strRoutine_nm
            End With
        End If
    End If
                                        
    If Me.o_old_production_run.StatusCd = "CA" Then
        With Me.o_error
            .p_type_cd = "E"
            .p_err_cd = "0550"
            .p_message_id = 3306
            .p_routine_nm = strRoutine_nm
        End With
    End If
                                                                                   
    If Me.StatusCd = "CA" And Me.o_old_production_run.CompletedQty > 0 Then
        With Me.o_error
            .p_type_cd = "E"
            .p_err_cd = "0550"
            .p_message_id = 0
            .p_routine_nm = strRoutine_nm
            .p_desc = "Production Run can't be cancelled because inventory has been recieved against it.  Close it instead"
        End With
    End If
                                                                                   
    If Me.StatusCd = "CL" And Me.o_old_production_run.CompletedQty < 1 Then
        With Me.o_error
            .p_type_cd = "E"
            .p_err_cd = "0550"
            .p_message_id = 0
            .p_routine_nm = strRoutine_nm
            .p_desc = "Production Run can't be closed because inventory has not been recieved against it. Cancel it instead"
        End With
    End If
                                                                                   
                                        ' Set Item ID Property
    strSql = "update prod_run set "
    strSql = strSql & fUpdateVariable("status_cd", Me.StatusCd, "S", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("production_run_id", Me.ProductionRunID, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, Me.o_error)
    
    If Me.StatusCd = "OP" Then
        Call SetStatus( _
                  Me.o_error _
                , Me.o_trans_data _
                , Me.ProductionRunID _
                )
    End If
                                        ' Call Assembly Order Update Quantities
    Set objAssemblyOrder = New assembly_ord
    Call objAssemblyOrder.UpdateQuantities( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Me.AssemOrdID _
                            )
    Set objAssemblyOrder = Nothing
    
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


Private Sub IssueDeleteData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdateQty        As Double
        
    Let strRoutine_nm = "prod_run.cls.IssueDeleteData"

    Let strSql = "delete from prod_run "
    Let strSql = strSql & "where "
    Let strSql = strSql & " " & fWhereVariable("production_run_id", ProductionRunID, "N", "=")
    
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


Private Function ValidateData(ByVal strFunction_cd As String) As Boolean

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm        As String, _
        strSql              As String, _
        objAssemblyOrder    As Object, _
        objRecordset        As ADODB.RecordSet, _
        dblUpdateQty        As Double
    
    Let strRoutine_nm = "prod_run.cls.ValidateData"
    If strFunction_cd = "Add" Or strFunction_cd = "Change" Then
        If fValuePresent(AssemOrdID) = False Then
            Let o_error.p_type_cd = "E"
            Let o_error.p_err_cd = "0100"
            Let o_error.p_message_id = 2124
            Let o_error.p_routine_nm = strRoutine_nm
            Let o_error.p_desc = "Assembly Order ID requires a value to be passed"
            Let o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
    
    If strFunction_cd = "Add" Then
        strSql = "SELECT count(*), status_cd " & _
                 "FROM assembly_order " & _
                 "WHERE assem_ord_id = " & AssemOrdID
        strSql = strSql & " GROUP BY status_cd "
        Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            'objRecordset.Close
            Set objRecordset = Nothing
            ValidateData = False
            Exit Function
        End If
        
        If objRecordset(0) = 0 Then
            Let o_error.p_type_cd = "E"
            Let o_error.p_err_cd = "0550"
            Let o_error.p_message_id = 2126
            Let o_error.p_routine_nm = strRoutine_nm
            Let o_error.p_desc = "Assembly Order ID does not exist"
            Let o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        Else
            If objRecordset("status_cd") = "PL" Then
                Let o_error.p_type_cd = "E"
                Let o_error.p_err_cd = "0550"
                Let o_error.p_message_id = 2128
                Let o_error.p_routine_nm = strRoutine_nm
                Let o_error.p_desc = "Assembly Order is in a planned status"
                Let o_error.p_sqlstate_cd = ""
                ValidateData = False
                Exit Function
            End If
        End If
        
        'objRecordset.Close
        Set objRecordset = Nothing
    End If
    
    If strFunction_cd = "Add" Or strFunction_cd = "Change" Then
        If fValuePresent(OrderedQty) = False Then
            Let o_error.p_type_cd = "E"
            Let o_error.p_err_cd = "0450"
            Let o_error.p_message_id = 2122
            Let o_error.p_routine_nm = strRoutine_nm
            Let o_error.p_desc = "Ordered Quantity requires a value to be passed"
            Let o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
        
        
    If strFunction_cd = "Add" Or strFunction_cd = "Change" Then
        If fValuePresent(PlannedStartDt) = False Then
            Let o_error.p_type_cd = "E"
            Let o_error.p_err_cd = "0500"
            Let o_error.p_message_id = 2125
            Let o_error.p_routine_nm = strRoutine_nm
            Let o_error.p_desc = "Planned Start Date requires a value to be passed"
            Let o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
        
    ValidateData = True
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


Private Sub IssueInsertData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double
        
    strRoutine_nm = "prod_run.cls.IssueInsertData"
    
    ProductionRunID = fGetId("production_run_id", o_error, o_trans_data)
                                        ' Get Item ID
    ItemID = fGetItemId(ItemNbr, o_error, o_trans_data)
                                        ' Build Insert Statement
    strSql = "insert into prod_run ( "
    strSql = strSql & "  production_run_id"
    strSql = strSql & ", assem_ord_id"
    strSql = strSql & ", ordered_qty"
    strSql = strSql & ", completed_qty"
    strSql = strSql & ", status_cd"
    strSql = strSql & ", planned_start_dt"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(ProductionRunID, "N")
    strSql = strSql & "," & fInsertVariable(AssemOrdID, "N")
    strSql = strSql & "," & fInsertVariable(OrderedQty, "N")
    strSql = strSql & "," & fInsertVariable("0", "N")
    strSql = strSql & "," & fInsertVariable("PL", "S")
    strSql = strSql & "," & fInsertVariable(PlannedStartDt, "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & "," & fInsertVariable(o_trans_data.p_assoc_id, "N")
    strSql = strSql & "," & fInsertVariable("current_ts", "D")
    strSql = strSql & ")"
  
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, o_error)
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



Private Sub IssueUpdateData()

    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double
        
    strRoutine_nm = "Prod_run.cls.IssueUpdateData"
    
    strSql = "update prod_run set "
    strSql = strSql & fUpdateVariable("status_cd", StatusCd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("ordered_qty", OrderedQty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("completed_qty", CompletedQty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("assem_ord_id", AssemOrdID, "N", "=")
    strSql = strSql & "," & fUpdateVariable("planned_start_dt", PlannedStartDt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & " " & fWhereVariable("production_run_id", ProductionRunID, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, o_error)
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

Public Sub UpdateQuantities( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblProdRunId As Variant _
            )

    On Error GoTo error_handler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    strRoutine_nm = "Prod_run.cls.UpdateQuantities"
    
    Dim strSql              As String, _
        dblUpdatedqty       As Double, _
        dblSumAdjQty        As Double, _
        dblItemId           As Double, _
        rstInventoryAdj     As ADODB.RecordSet, _
        rstItem             As ADODB.RecordSet, _
        objAssemblyOrder    As assembly_ord
        
                                        
                                        ' Sum adjust quantity
    strSql = "SELECT sum(adj_qty) as sum_adj_qty " & _
             "FROM inventory_adj " & _
             "WHERE " & _
                fWhereVariable("prod_run_id", v_dblProdRunId, "N", "=") & _
             " and ac_line_nbr is null"
                
    Call r_objTransData.OpenRecordset(rstInventoryAdj, strSql, r_objError)
    
    If r_objError.p_type_cd = "E" Or r_objError.p_type_cd = "F" Then
        rstInventoryAdj.Close
        Set rstInventoryAdj = Nothing
        rstItem.Close
        Set rstItem = Nothing
        Exit Sub
    End If
    
    If rstInventoryAdj.EOF = True Then
        dblSumAdjQty = 0
    Else
        dblSumAdjQty = rstInventoryAdj("sum_adj_qty")
    End If
    
    rstInventoryAdj.Close
    Set rstInventoryAdj = Nothing
                                        ' Update prod_run with summed adjust qty
    strSql = "UPDATE prod_run SET "
    strSql = strSql & _
        fUpdateVariable("completed_qty", dblSumAdjQty, "N", "=")
    strSql = strSql & " WHERE "
    strSql = strSql & fWhereVariable("production_run_id", v_dblProdRunId, "N", "=")
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
                                        
    Call SetStatus( _
              r_objError _
            , r_objTransData _
            , v_dblProdRunId _
            )
                    ' Call Assembly Order Update Quantities
                    ' note that set status instantiates the object
                    ' so, me.assemordid is accurate
    Set objAssemblyOrder = New assembly_ord
    Call objAssemblyOrder.UpdateQuantities( _
                              r_objError _
                            , r_objTransData _
                            , Me.AssemOrdID _
                            )
        
    Set objAssemblyOrder = Nothing
                                        ' Set status
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


Private Sub SetStatus( _
              ByRef r_objError As Object _
            , ByRef r_objTransData As Object _
            , ByVal v_dblProdRunId As Variant _
            )

    On Error GoTo error_handler
    
    If r_objError.type_cd = "E" Or r_objError.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String
    strRoutine_nm = "Prod_run.cls.SetStatus"
    
    Dim strSql              As String, _
        strStatusCd         As String, _
        dblUpdatedqty       As Double, _
        dblSumAdjQty        As Double, _
        rstProdRun          As ADODB.RecordSet
   
                    'Select quantities
    Set Me.o_error = r_objError
    Set Me.o_trans_data = r_objTransData
    Me.ProductionRunID = v_dblProdRunId
    Call Me.GetProductionRun
    
    If Me.StatusCd = "CA" Or _
        Me.StatusCd = "OH" Or _
        Me.StatusCd = "PL" Or _
        Me.StatusCd = "CL" Then
        Exit Sub
    End If
    
    If Me.CompletedQty >= Me.OrderedQty Then
       strStatusCd = "CP"
    Else
        strStatusCd = "OP"
    End If
        
    If Me.StatusCd = strStatusCd Then
        Exit Sub
    End If
                                        ' Update prod_run with summed adjust qty
    strSql = "UPDATE prod_run SET "
    strSql = strSql & _
        fUpdateVariable("status_cd", strStatusCd, "S", "=")
    strSql = strSql & " WHERE "
    strSql = strSql & fWhereVariable("production_run_id", Me.ProductionRunID, "N", "=")
    Call r_objTransData.ExecuteDatabase(strSql, dblUpdatedqty, r_objError)
    
    Exit Sub
error_handler:
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



Private Sub GetOldProductionRun()
    
    On Error GoTo error_handler
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double, _
        dblSumAdjQty        As Double, _
        rstProdRun          As ADODB.RecordSet
        
    strRoutine_nm = "Prod_run.cls.GetOldProductionRun"
    
    If Me.o_old_production_run Is Nothing Then
        Set Me.o_old_production_run = New prod_run
        Set Me.o_old_production_run.o_error = o_error
        Set Me.o_old_production_run.o_trans_data = o_trans_data
        Me.o_old_production_run.ProductionRunID = ProductionRunID
        Me.o_old_production_run.GetProductionRun
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

