VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "assembly_comp"
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
'               order components.
'//////////////////////////////////////////////////////////////////////////////

Private m_intAssemOrdID             As Variant, _
        m_dblItemID                 As Variant, _
        m_intLineNbr                As Variant, _
        m_dblRequiredQty            As Variant, _
        m_intSupplierID             As Variant, _
        m_intLocID                  As Variant, _
        m_dblUpdatedAssocID         As Variant, _
        m_dblEnteredAssocID         As Variant, _
        m_strUomCd                  As Variant, _
        m_strUomDesc                As Variant, _
        m_strStatusCd               As Variant, _
        m_strStatusDesc             As Variant, _
        m_strItemNbr                As Variant, _
        m_strTypeCD                 As Variant, _
        m_strTypeDesc               As Variant, _
        m_dtmRequiredDt             As Variant, _
        m_objError                  As Object, _
        m_objTransData              As Object, _
        m_rstRecordSet              As ADODB.RecordSet
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
                                        ' The assembly order ID
Public Property Let AssemOrdID(ByVal v_intAssemOrdID As Variant)
    m_intAssemOrdID = v_intAssemOrdID
End Property

Public Property Get AssemOrdID() As Variant
    AssemOrdID = m_intAssemOrdID
End Property
                                        ' The item ID
Public Property Let ItemID(ByVal v_intItemid As Variant)
    m_dblItemID = v_intItemid
End Property

Public Property Get ItemID() As Variant
    ItemID = m_dblItemID
End Property
                                        ' The required qty
Public Property Let RequiredQty(ByVal v_intRequiredQty As Variant)
    m_dblRequiredQty = v_intRequiredQty
End Property

Public Property Get RequiredQty() As Variant
    RequiredQty = m_dblRequiredQty
End Property
                                        ' The supplier ID
Public Property Let SupplierID(ByVal v_intSupplierID As Variant)
    m_intSupplierID = v_intSupplierID
End Property

Public Property Get SupplierID() As Variant
    SupplierID = m_intSupplierID
End Property
                                        ' The location ID
Public Property Let LocID(ByVal v_intLocID As Variant)
    m_intLocID = v_intLocID
End Property

Public Property Get LocID() As Variant
    LocID = m_intLocID
End Property
                                        ' The line number
Public Property Let LineNbr(ByVal v_intLineNbr As Variant)
    m_intLineNbr = v_intLineNbr
End Property

Public Property Get LineNbr() As Variant
    LineNbr = m_intLineNbr
End Property
                                        ' The Updating Associate ID
Public Property Let UpdatedAssocID(ByVal v_intUpdatedAssocID As Variant)
    m_dblUpdatedAssocID = v_intUpdatedAssocID
End Property

Public Property Get UpdatedAssocID() As Variant
    UpdatedAssocID = m_dblUpdatedAssocID
End Property
                                        ' The Associate ID of the first entry.
Public Property Let EnteredAssocID(ByVal v_intEnteredAssocID As Variant)
    m_dblEnteredAssocID = v_intEnteredAssocID
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
                                        ' The unit of measure.
Public Property Let UomCd(ByVal v_strUomCd As Variant)
    m_strUomCd = v_strUomCd
End Property

Public Property Get UomCd() As Variant
    UomCd = m_strUomCd
End Property
                                        ' The unit of measure description
Public Property Let UomDesc(ByVal v_strUomDesc As Variant)
    m_strUomDesc = v_strUomDesc
End Property

Public Property Get UomDesc() As Variant
    UomDesc = m_strUomDesc
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
                                        ' The close date.
Public Property Let RequiredDt(ByVal v_dtmRequiredDt As Variant)
    m_dtmRequiredDt = v_dtmRequiredDt
End Property

Public Property Get RequiredDt() As Variant
    RequiredDt = m_dtmRequiredDt
End Property
                                        ' Returning recordset
Public Property Set RecordSet(ByVal r_rstRecordSet As ADODB.RecordSet)
    Set m_rstRecordSet = r_rstRecordSet
End Property
Public Property Get RecordSet() As ADODB.RecordSet
    Set RecordSet = m_rstRecordSet
End Property
                                        ' The item type code.
Public Property Let TypeCd(ByVal v_strTypeCd As Variant)
    m_strTypeCD = v_strTypeCd
End Property

Public Property Get TypeCd() As Variant
    TypeCd = m_strTypeCD
End Property
                                        ' The item type description.
Public Property Let TypeDesc(ByVal v_strTypeDesc As Variant)
    m_strTypeDesc = v_strTypeDesc
End Property

Public Property Get TypeDesc() As Variant
    TypeDesc = m_strTypeDesc
End Property

Public Sub GetAssemblyComponent()
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_comp.cls.GetAssemblyOrderComp"
  
    Call SelectAssemblyOrderComp
    
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

Public Sub GetAssemblyOrderComps()
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_comp.cls.GetAssemblyOrderComp"
    
    Dim strSql As String, _
        rstRecordSet As ADODB.RecordSet
    
    strSql = "SELECT a.assem_ord_id " & _
             " ,a.status_cd as assem_status_cd " & _
             " ,a.required_dt, a.uom_cd, a.line_nbr " & _
             " ,a.per_item_qty, a.required_qty " & _
             " ,a.supplier_id " & _
             " ,b.type_cd " & _
             " ,b.item_nbr " & _
             " ,b.item_id " & _
             " ,c.status_cd as mat_req_status_cd " & _
             " ,c.back_ordered_qty " & _
             " ,c.ready_qty " & _
             " ,c.consumed_qty " & _
             " FROM (assembly_comp a INNER JOIN item b ON a.item_id = b.item_id) " & _
             " LEFT JOIN mat_req c ON (a.line_nbr = c.ac_line_nbr) " & _
             " AND (a.assem_ord_id = c.assem_ord_id)" & _
             " where a.assem_ord_id = " & AssemOrdID
    
    Call Me.o_trans_data.OpenRecordset(rstRecordSet, strSql, Me.o_error)
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        rstRecordSet.Close
        Set rstRecordSet = Nothing
        Exit Sub
    End If
    
    Set RecordSet = rstRecordSet
    
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

Public Sub AddComponents()
           
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm       As String
    strRoutine_nm = "assembly_comp.cls.AddComponents"
    
    Call ValidateData("Add")
    Call GetLineNbr
    Call IssueInsertData
    
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


Public Sub Change()

    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "assembly_comp.cls.Change"
    
    Dim objMatReq As mat_req
    
    Call ValidateData("Change")
    Call IssueUpdateData
    Call Me.GetAssemblyComponent
    Set objMatReq = New mat_req
    Call objMatReq.Change( _
                  Me.o_error _
                , Me.o_trans_data _
                , Empty _
                , "AO" _
                , Empty _
                , Empty _
                , Me.AssemOrdID _
                , Me.LineNbr _
                , Me.RequiredQty _
                , Me.RequiredDt _
                , Empty _
                , Empty _
                , Empty _
                , Empty _
                )
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


Public Sub Delete()
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm           As String
    strRoutine_nm = "assembly_comp.cls.Delete"
    
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

Private Sub GetLineNbr()
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Dim rstRecordSet        As ADODB.RecordSet, _
        strRoutine_nm       As String, _
        dblLineNbr          As Double, _
        strSql              As String

    strRoutine_nm = "assembly_comp.cls.GetLineNbr"
    
    strSql = "SELECT line_nbr FROM assembly_comp " & _
               " WHERE " & fWhereVariable("assem_ord_id", AssemOrdID, "N", "=") & _
               " ORDER BY line_nbr desc "
    Call Me.o_trans_data.OpenRecordset(rstRecordSet, strSql, Me.o_error)
    If rstRecordSet.EOF Then
        LineNbr = 1
    Else
        dblLineNbr = rstRecordSet("line_nbr")
    
        If IsNull(dblLineNbr) = True Then
            LineNbr = 1
        Else
            LineNbr = dblLineNbr + 1
        End If
    End If
    rstRecordSet.Close
    Set rstRecordSet = Nothing
    
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

Private Sub IssueInsertData()

    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double, _
        dblRequiredQty      As Double, _
        dtmRequiredDt       As Variant, _
        clsBOM              As Bill_of_Materials, _
        rstBOM              As ADODB.RecordSet, _
        rstItem             As ADODB.RecordSet
        
    strRoutine_nm = "assembly_comp.cls.IssueInsertData"
                                        ' Compute required date
                                        ' Explode Bill of Material one level
                                        ' to get assembly components
    Set clsBOM = New Bill_of_Materials
    clsBOM.ParentPartNbr = ItemNbr
    Set clsBOM.o_error = Me.o_error
    Set clsBOM.o_trans_data = Me.o_trans_data
    Call clsBOM.BOMExplosion_SingleLevel
    Set rstBOM = clsBOM.BOM
    
    Set clsBOM = Nothing
                                        ' If no components exists, exit sub
    If rstBOM.EOF Then
        Exit Sub
    End If
                                        ' Scan the BOM
    rstBOM.MoveFirst
    Do While Not rstBOM.EOF
        'dtmRequiredDt = ComputeRequiredDt(rstBOM("comp_item_id"))
         dtmRequiredDt = Me.RequiredDt
         
                                        ' Get the item information for the
                                        ' component
        strSql = "SELECT uom_cd, preferred_supplier_id " & _
                 "FROM item " & _
                 "WHERE item_id = " & rstBOM("comp_item_id")
        Call Me.o_trans_data.OpenRecordset(rstItem, strSql, Me.o_error)
        If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
            Set rstItem = Nothing
            Exit Sub
        End If
                                        ' Set Item ID
        Call GetLineNbr
                                        ' Compute required qty
        dblRequiredQty = RequiredQty * rstBOM("item_qty")
                                        ' Build Insert statement
        strSql = "insert into assembly_comp( "
        strSql = strSql & "  assem_ord_id"
        strSql = strSql & ", line_nbr"
        strSql = strSql & ", status_cd"
        strSql = strSql & ", item_id"
        strSql = strSql & ", required_qty"
        strSql = strSql & ", per_item_qty"
        strSql = strSql & ", required_dt"
        strSql = strSql & ", uom_cd"
        strSql = strSql & ", supplier_id"
        strSql = strSql & ", updated_assoc_id"
        strSql = strSql & ", updated_ts"
        strSql = strSql & ", entered_assoc_id"
        strSql = strSql & ", entered_ts"
        strSql = strSql & " ) "
        strSql = strSql & "values ("
        strSql = strSql & " " & fInsertVariable(AssemOrdID, "N")
        strSql = strSql & "," & fInsertVariable(LineNbr, "N")
        strSql = strSql & "," & fInsertVariable("PL", "S")
        strSql = strSql & "," & fInsertVariable(rstBOM("comp_item_id"), "N")
        strSql = strSql & "," & fInsertVariable(dblRequiredQty, "N")
        strSql = strSql & "," & fInsertVariable(rstBOM("item_qty"), "N")
        strSql = strSql & "," & fInsertVariable(dtmRequiredDt, "D")
        strSql = strSql & "," & fInsertVariable(rstItem("uom_cd"), "S")
        strSql = strSql & "," & fInsertVariable(rstItem("preferred_supplier_id"), "N")
        strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
        strSql = strSql & "," & fInsertVariable("current_ts", "D")
        strSql = strSql & "," & fInsertVariable(Me.o_trans_data.p_assoc_id, "N")
        strSql = strSql & "," & fInsertVariable("current_ts", "D")
        strSql = strSql & ")"
                                        ' Add the component
        Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, Me.o_error)
                                        ' Move to the next component on Bill
        rstBOM.MoveNext
        
        rstItem.Close
        Set rstItem = Nothing
    Loop
    rstBOM.Close
    Set rstBOM = Nothing
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

Private Function ValidateData(strFunction_cd) As Boolean
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm           As String, _
        strSql                  As String
        
    strRoutine_nm = "assembly_comp.cls.ValidateAdd"
    
    If strFunction_cd = "Add" Or strFunction_cd = "Change" Then
        If fValuePresent(AssemOrdID) = False Then
            Let Me.o_error.p_type_cd = "E"
            Let Me.o_error.p_err_cd = "0100"
            Let Me.o_error.p_message_id = 2124
            Let Me.o_error.p_routine_nm = strRoutine_nm
            Let Me.o_error.p_desc = "Assembly ID requires a value to be passed"
            Let Me.o_error.p_sqlstate_cd = ""
            ValidateData = False
            Exit Function
        End If
    End If
    ValidateData = True
    Exit Function
errHandler:
    ValidateData = False
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

Private Sub IssueUpdateData()
    
    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm       As String
    strRoutine_nm = "assembly_comp.cls.IssueUpdateData"
    
    Dim strSql As String, _
        dblUpdatedqty As Double
    
    strSql = "update assembly_comp set "
    strSql = strSql & " " & fUpdateVariable("required_dt", RequiredDt, "D", "=")
    strSql = strSql & "," & fUpdateVariable("required_qty", RequiredQty, "N", "=")
    strSql = strSql & "," & fUpdateVariable("supplier_id", SupplierID, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Me.o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "    " & fWhereVariable("assem_ord_id", AssemOrdID, "N", "=")
    strSql = strSql & "and " & fWhereVariable("line_nbr", Me.LineNbr, "N", "=")
    
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
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm       As String, _
        strSql              As String, _
        dblUpdatedqty       As Double
        
    strRoutine_nm = "assembly_comp.cls.IssueDeleteData"

    Let strSql = "delete from assembly_comp "
    Let strSql = strSql & "where "
    Let strSql = strSql & " " & fWhereVariable("assem_ord_id", AssemOrdID, "N", "=")
    Let strSql = strSql & " AND " & fWhereVariable("line_nbr", LineNbr, "N", "=")
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:   This function will change the item's status equal to the status
'               of the purchase order record.
'
'Arguments:     none
'//////////////////////////////////////////////////////////////////////////////
Public Sub ChangeStatus()

    On Error GoTo errHandler
    
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Exit Sub
    End If
    
    Dim objRecordset As ADODB.RecordSet
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdatedqty As Double
    Dim objMatReq As mat_req
    Dim objInventoryAlloc As Object
    strRoutine_nm = "Assembly_comp.cls.changestatus"
                
    If Me.StatusCd = "OP" Or _
       Me.StatusCd = "CP" Or _
       Me.StatusCd = "CL" Or _
       Me.StatusCd = "OH" Or _
       Me.StatusCd = "CA" Then
    Else
        Me.o_error.p_type_cd = "E"
        Me.o_error.p_err_cd = "0200"
        Me.o_error.p_routine_nm = strRoutine_nm
        Me.o_error.p_message_id = 3301
        Exit Sub
    End If
    
    strSql = "update assembly_comp set "
    strSql = strSql & fUpdateVariable("status_cd", StatusCd, "S", "=")
    strSql = strSql & "," & fUpdateVariable("updated_assoc_id", Me.o_trans_data.p_assoc_id, "N", "=")
    strSql = strSql & "," & fUpdateVariable("updated_ts", "current_ts", "D", "=")
    strSql = strSql & " where"
    strSql = strSql & "   " & fWhereVariable("assem_ord_id", AssemOrdID, "N", "=")
    strSql = strSql & "and" & fWhereVariable("line_nbr", LineNbr, "N", "=")
    
    Call Me.o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, Me.o_error)
    
    If dblUpdatedqty = 0 Then
        Me.o_error.type_cd = "F"
        Me.o_error.err_cd = "0100"
        Me.o_error.routine_nm = strRoutine_nm
        Me.o_error.message_id = 1005
        Exit Sub
    End If
    
    Call GetAssemblyComponent
    
    Set objMatReq = New mat_req
                                        ' Mat req get this value.
    Select Case Me.StatusCd
        Case "OP"
            Set objMatReq.o_error = Me.o_error
            Set objMatReq.o_trans_data = Me.o_trans_data
            objMatReq.mat_req_id = Empty
            objMatReq.assem_ord_id = Me.AssemOrdID
            objMatReq.ac_line_nbr = Me.LineNbr
            objMatReq.sales_ord_id = Empty
            objMatReq.so_line_nbr = Empty
            Call objMatReq.GetMatReq
            
            If IsNull(objMatReq.mat_req_id) = True Then
                Call objMatReq.addMatReq( _
                          Me.o_error _
                        , Me.o_trans_data _
                        , "AO" _
                        , Empty _
                        , Empty _
                        , Me.AssemOrdID _
                        , Me.LineNbr _
                        , Me.ItemID _
                        , Me.LocID _
                        , Me.RequiredQty _
                        , Me.RequiredDt _
                        , 0 _
                        , 0 _
                        , Null _
                        )
            
                Call objMatReq.openMatReq( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Empty _
                            , "AO" _
                            , Empty _
                            , Empty _
                            , Me.AssemOrdID _
                            , Me.LineNbr _
                            )
                Call objMatReq.generatePlan
            Else
                Call objMatReq.openMatReq( _
                              Me.o_error _
                            , Me.o_trans_data _
                            , Empty _
                            , "AO" _
                            , Empty _
                            , Empty _
                            , Me.AssemOrdID _
                            , Me.LineNbr _
                            )
                
                    'after opening the material requirement, reallocate the
                    'inventory
                Set objInventoryAlloc = CreateObject("converge_inv.inventory_alloc")
                With objInventoryAlloc
                    Set .o_error = Me.o_error
                    Set .o_trans_data = Me.o_trans_data
                    .p_mat_req_id = objMatReq.mat_req_id
                    .p_item_id = objMatReq.item_id
                    .p_inv_loc_id = objMatReq.inv_loc_id
                    .Allocate
                End With
                Set objInventoryAlloc = Nothing
            End If
                        
        Case "CP"
            Call objMatReq.closeMatReq( _
                          o_error _
                        , o_trans_data _
                        , Empty _
                        , "AO" _
                        , Empty _
                        , Empty _
                        , Me.AssemOrdID _
                        , Me.LineNbr _
                        )
        Case "CL"
            Call objMatReq.closeMatReq( _
                          o_error _
                        , o_trans_data _
                        , Empty _
                        , "AO" _
                        , Empty _
                        , Empty _
                        , Me.AssemOrdID _
                        , Me.LineNbr _
                        )
        Case "OH"
            Call objMatReq.holdMatReq( _
                          o_error _
                        , o_trans_data _
                        , Empty _
                        , "AO" _
                        , Empty _
                        , Empty _
                        , Me.AssemOrdID _
                        , Me.LineNbr _
                        )

        Case "CA"
            Call objMatReq.cancelMatReq( _
                          o_error _
                        , o_trans_data _
                        , Empty _
                        , "AO" _
                        , Empty _
                        , Empty _
                        , Me.AssemOrdID _
                        , Me.LineNbr _
                        )
    End Select
    
    Set objMatReq = Nothing
    
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

Private Sub SelectAssemblyOrderComp()

    On Error GoTo errHandler
    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm           As String, _
        strSql                  As String, _
        rstRecordSet            As ADODB.RecordSet
    
    strRoutine_nm = "assembly_comp.cls.SelectAssemblyOrderComp"
        
    strSql = "SELECT a.assem_ord_id, a.line_nbr, "
    strSql = strSql & "a.item_id, a.required_qty, "
    strSql = strSql & "a.required_dt, a.supplier_id, "
    strSql = strSql & "a.uom_cd, a.status_cd, "
    strSql = strSql & "b.type_cd, c.inv_loc_id "
    strSql = strSql & "FROM assembly_comp a, item b, assembly_order c "
    strSql = strSql & "WHERE "
    strSql = strSql & "     a.assem_ord_id = c.assem_ord_id "
    strSql = strSql & " AND a.item_id = b.item_id "
    strSql = strSql & " AND " & fWhereVariable("a.assem_ord_id", Me.AssemOrdID, "N", "=")
    strSql = strSql & " AND " & fWhereVariable("a.line_nbr", Me.LineNbr, "N", "=")
    Call Me.o_trans_data.OpenRecordset(rstRecordSet, strSql, Me.o_error)

    If Me.o_error.p_type_cd = "E" Or Me.o_error.p_type_cd = "F" Then
        Set rstRecordSet = Nothing
        Exit Sub
    End If
    
    If rstRecordSet.EOF = True Then
        LineNbr = Null
    Else
        Me.AssemOrdID = rstRecordSet("assem_ord_id")
        Me.LineNbr = rstRecordSet("line_nbr")
        Me.ItemID = rstRecordSet("item_id")
        Me.TypeCd = rstRecordSet("type_cd")
        Me.TypeDesc = fGetDesc("item_type_cd", rstRecordSet("type_cd"), Me.o_error, Me.o_trans_data)
        Me.ItemNbr = fGetItemNbr(rstRecordSet("item_id"), Me.o_error, Me.o_trans_data)
        Me.RequiredQty = rstRecordSet("required_qty")
        Me.RequiredDt = rstRecordSet("required_dt")
        Me.SupplierID = rstRecordSet("supplier_id")
        Me.UomCd = rstRecordSet("uom_cd")
        Me.UomDesc = fGetDesc("uom_cd", rstRecordSet("uom_cd"), Me.o_error, Me.o_trans_data)
        Me.StatusCd = rstRecordSet("status_cd")
        Me.StatusDesc = fGetDesc("assembly_order_status_cd", rstRecordSet("status_cd"), Me.o_error, Me.o_trans_data)
        Me.LocID = rstRecordSet("inv_loc_id")
    End If
        
    rstRecordSet.Close
    Set rstRecordSet = Nothing
    
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

