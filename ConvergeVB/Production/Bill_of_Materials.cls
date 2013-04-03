VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bill_of_Materials"
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
'Description:   This class object is responsible for maintaining the bill of
'               material for an item.
'//////////////////////////////////////////////////////////////////////////////

Private m_varParentPartNbr          As Variant, _
        m_varComponentPartNbr       As Variant, _
        m_strRoutine_nm             As String, _
        m_strConnectString          As String, _
        m_intChaseSub               As Double, _
        m_rstBOM                    As ADODB.RecordSet, _
        m_rstBillOfMaterial         As ADODB.RecordSet, _
        m_objError                  As Object, _
        m_objTransData              As Object
                                        ' Transaction Object
Public Property Set o_trans_data(ByVal v_objTransData As Object)
    Set m_objTransData = v_objTransData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTransData
End Property
                                        ' Error Object
Public Property Set o_error(ByVal v_objError As Object)
    Set m_objError = v_objError
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
                                        ' The item number that is being exploded
Public Property Let ParentPartNbr(ByVal v_strParentPartNbr As String)
    m_varParentPartNbr = v_strParentPartNbr
End Property

Public Property Get ParentPartNbr() As String
    ParentPartNbr = m_varParentPartNbr
End Property
                                        ' The component item number that is being
                                        ' imploded
Public Property Let ComponentPartNbr(ByVal v_strComponentPartNbr As String)
    m_varComponentPartNbr = v_strComponentPartNbr
End Property

Public Property Get ComponentPartNbr() As String
    ComponentPartNbr = m_varComponentPartNbr
End Property
                                        ' Returning Recordset
Public Property Set BOM(ByRef r_rstBillOfMaterial As ADODB.RecordSet)
    Set m_rstBillOfMaterial = r_rstBillOfMaterial
End Property
Public Property Get BOM() As ADODB.RecordSet
    Set BOM = m_rstBillOfMaterial
End Property
                                        ' The database connect string
Public Property Let ConnectString(ByVal v_strConnectString As String)
    m_strConnectString = v_strConnectString
End Property

Public Property Get ConnectString() As String
    ConnectString = m_strConnectString
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure explodes the bill of material for a given part
'               and builds the recordset for display purposes.
'
'Arguments:     none
'Pre-Requisite: This routine expects the following properties to have a value:
'               ParentPartNbr        -   Item to be exploded
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub BOMExplosion_MultiLevel()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm            As String
    Dim dblCtr      As Double
    strRoutine_nm = "bill_of_materials.cls.BOMExplosion_MultiLevel"
    
                                        ' Initiate Display Level Array
                                        ' Default Values
    m_intChaseSub = 1
                                        ' Create the Bill of Material recordset
                                        ' to be returned to the calling routine.
    Set BOM = New ADODB.RecordSet
    With BOM.Fields
        .Append "intLevelNbr", adInteger, , adFldUpdatable
        .Append "strParentPartNbr", adVarChar, 25, adFldUpdatable
        .Append "strComponentPartNbr", adVarChar, 25, adFldUpdatable
        .Append "intItemQty", adBigInt, , adFldUpdatable
        .Append "intItemCost", adDouble, , adFldUpdatable
        .Append "dblAssemCost", adDouble, , adFldUpdatable
    End With
    BOM.Open
                                        ' Traverse the Bill of Material and
                                        ' Build the recordset.
    Call Traverse_Bill(ParentPartNbr)
                                        ' Set the Bill of Material property equal
                                        ' to the newly created BOM recordset.
    BOM.MoveFirst
    
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure traverses through the bill of material for the
'               part number passed to it.
'
'Arguments:     strChasePartNbr     -   The part to traverse
'
'//////////////////////////////////////////////////////////////////////////////

Private Sub Traverse_Bill(ByVal v_strChasePartNbr As String)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm            As String
    Dim varComponentPartNbr      As Variant
    Dim rstBillTree              As ADODB.RecordSet
        
    strRoutine_nm = "Bill_of_Materials.cls.Traverse_Bill"
    
    Call ReadBillUsingParent(v_strChasePartNbr, rstBillTree)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    If rstBillTree.EOF Then
                                        ' Add the part to the recordset
                                        ' with no component
        BOM.AddNew
        BOM("intLevelNbr") = m_intChaseSub
        BOM("strParentPartNbr") = v_strChasePartNbr
        BOM("strComponentPartNbr") = ""
        BOM("intItemQty") = 0
        BOM("intItemCost") = 0
        ' JRP4
        BOM("dblAssemCost") = 0
                                        
                                        ' Drop down a level
        m_intChaseSub = m_intChaseSub - 1
        rstBillTree.Close
        Set rstBillTree = Nothing
        Exit Sub
    End If
    
    Do While Not rstBillTree.EOF
        varComponentPartNbr = fGetItemNbr(rstBillTree("comp_item_id"), o_error, o_trans_data)
        
        If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
            Exit Sub
        End If
        
        If IsNull(varComponentPartNbr) = True Then
            With o_error
                .p_type_cd = "F"
                .p_err_cd = "0100"
                .p_desc = Err.Description
                .p_routine_nm = strRoutine_nm
                .p_message_id = 2112
            End With
            Exit Sub
        End If
                                        ' Add the part to the recordset
        BOM.AddNew
        BOM("intLevelNbr") = m_intChaseSub
        BOM("strParentPartNbr") = v_strChasePartNbr
        BOM("strComponentPartNbr") = varComponentPartNbr
        BOM("intItemQty") = rstBillTree("item_qty")
        BOM("intItemCost") = rstBillTree("item_cost")

        ' JRP4
        BOM("dblAssemCost") = rstBillTree("assem_cost")


        m_intChaseSub = m_intChaseSub + 1
        Call Traverse_Bill(varComponentPartNbr)
        
        rstBillTree.MoveNext
    Loop
    
    rstBillTree.Close
    Set rstBillTree = Nothing

    m_intChaseSub = m_intChaseSub - 1
    
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure implodes the bill of material for a given part
'               and builds the recordset for display purposes.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub BOMImplosion_MultiLevel()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm            As String
    Dim dblCtr      As Double
    strRoutine_nm = "bill_of_materials.cls.BOMImplosion_MultiLevel"
                                        ' Initiate Display Level Array
                                        ' Default Values
    m_intChaseSub = 1
                                        ' Create the Bill of Material recordset
                                        ' to be returned to the calling routine.
    Set BOM = New ADODB.RecordSet
    With BOM.Fields
        .Append "intLevelNbr", adInteger, , adFldUpdatable
        .Append "strParentPartNbr", adVarChar, 25, adFldUpdatable
        .Append "strComponentPartNbr", adVarChar, 25, adFldUpdatable
        .Append "intItemQty", adBigInt, , adFldUpdatable
        .Append "dblItemCost", adDouble, , adFldUpdatable
        .Append "dblAssemCost", adDouble, , adFldUpdatable
    End With
    BOM.Open
                                        ' Traverse the Bill of Material and
                                        ' Build the recordset.
    Call Traverse_Bill_Implosion(ComponentPartNbr)
                                        ' Set the Bill of Material property equal
                                        ' to the newly created BOM recordset.
    BOM.MoveFirst
    
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure traverses up the bill of material for the
'               component part number passed to it.
'
'Arguments:     strChasePartNbr     -   The part to traverse
'
'//////////////////////////////////////////////////////////////////////////////

Private Sub Traverse_Bill_Implosion(ByVal v_strChasePartNbr As String)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strRoutine_nm            As String
    Dim varParentPartNbr         As Variant
    Dim rstBillTree              As ADODB.RecordSet
        
    strRoutine_nm = "Bill_of_Materials.cls.Traverse_Bill_Implosion"
    
    Call ReadBillUsingComponent(v_strChasePartNbr, rstBillTree)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    If rstBillTree.EOF Then
                                        ' Add the part to the recordset
                                        ' with no Parent Part Number.
        BOM.AddNew
        BOM("intLevelNbr") = m_intChaseSub
        BOM("strParentPartNbr") = ""
        BOM("strComponentPartNbr") = v_strChasePartNbr
        BOM("intItemQty") = 0
        BOM("intItemCost") = 0
        ' JRP4
        BOM("dblAssemCost") = 0
                                        
                                        ' Drop down a level
        m_intChaseSub = m_intChaseSub - 1
        rstBillTree.Close
        Set rstBillTree = Nothing
        Exit Sub
    End If
    
    Do While Not rstBillTree.EOF
        varParentPartNbr = fGetItemNbr(rstBillTree("comp_id"), o_error, o_trans_data)
        
        If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
            Exit Sub
        End If
        
        If IsNull(varParentPartNbr) = True Then
            With o_error
                .p_type_cd = "F"
                .p_err_cd = "0100"
                .p_desc = Err.Description
                .p_routine_nm = strRoutine_nm
                .p_message_id = 2111
            End With
            Exit Sub
        End If
                                        ' Add the part to the recordset
        BOM.AddNew
        BOM("intLevelNbr") = m_intChaseSub
        BOM("strParentPartNbr") = varParentPartNbr
        BOM("strComponentPartNbr") = v_strChasePartNbr
        BOM("intItemQty") = rstBillTree("item_qty")
        BOM("intItemCost") = rstBillTree("item_cost")
        ' JRP4
        BOM("dblAssemCost") = rstBillTree("assem_cost")


        m_intChaseSub = m_intChaseSub + 1
        Call Traverse_Bill_Implosion(varParentPartNbr)
        
        rstBillTree.MoveNext
    Loop
    
    rstBillTree.Close
    Set rstBillTree = Nothing

    m_intChaseSub = m_intChaseSub - 1
    
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure reads the bill of material file using the
'               parent part number passed to it.
'
'Arguments:     strChasePartNbr     -   The part to traverse
'
'//////////////////////////////////////////////////////////////////////////////

Private Sub ReadBillUsingParent(ByVal v_strChasePartNbr As String, ByRef rstBillTree As ADODB.RecordSet)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim varChaseItemId          As Variant, _
        strSql                  As String, _
        strRoutine_nm           As String
    
    strRoutine_nm = "Bill_of_Materials.cls.ReadBillUsingParent"
    
    varChaseItemId = fGetItemId(v_strChasePartNbr, o_error, o_trans_data)
    
    If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
        Exit Sub
    End If
    
    If IsNull(varChaseItemId) = True Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_nbr = Err.Number
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 3304
        End With
        Exit Sub
    End If
    
    strSql = "SELECT a.*, b.item_cost, b.assem_cost " & _
             " FROM bill_of_material a, item b" & _
             " WHERE a.item_id = " & varChaseItemId & " AND a.comp_item_id = b.item_id "
    Call o_trans_data.OpenRecordset(rstBillTree, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstBillTree.Close
        Set rstBillTree = Nothing
        Exit Sub
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure reads the bill of material file using the
'               component part number passed to it for implosion purposes.
'
'Arguments:     strChasePartNbr     -   The part to traverse
'
'//////////////////////////////////////////////////////////////////////////////

Private Sub ReadBillUsingComponent(ByVal v_strChaseComponentNbr As String, ByRef rstBillTree As ADODB.RecordSet)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim varChaseItemId          As Variant, _
        strSql                  As String, _
        strRoutine_nm           As String
    
    strRoutine_nm = "Bill_of_Materials.cls.ReadBillUsingComponent"
    
    varChaseItemId = fGetItemId(v_strChaseComponentNbr, o_error, o_trans_data)
    
    If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
        Exit Sub
    End If
    
    If IsNull(varChaseItemId) = True Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_nbr = Err.Number
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 0
        End With
        Exit Sub
    End If
    
    strSql = "SELECT a.*, b.item_cost, b.assem_cost " & _
             "FROM bill_of_material a, item b " & _
             "WHERE comp_item_id = " & varChaseItemId & _
             " AND a.comp_item_id = b.item_id "

    Call o_trans_data.OpenRecordset(rstBillTree, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstBillTree.Close
        Set rstBillTree = Nothing
        Exit Sub
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This function determines if component part numbers exists for
'               the part number passed to it.
'
'Arguments:     strPartNbr     -   The part to traverse
'
'//////////////////////////////////////////////////////////////////////////////

Private Function ComponentsExist(ByVal v_strPartNbr As String) As Boolean
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If
    
    Dim varChaseItemId          As Variant, _
        rstBill                 As ADODB.RecordSet, _
        strSql                  As String, _
        strRoutine_nm           As String
    
    strRoutine_nm = "Bill_of_Materials.cls.ComponentsExist"
    
    varChaseItemId = fGetItemId(v_strPartNbr, o_error, o_trans_data)
    
    If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
        Exit Function
    End If
      
    If IsNull(varChaseItemId) = True Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 0
        End With
        Exit Function
    End If
    
    strSql = "SELECT * " & _
             "FROM bill_of_material " & _
             "WHERE item_id = " & varChaseItemId
             
    Call o_trans_data.OpenRecordset(rstBill, strSql, o_error)
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstBill.Close
        Set rstBill = Nothing
        ComponentsExist = False
        Exit Function
    End If
    
    If rstBill.EOF Then
        ComponentsExist = False
    Else
        ComponentsExist = True
    End If
    
    rstBill.Close
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure explodes the bill, one level down.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub BOMExplosion_SingleLevel()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim varParentItemID         As Variant, _
        strComponentItemNbr     As String, _
        strRoutine_nm           As String, _
        strSql                  As String, _
        rstWorkingBOM           As ADODB.RecordSet, _
        rstItem                 As ADODB.RecordSet
    
    strRoutine_nm = "Bill_of_Material.cls.BOMExplosion_SingleLevel"
                                        ' Obtain Parent ID from Parent
                                        ' Part Number
    varParentItemID = fGetItemId(ParentPartNbr, o_error, o_trans_data)
    
    If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
        Exit Sub
    End If
                                        
    If IsNull(varParentItemID) = True Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 0
        End With
        Exit Sub
    End If
                                        ' Build Recordset
    strSql = "SELECT a.item_id, a.comp_item_id,  " & _
             "       a.item_qty, b.item_nbr, b.item_desc, b.item_cost " & _
             "FROM bill_of_material a, item b " & _
             "WHERE a.item_id = " & CStr(varParentItemID) & _
             " and a.comp_item_id = b.item_id "
             
    Call o_trans_data.OpenRecordset(rstWorkingBOM, strSql, o_error)
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        Set rstWorkingBOM = Nothing
        Exit Sub
    End If
    
    Set BOM = rstWorkingBOM
   
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



