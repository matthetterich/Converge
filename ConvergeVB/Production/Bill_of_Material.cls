VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bill_of_Material"
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

Private m_varParentPartNbr              As Variant, _
        m_varComponentPartNbr           As Variant, _
        m_strConnectString              As String, _
        m_dblItemQty                    As Double, _
        m_dblAssociateID                As Double, _
        m_objError                      As Object, _
        m_objTransData                  As Object
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
                                        ' The Parent Part Number
Public Property Let ParentPartNbr(ByVal v_strParentPartNbr As String)
    m_varParentPartNbr = v_strParentPartNbr
End Property

Public Property Get ParentPartNbr() As String
    ParentPartNbr = m_varParentPartNbr
End Property
                                        ' The component part being added to the
                                        ' parent part number.
Public Property Let ComponentPartNbr(ByVal v_strParentPartNbr As String)
    m_varComponentPartNbr = v_strParentPartNbr
End Property

Public Property Get ComponentPartNbr() As String
    ComponentPartNbr = m_varComponentPartNbr
End Property
                                        ' The connect string to the database.
Public Property Let ConnectString(ByVal v_strConnectString As String)
    m_strConnectString = v_strConnectString
End Property

Public Property Get ConnectString() As String
    ConnectString = m_strConnectString
End Property
                                        ' The quantity per for the component
Public Property Let ItemQty(ByVal v_dblItemQty As Double)
    m_dblItemQty = v_dblItemQty
End Property

Public Property Get ItemQty() As Double
    ItemQty = m_dblItemQty
End Property
                                        ' The associate id
Public Property Let AssociateID(ByVal v_dblAssociateID As Double)
    m_dblAssociateID = v_dblAssociateID
End Property

Public Property Get AssociateID() As Double
    AssociateID = m_dblAssociateID
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   This procedure adds a component part to a parent part within
'               the bill of material.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub AddToBill()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strSql              As String, _
        strRoutine_nm       As String, _
        varParentItemID     As Variant, _
        varComponentItemID  As Variant, _
        dblUpdateQty        As Double, _
        rstBill             As ADODB.RecordSet
        
        
    
    strRoutine_nm = "Bill_of_Material.cls.AddToBill"
        
    If Not ValidArguments(varParentItemID, varComponentItemID, "AddToBill") Then
        Exit Sub
    End If
    
    
    strSql = "INSERT INTO bill_of_material " & _
             "(item_id, comp_item_id, item_qty, " & _
             " updated_assoc_id, updated_ts, entered_assoc_id, " & _
             " entered_ts) " & _
             "VALUES " & _
             "(" & CStr(varParentItemID) & ", " & _
                   CStr(varComponentItemID) & ", " & _
                   CStr(ItemQty) & ", " & _
                   CStr(AssociateID) & ", " & _
             "'" & CStr(Now()) & "', " & _
                   CStr(AssociateID) & ", " & _
             "'" & CStr(Now()) & "')"
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
    Call updateItemCost(varParentItemID)

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
'Description:   This procedure modifies the component part for a parent part
'               within the bill of material.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub ModifyBill()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strSql                          As String, _
        strRoutine_nm                   As String, _
        varParentItemID                 As Variant, _
        varComponentItemID              As Variant, _
        dblUpdatedqty                   As Double, _
        rstBill                         As ADODB.RecordSet
    
    strRoutine_nm = "Bill_of_Material.cls.ModifyBill"
        
    If Not ValidArguments(varParentItemID, varComponentItemID, "ModifyBill") Then
        Exit Sub
    End If
            
    strSql = "UPDATE bill_of_material SET " & _
             " item_qty         = " & CStr(ItemQty) & _
             " WHERE item_id = " & CStr(varParentItemID) & " AND " & _
             "       comp_item_id = " & CStr(varComponentItemID)
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, o_error)
    
    Call updateItemCost(varParentItemID)

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
'Description:   This function validates that the appropriate arguments were
'               passed.  It also passes back the Item ID and the Component
'               Item ID for the parent part number and the component part
'               number.
'
'Arguments:     r_varParentItemID       -   Returning Parent Item ID
'               r_varComponentItemID    -   Returning Component Item ID
'Returns:
'               TRUE                    -   All arguments are valid
'               FALSE                   -   Not all arguments are valid
'
'//////////////////////////////////////////////////////////////////////////////

Private Function ValidArguments(ByRef r_varParentItemID As Variant, _
                                ByRef r_varComponentItemID As Variant, _
                                ByVal v_strFunction) As Boolean
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If
    
    Dim strRoutine_nm       As String, _
        strSql              As String, _
        blnDone             As Boolean, _
        rstBill             As ADODB.RecordSet, _
        clsBillOfMaterials  As Bill_of_Materials, _
        objItem             As Object
    
    strRoutine_nm = "bill_of_material.cls.ValidArguments"
                                        ' The parent part number is required
    If Len(Trim(ParentPartNbr)) = 0 Or ParentPartNbr = Null Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0100"
        Let o_error.p_message_id = 2111
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Parent Part Number is requires a value to be passed"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
        Exit Function
    End If
                                        
    Set objItem = CreateObject("converge_item.item")
    Set objItem.o_error = o_error
    Set objItem.o_trans_data = o_trans_data
    objItem.p_item_id = Null
    objItem.p_item_nbr = ParentPartNbr
    objItem.getItem
        
    If IsNull(objItem.p_item_id) = True Then
        With o_error
            .p_type_cd = "E"
            .p_err_cd = "0100"
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 1050
        End With
        ValidArguments = False
        Exit Function
    End If
        
    r_varParentItemID = objItem.p_item_id
    
    If v_strFunction = "AddToBill" Then
        If objItem.p_type_cd = "AS" Then
        Else
            With o_error
                .p_type_cd = "E"
                .p_err_cd = "0110"
                .p_desc = Err.Description
                .p_routine_nm = strRoutine_nm
                .p_message_id = 3303
            End With
            ValidArguments = False
            Exit Function
        End If
    End If

                                        ' Stop here if the function is
                                        ' DeleteComponents
    If v_strFunction = "DeleteComponents" Then
        ValidArguments = True
        Exit Function
    End If
                                        ' The component part number is required
    If Len(Trim(ComponentPartNbr)) = 0 Or ComponentPartNbr = Null Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0150"
        Let o_error.p_message_id = 2112
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Component Part Number requires a value to be passed"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
        Exit Function
    End If
                                        ' Pass back the Component Item ID
    r_varComponentItemID = fGetItemId(ComponentPartNbr, o_error, o_trans_data)
    
    If o_error.p_type_cd = "F" Or o_error.p_type_cd = "E" Then
        ValidArguments = False
        Exit Function
    End If
    
    If IsNull(r_varComponentItemID) = True Then
        With o_error
            .p_type_cd = "F"
            .p_err_cd = "0100"
            .p_desc = Err.Description
            .p_routine_nm = strRoutine_nm
            .p_message_id = 2112
        End With
        Exit Function
    End If
                                        
    If r_varComponentItemID = r_varParentItemID Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0175"
        Let o_error.p_message_id = 2116
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Component Part Number and Parent Part " & _
                        "Number cannot have the same value"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
        Exit Function
    End If
                                        ' Stop here if the function is
                                        ' RemoveComponent
    If v_strFunction = "RemoveComponent" Then
        ValidArguments = True
        Exit Function
    End If
                                        ' If the item qty is not passed, exit
                                        ' the routine.
    If ItemQty = 0 Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0200"
        Let o_error.p_message_id = 2114
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Item Quantity requires a value to be passed"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
        Exit Function
    End If
                                        ' Stop here if the function is ModifyBill
    If v_strFunction = "ModifyBill" Then
        ValidArguments = True
        Exit Function
    End If
                                        ' Check to see if the component isn't
                                        ' already assigned to the parent.
    strSql = "SELECT count(*) " & _
             "FROM bill_of_material " & _
             "WHERE item_id = " & CStr(r_varParentItemID) & " AND " & _
             "      comp_item_id = " & CStr(r_varComponentItemID)
             
    Call o_trans_data.OpenRecordset(rstBill, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstBill.Close
        Set rstBill = Nothing
        ValidArguments = False
        Exit Function
    End If
        
    If rstBill(0) > 0 Then
        rstBill.Close
        Set rstBill = Nothing
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0300"
        Let o_error.p_message_id = 2113
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Component part already exists for parent part"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
        Exit Function
    End If
                                        ' Close recordset
    rstBill.Close
    Set rstBill = Nothing
                                        ' Check to ensure that no parent part
                                        ' number equals the new component part
                                        ' number for add functions.
    If v_strFunction = "AddToBill" Then
                                        ' Create bill of materials object
        Set clsBillOfMaterials = New Bill_of_Materials
        Set clsBillOfMaterials.o_error = o_error
        Set clsBillOfMaterials.o_trans_data = o_trans_data
        clsBillOfMaterials.ParentPartNbr = ComponentPartNbr
        
        Call clsBillOfMaterials.BOMExplosion_MultiLevel
                                        ' Process BOM recordset
        Set rstBill = clsBillOfMaterials.BOM
        blnDone = False
        Do While rstBill.EOF = False And blnDone = False
            If rstBill("strComponentPartNbr") = ParentPartNbr Then
                blnDone = True
            End If
            rstBill.MoveNext
        Loop
                                        ' Close recordset
        rstBill.Close
        Set rstBill = Nothing
    End If
                                        ' If the loop ended in error, exit function
    If blnDone = True Then
        Let o_error.p_type_cd = "E"
        Let o_error.p_err_cd = "0400"
        Let o_error.p_message_id = 2127
        Let o_error.p_routine_nm = strRoutine_nm
        Let o_error.p_desc = "Cannot have a component part that is also a " & _
                             "parent part number within the same assembly"
        Let o_error.p_sqlstate_cd = ""
        ValidArguments = False
    End If
                                        ' All arguments are valid
    ValidArguments = True

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
'Description:   This procedure removes a component from the bill of material
'               for a parent part.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub RemoveComponent()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strSql                          As String, _
        strRoutine_nm                   As String, _
        varParentItemID                 As Variant, _
        varComponentItemID              As Variant, _
        dblUpdatedqty                   As Double, _
        rstBill                         As ADODB.RecordSet
    
    strRoutine_nm = "Bill_of_Material.cls.RemoveBill"
    
    If Not ValidArguments(varParentItemID, varComponentItemID, "RemoveComponent") Then
        Exit Sub
    End If
            
    strSql = "DELETE FROM bill_of_material " & _
             "WHERE item_id = " & CStr(varParentItemID) & " AND " & _
             "      comp_item_id = " & CStr(varComponentItemID)
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, o_error)
    
    Call updateItemCost(varParentItemID)
    
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
'Description:   This procedure deletes all components, one level down, from
'               the bill of material for a parent part.
'
'Arguments:     none
'
'//////////////////////////////////////////////////////////////////////////////

Public Sub DeleteComponents()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim strSql                          As String, _
        strRoutine_nm                   As String, _
        varParentItemID                 As Variant, _
        varComponentItemID              As Variant, _
        dblUpdatedqty                   As Double, _
        rstBill                         As ADODB.RecordSet
    
    strRoutine_nm = "Bill_of_Material.cls.DeleteComponents"
    
    If Not ValidArguments(varParentItemID, varComponentItemID, "DeleteComponents") Then
        Exit Sub
    End If
                                            ' Scan components of parent
    strSql = "SELECT * " & _
             "FROM bill_of_material " & _
             "WHERE item_id = " & CStr(varParentItemID)
    
    Call o_trans_data.OpenRecordset(rstBill, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        rstBill.Close
        Set rstBill = Nothing
        Exit Sub
    End If
    
    Do While Not rstBill.EOF
                                        ' Remove component from parent
        strSql = "DELETE FROM bill_of_material " & _
                 "WHERE item_id = " & CStr(varParentItemID) & " AND " & _
                 "      comp_item_id = " & CStr(rstBill("comp_item_id"))
        Call o_trans_data.ExecuteDatabase(strSql, dblUpdatedqty, o_error)
        
        rstBill.MoveNext
    Loop
    rstBill.Close
    Set rstBill = Nothing

    Call updateItemCost(varParentItemID)
    
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


Private Sub updateItemCost(ByVal v_lngItemId)
    Dim item As Object
    Set item = CreateObject("converge_item.Item")
    Set item.o_error = Me.o_error
    Set item.o_trans_data = Me.o_trans_data
    item.p_item_id = v_lngItemId
    Call item.getItem
    item.Change
End Sub


