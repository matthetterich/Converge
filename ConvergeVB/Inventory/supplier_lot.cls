VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "supplier_lot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_supplier_id As Variant
Private mvarp_supplier_nbr As Variant
Private mvarp_item_nbr As Variant
Private mvarp_item_id As Variant
Private mvarp_supplier_item_nbr As Variant
Private mvarp_supplier_item_desc As Variant
Private mvarp_supplier_lot_id As Variant
Private mvarp_old_supplier_lot_id As Variant
Private mvaro_record_set As ADODB.Recordset
'
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property
'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_lot_id 'Let' and 'Get' are used to
'               for the new supplier lot id
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_lot_id(ByVal vData As Variant)
    mvarp_supplier_lot_id = vData
End Property
Public Property Get p_supplier_lot_id() As Variant
        p_supplier_lot_id = mvarp_supplier_lot_id
End Property
'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_nbr 'Let' and 'Get' are used to
'               for the new supplier lot id
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_nbr(ByVal vData As Variant)
    mvarp_supplier_nbr = vData
End Property
Public Property Get p_supplier_nbr() As Variant
        p_supplier_nbr = mvarp_supplier_nbr
End Property
'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_item_nbr 'Let' and'Get' are used to
'               for the new item number '
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_item_nbr(ByVal vData As Variant)
    mvarp_item_nbr = vData
End Property
Public Property Get p_item_nbr() As Variant
        p_item_nbr = mvarp_item_nbr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_lot_id 'Let' and 'Get' are used to
'               for the old supplier lot id
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_old_supplier_lot_id(ByVal vData As Variant)
    mvarp_old_supplier_lot_id = vData
End Property
Public Property Get p_old_supplier_lot_id() As Variant
        p_old_supplier_lot_id = mvarp_old_supplier_lot_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_item_desc 'Let' and 'Get' are used to
'               for the supplier item description.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_item_desc(ByVal vData As Variant)
    mvarp_supplier_item_desc = vData
End Property
Public Property Get p_supplier_item_desc() As Variant
        p_supplier_item_desc = mvarp_supplier_item_desc
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_item_nbr 'Let' and 'Get' are used to
'               for the supplier item number.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_item_nbr(ByVal vData As Variant)
    mvarp_supplier_item_nbr = vData
End Property
Public Property Get p_supplier_item_nbr() As Variant
        p_supplier_item_nbr = mvarp_supplier_item_nbr
   
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_item_id 'Let' and 'Get' are used to
'               for the Item Identifier.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The p_supplier_id 'Let' and 'Get' are used to
'               for the supplier Identifier.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
        p_supplier_id = mvarp_supplier_id
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The 'Set' and 'Get' are used to
'               for the common object declarations.
'
'//////////////////////////////////////////////////////////////////////////////
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

'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to get the supplier lot - using p_supplier_lot_id
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetSupplierLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_lot.cls.GetSupplierLot"
  
    Call SelectSupplierLot
    
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
'Description:   Function to Add the supplier lot
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    strRoutine_nm = "Supplier_Lot.cls.AddSupplierLot"
    
    ' Check to see if the supplier lot already exists.
    strSql = "SELECT * " & _
            " FROM supplier_lot " & _
            "WHERE " & fWhereVariable("supplier_lot_id", p_supplier_lot_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If
    
    If objrecordset.EOF = True Then
        Call AddLot
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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
'Description:   Function to change the supplier lot record
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_lot.cls.Change"
    
    Call Validate("change")
    Call IssueUpdateData
    
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
'Description:  Function to delete the supplier lot record
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub Delete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    ' Initialize
    strRoutine_nm = "supplier_lot.cls.Delete"
    
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

'//////////////////////////////////////////////////////////////////////////////
'Description:  Called from purchase order receipt and adds the lot
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub AddPORecSupplierLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    strRoutine_nm = "Supplier_Lot.cls.AddPORecSupplierLot"
    
    ' Check to see if the supplier lot already exists.
    strSql = "SELECT * " & _
            " FROM supplier_lot " & _
            "WHERE " & fWhereVariable("supplier_lot_id", p_supplier_lot_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If
    
    If objrecordset.EOF = True Then
        Call AddLot
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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
'Description:   Function to Add the supplier lot
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub AddLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "supplier_lot.cls.Add"
    
    Call Validate("add")
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


'//////////////////////////////////////////////////////////////////////////////
'Description:  Function to perform the delete of supplier lot
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueDeleteData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "supplier_lot.cls.IssueDeleteData"

    strSql = "DELETE FROM supplier_lot " & _
            " WHERE " & fWhereVariable("supplier_lot_id", p_supplier_lot_id, "S", "=")
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:  Function to Validate the information for supplier lot
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub Validate(sFunction_cd)

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim varItem_nbr
    Dim varSupplier_nbr
    
    strRoutine_nm = "supplier_lot.cls.Validate"

    If sFunction_cd = "add" Then
        If fValuePresent(p_supplier_lot_id) = False Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1100"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2215
        End If
        
        varItem_nbr = fGetItemNbr(p_item_id, o_error, o_trans_data)
        If IsNull(varItem_nbr) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1110"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2090
        End If
        
        varSupplier_nbr = fGetSupplierNbr(p_supplier_id, o_error, o_trans_data)
        If IsNull(varSupplier_nbr) = True Then
            o_error.p_type_cd = "E"
            o_error.p_err_cd = "1110"
            o_error.p_routine_nm = strRoutine_nm
            o_error.p_message_id = 2100
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




'//////////////////////////////////////////////////////////////////////////////
'Description:  Function to insert data into supplier lot table
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueInsertData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "supplier_lot.cls.IssueInsert"
    
    strSql = "insert into supplier_lot ( "
    strSql = strSql & "  supplier_lot_id"
    strSql = strSql & ", Supplier_id"
    strSql = strSql & ", item_id"
    strSql = strSql & ", supplier_item_nbr"
    strSql = strSql & ", supplier_item_desc"
    strSql = strSql & ", updated_assoc_id"
    strSql = strSql & ", updated_ts"
    strSql = strSql & ", entered_assoc_id"
    strSql = strSql & ", entered_ts"
    strSql = strSql & " ) "
    strSql = strSql & "values ("
    strSql = strSql & " " & fInsertVariable(p_supplier_lot_id, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_id, "N")
    strSql = strSql & "," & fInsertVariable(p_item_id, "N")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_nbr, "S")
    strSql = strSql & "," & fInsertVariable(p_supplier_item_desc, "S")
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
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub


'//////////////////////////////////////////////////////////////////////////////
'Description:  Function to insert data into supplier lot table
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub IssueUpdateData()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "supplier_lot.cls.IssueUpdateData"
    
    strSql = "UPDATE supplier_lot SET " & _
             "       " & fUpdateVariable("supplier_item_nbr", p_supplier_item_nbr, "S", "=") & _
             ",      " & fUpdateVariable("supplier_item_desc", p_supplier_item_desc, "S", "=") & _
             " WHERE " & fWhereVariable("supplier_lot_id", p_supplier_lot_id, "S", "=") & _
             " and   " & fWhereVariable("supplier_id", p_supplier_id, "N", "=")
    
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


'//////////////////////////////////////////////////////////////////////////////
'Description:  Function to insert data into supplier lot table
'  Arguments:  None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub SelectSupplierLot()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    ' Initialize
    strRoutine_nm = "Supplier_Lot.cls.SelectSupplierLot"
    
    strSql = "SELECT * " & _
            " FROM supplier_lot " & _
            "WHERE " & fWhereVariable("supplier_lot_id", p_supplier_lot_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If
    
    If objrecordset.EOF = True Then
        p_supplier_lot_id = Null
    Else
        p_supplier_lot_id = objrecordset("supplier_lot_id")
        p_item_id = objrecordset("item_id")
        p_supplier_item_nbr = objrecordset("supplier_item_nbr")
        p_supplier_id = objrecordset("supplier_id")
        p_supplier_item_desc = objrecordset("supplier_item_desc")
    End If
    
    objrecordset.Close
    Set objrecordset = Nothing
    
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


Public Sub GetSupplierLots()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objrecordset As ADODB.Recordset
    
    ' Initialize
    strRoutine_nm = "inventory_adjs.cls.getSupplierLots"
        
    strSql = "SELECT " & _
                    " a.supplier_lot_id, " & _
                    " a.supplier_item_nbr, " & _
                    " b.item_nbr, " & _
                    " b.item_desc," & _
                    " c.supplier_nbr " & _
               "  FROM supplier_lot a, item b, supplier c " & _
               " WHERE a.item_id = b.item_id " & _
                "   AND a.supplier_id = c.supplier_id "
                
    If IsNull(p_item_nbr) = True Or p_item_nbr = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("b.item_nbr", p_item_nbr, "S", "=")
    End If
    
    If IsNull(p_supplier_nbr) = True Or p_supplier_nbr = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("c.supplier_nbr", p_supplier_nbr, "S", "=")
    End If
    
    If IsNull(p_supplier_lot_id) = True Or p_supplier_lot_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_lot_id", p_supplier_lot_id, "S", "=")
    End If
    
    If IsNull(p_supplier_item_nbr) = True Or p_supplier_item_nbr = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_item_nbr", p_supplier_item_nbr, "S", "=")
    End If
    
    strSql = strSql & " ORDER BY a.supplier_lot_id "
    
    Call o_trans_data.OpenRecordset(objrecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objrecordset.Close
        Set objrecordset = Nothing
        Exit Sub
    End If

    Set o_record_set = objrecordset

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


