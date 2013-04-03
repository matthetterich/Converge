Attribute VB_Name = "modUtilites"
Option Explicit

Global objUtilities As Object

Public Function fInsertVariable(sHost_var, sData_type_cd) As String
    
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    objUtilities.p_host_var = sHost_var
    objUtilities.p_data_type_cd = sData_type_cd
    objUtilities.p_format_str = ""
    Call objUtilities.FormatInsertVariable
    fInsertVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function


Public Function fUpdateVariable(sUpdate_var_nm, sHost_var, sData_type_cd, sOperation_cd) As String
     
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    objUtilities.p_host_var = sHost_var
    objUtilities.p_data_type_cd = sData_type_cd
    objUtilities.p_host_var_nm = sUpdate_var_nm
    objUtilities.p_Operation_cd = sOperation_cd
    objUtilities.p_format_str = ""
    Call objUtilities.FormatUpdateVariable
    fUpdateVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function




Public Function fWhereVariable(sUpdate_var_nm, sHost_var, sData_type_cd, sOperation_cd) As String
     
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    objUtilities.p_host_var = sHost_var
    objUtilities.p_data_type_cd = sData_type_cd
    objUtilities.p_host_var_nm = sUpdate_var_nm
    objUtilities.p_Operation_cd = sOperation_cd
    objUtilities.p_format_str = ""
    Call objUtilities.FormatWhereVariable
    fWhereVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function
Public Function fGetDesc(ByVal v_strName As String, ByVal v_strCode As Variant, ByRef o_error As Object, ByRef o_trans_data As Object) As String
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm
    strRoutine_nm = "Utilities.bas.fGetDesc"

    If IsNull(v_strCode) = True Then
        fGetDesc = Empty
        Exit Function
    End If
   
    Dim varCodeDesc As Variant
    Dim intCount As Integer
    
    strRoutine_nm = "transaction.cls.CheckCodeDesc"

    varCodeDesc = o_trans_data.p_CodeDesc
    
    fGetDesc = Empty
    
    For intCount = 0 To UBound(varCodeDesc, 2)
        If LCase(varCodeDesc(2, intCount)) = LCase(v_strCode) Then
            If LCase(varCodeDesc(0, intCount)) = LCase(v_strName) Then
                fGetDesc = varCodeDesc(3, intCount)
                Exit For
            End If
        End If
    Next intCount
 
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
    Err.clear
End Function

Public Function fValuePresent(sField)

    If IsNull(sField) = True Then
        fValuePresent = False
        Exit Function
    End If
    
    If IsEmpty(sField) = True Then
        fValuePresent = False
        Exit Function
    End If
    
    If Len(sField) > 0 Then
    Else
        fValuePresent = False
        Exit Function
    End If
    
    If sField = "" Then
        fValuePresent = False
        Exit Function
    End If
    
    fValuePresent = True
    
End Function


Public Function fGetAssocNbr(ByVal varAssoc_id As Variant, ByRef o_error As Object, o_trans_data)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Utilities.bas.fGetAssocNbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varAssoc_id) = False Then
        fGetAssocNbr = Null
        Exit Function
    End If
    
    strSql = "SELECT assoc_nbr      "
    strSql = strSql & "FROM associate "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("assoc_id", varAssoc_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetAssocNbr = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetAssocNbr = objRecordset("assoc_nbr")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function


Public Function fGetAssocId(ByVal varAssoc_nbr As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "Utilities.bas.fGetAssocId"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varAssoc_nbr) = False Then
        fGetAssocId = Null
        Exit Function
    End If
    
    strSql = "SELECT assoc_id      "
    strSql = strSql & "FROM associate "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("assoc_nbr", varAssoc_nbr, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetAssocId = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetAssocId = objRecordset("assoc_id")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function


Public Function fGetItemNbr(ByVal varItem_id As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetItemNbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varItem_id) = False Then
        fGetItemNbr = Null
        Exit Function
    End If
    
    strSql = "SELECT item_nbr FROM item " & _
               " WHERE " & fWhereVariable("item_id", varItem_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetItemNbr = Null
        Exit Function
    End If
    
    fGetItemNbr = objRecordset("item_nbr")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function


Public Function fGetItemId(ByVal varItem_nbr As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetItemId"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varItem_nbr) = False Then
        fGetItemId = Null
        Exit Function
    End If
    
    strSql = "SELECT item_id FROM item " & _
               " WHERE " & fWhereVariable("item_nbr", varItem_nbr, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetItemId = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetItemId = objRecordset("item_id")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function


Public Function fGetCustNbr(ByVal varCust_id As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetCustNbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varCust_id) = False Then
        fGetCustNbr = Null
        Exit Function
    End If
    
    strSql = "SELECT cust_nbr FROM customer " & _
               " WHERE " & fWhereVariable("cust_id", varCust_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetCustNbr = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetCustNbr = objRecordset("cust_nbr")
    
    ' Close RecordSet
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function


Public Function fGetCustId(ByVal varCust_nbr As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetCustId"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varCust_nbr) = False Then
        fGetCustId = Null
        Exit Function
    End If
    
    strSql = "SELECT cust_id FROM customer " & _
               " WHERE " & fWhereVariable("cust_nbr", varCust_nbr, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetCustId = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetCustId = objRecordset("cust_id")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function



Public Sub ValidateAssoc_nbr(ByRef varAssoc_nbr As Variant, ByRef varAssoc_id As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.base.ValidateAssoc_nbr"
    
    If fValuePresent(varAssoc_nbr) = True Then
    Else
        Exit Sub
    End If
    
    If IsNull(varAssoc_nbr) = True Then
        varAssoc_id = Null
        Exit Sub
    End If
    
    varAssoc_id = fGetAssocId(varAssoc_nbr, o_error, o_trans_data)
    
    If IsNull(varAssoc_id) = True Then
        o_error.p_type_cd = "F"
        o_error.p_err_cd = "1150"
        o_error.p_routine_nm = strRoutine_nm
        o_error.p_message_id = 2165
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
    Err.clear
End Sub




Public Function fGetSupplierId(ByVal varSupplier_nbr As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetSupplierId"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varSupplier_nbr) = False Then
        fGetSupplierId = Null
        Exit Function
    End If
    
    strSql = "SELECT supplier_id FROM supplier " & _
               " WHERE " & fWhereVariable("supplier_nbr", varSupplier_nbr, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetSupplierId = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetSupplierId = objRecordset("supplier_id")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function



Public Function fGetSupplierNbr(ByVal varSupplier_id As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "utilities.bas.fGetSupplierNbr"
    Dim objRecordset As adodb.Recordset
    Dim dblUpdateQty As Double
    Dim strSql As String
    
    If fValuePresent(varSupplier_id) = False Then
        fGetSupplierNbr = Null
        Exit Function
    End If
    
    strSql = "SELECT Supplier_nbr FROM supplier " & _
               " WHERE " & fWhereVariable("Supplier_id", varSupplier_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetSupplierNbr = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetSupplierNbr = objRecordset("Supplier_nbr")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function



Public Function fGetId(ByVal strSeq_nm As String, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "inventory_adj.cls.GetId"
    
    Dim objSequence As Object
    Set objSequence = CreateObject("converge.sequence")
    objSequence.p_seq_nm = strSeq_nm
    objSequence.p_seq_nbr = 0
    Set objSequence.o_error = o_error
    Set objSequence.o_trans_data = o_trans_data
    Call objSequence.getSeqNbr
    fGetId = objSequence.p_seq_nbr
    
    Set objSequence = Nothing

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
    Err.clear
End Function


Public Function fGetInvLocNm(ByVal v_varInv_loc_id As Variant, ByRef o_error As Object, ByRef o_trans_data As Object)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim dblUpdateQty As Double
    Dim strRoutine_nm As String, _
        objRecordset As adodb.Recordset, _
        strSql As String
    
    strRoutine_nm = "Utilities.bas.fGetInvLocNm"
    
    If fValuePresent(v_varInv_loc_id) = False Then
        fGetInvLocNm = Null
        Exit Function
    End If
    
    strSql = "SELECT inv_loc_nm      " & _
            "FROM inventory_loc " & _
            "WHERE         " & _
            "   " & fWhereVariable("inv_loc_id", v_varInv_loc_id, "S", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    If objRecordset.EOF = True Then
        fGetInvLocNm = Null
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Function
    End If
    
    fGetInvLocNm = objRecordset("inv_loc_nm")
    
    objRecordset.Close
    Set objRecordset = Nothing
    
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
    Err.clear
End Function

Public Function AddName(objName, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.AddName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    Call objName.AddName
    
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
    Err.clear
End Function

Public Function ChangeName(objName, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.ChangeName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    Call objName.ChangeName
    
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
    Err.clear
End Function

Public Function DeleteName(objName, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.DeleteName"
    
    Set objName.o_error = o_error
    Set objName.o_trans_data = o_trans_data
    Call objName.DeleteAll
  
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
    Err.clear
End Function

Public Function AddAddr(objAddress, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.AddAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    Call objAddress.AddAddress
    
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
    Err.clear
End Function

Public Function ChangeAddr(objAddress, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.ChangeAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    Call objAddress.ChangeAddress
    
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
    Err.clear
End Function

Public Function DeleteAddr(objAddress, o_trans_data, o_error)
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Function
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "sales_quote.cls.DeleteAddr"
    
    Set objAddress.o_error = o_error
    Set objAddress.o_trans_data = o_trans_data
    Call objAddress.DeleteAll
    
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
    Err.clear
End Function
