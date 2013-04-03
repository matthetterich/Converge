VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "supplier_lots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As adodb.Recordset
Private mvarp_item_id As Variant
Private mvarp_supplier_id As Variant
Private mvarp_supplier_item_nbr As Variant
Private mvarp_supplier_lot_id As Variant
'

Public Property Let p_supplier_lot_id(ByVal vData As Variant)
    mvarp_supplier_lot_id = vData
End Property
Public Property Get p_supplier_lot_id() As Variant
    p_supplier_lot_id = mvarp_supplier_lot_id
End Property
Public Property Let p_supplier_item_nbr(ByVal vData As Variant)
    mvarp_supplier_item_nbr = vData
End Property
Public Property Get p_supplier_item_nbr() As Variant
    p_supplier_item_nbr = mvarp_supplier_item_nbr
End Property
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
    p_supplier_id = mvarp_supplier_id
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
     p_item_id = mvarp_item_id
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
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
Public Sub GetSupplierLots()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    
    ' Define
    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
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
                
    If IsNull(p_item_id) = True Or p_item_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If IsNull(p_supplier_id) = True Or p_supplier_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.supplier_id", p_supplier_id, "N", "=")
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
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If

    Set o_record_set = objRecordset
    Call Err.Clear

End Sub

