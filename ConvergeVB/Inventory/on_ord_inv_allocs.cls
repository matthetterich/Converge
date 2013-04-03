VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "on_ord_inv_allocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_item_id As Variant
Private mvarp_inv_loc_id As Variant
Private mvaro_record_set As adodb.Recordset
'

Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
End Property
Public Property Let p_inv_loc_id(ByVal vData As Variant)
    mvarp_inv_loc_id = vData
End Property
Public Property Get p_inv_loc_id() As Variant
        p_inv_loc_id = mvarp_inv_loc_id
End Property
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
        p_item_id = mvarp_item_id
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
'//////////////////////////////////////////////////////////////////////////////
'Description:   Gets the inventory allocation class
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetOnOrderInventoryAllocs()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    Dim objRecordset As adodb.Recordset
    Dim strSql As String
    strRoutine_nm = "on_ord_inv_allocs.cls.getOnOrderInventoryallocs"
    
    strSql = "SELECT " & _
                " a.on_ord_inv_alloc_id, " & _
                " a.alloc_qty, " & _
                " a.mat_req_id, " & _
                " a.inv_loc_id, " & _
                " b.assem_ord_id, " & _
                " b.sales_ord_id, " & _
                " c.item_nbr, " & _
                " c.item_desc " & _
                " FROM (on_ord_inv_alloc a inner join mat_req b on " & _
                " a.mat_req_id = b.mat_req_id ) inner join item c on b.item_id = c.item_id" & _
                " WHERE  1 = 1 "
                
    If IsNull(p_item_id) = True Or p_item_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.item_id", p_item_id, "N", "=")
    End If
    
    If IsNull(p_inv_loc_id) = True Or p_inv_loc_id = "" Then
    Else
       strSql = strSql & "  AND " & fWhereVariable("a.inv_loc_id", p_inv_loc_id, "S", "=")
    End If
    
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
