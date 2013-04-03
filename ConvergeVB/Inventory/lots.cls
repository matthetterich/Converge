VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_trans_data As Object
Private mvaro_error As Object
Private mvaro_record_set As adodb.Recordset
Private mvarp_item_id As Variant
Public Property Let p_item_id(ByVal vData As Variant)
    mvarp_item_id = vData
End Property
Public Property Set p_item_id(ByVal vData As Variant)
    Set mvarp_item_id = vData
End Property
Public Property Get p_item_id() As Variant
    If IsObject(mvarp_item_id) Then
        Set p_item_id = mvarp_item_id
    Else
        p_item_id = mvarp_item_id
    End If
End Property
Public Property Set o_record_set(ByVal vData As adodb.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As adodb.Recordset
    Set o_record_set = mvaro_record_set
End Property
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
Public Sub GetLots()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.getLots"
    Dim strSql As String
    
    strSql = "select * from lot "

    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        o_record_set.Close
        Set o_record_set = Nothing
        Err.Clear
        Exit Sub
    End If

    Call Err.Clear

End Sub
Public Sub GetInventoryItembyLots()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.getLots"
    Dim strSql As String
    
    strSql = "select * from lot " & _
                "WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        o_record_set.Close
        Set o_record_set = Nothing
        Err.Clear
        Exit Sub
    End If

    Call Err.Clear

End Sub
Public Sub GetInventoryItembyLotsbyBin()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "Lot.cls.GetInventoryItembyLotsbyBin"
    Dim strSql As String
    
    strSql = "select a.lot_id, " & _
                     " a.supplier_lot_id, " & _
                     " a.unit_cost, " & _
                     " a.status_cd " & _
                     " b.bin_id " & _
                     " from lot a, inventory_bin b " & _
                "WHERE " & fWhereVariable("item_id", p_item_id, "N", "=") & _
                "  AND a.lot_id = b.lot_id " & _
                " ORDER BY a.lot_id, b.bin_id "
    

    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        o_record_set.Close
        Set o_record_set = Nothing
        Err.Clear
        Exit Sub
    End If

    Call Err.Clear

End Sub
