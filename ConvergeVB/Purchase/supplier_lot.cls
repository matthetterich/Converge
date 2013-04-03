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
Private mvarp_supplier_lot_nbr As Variant 'local copy
Private mvarp_supplier_id As Variant 'local copy
Private mvarp_item_id As Variant 'local copy
Private mvarsupplier_item_nbr As Variant 'local copy
Private mvarsupplier_item_desc As Variant 'local copy
Private mvaro_error As Object 'local copy
Private mvaro_dbconnection As Object 'local copy
Private mvaro_trans_data As Object 'local copy
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Property Let supplier_item_desc(ByVal vData As Variant)
    mvarsupplier_item_desc = vData
End Property
Public Property Set supplier_item_desc(ByVal vData As Variant)
    Set mvarsupplier_item_desc = vData
End Property
Public Property Get supplier_item_desc() As Variant
    If IsObject(mvarsupplier_item_desc) Then
        Set supplier_item_desc = mvarsupplier_item_desc
    Else
        supplier_item_desc = mvarsupplier_item_desc
    End If
End Property
Public Property Let supplier_item_nbr(ByVal vData As Variant)
    mvarsupplier_item_nbr = vData
End Property
Public Property Set supplier_item_nbr(ByVal vData As Variant)
    Set mvarsupplier_item_nbr = vData
End Property
Public Property Get supplier_item_nbr() As Variant
    If IsObject(mvarsupplier_item_nbr) Then
        Set supplier_item_nbr = mvarsupplier_item_nbr
    Else
        supplier_item_nbr = mvarsupplier_item_nbr
    End If
End Property
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
Public Property Let p_supplier_id(ByVal vData As Variant)
    mvarp_supplier_id = vData
End Property
Public Property Set p_supplier_id(ByVal vData As Variant)
    Set mvarp_supplier_id = vData
End Property
Public Property Get p_supplier_id() As Variant
    If IsObject(mvarp_supplier_id) Then
        Set p_supplier_id = mvarp_supplier_id
    Else
        p_supplier_id = mvarp_supplier_id
    End If
End Property
Public Property Let p_supplier_lot_nbr(ByVal vData As Variant)
    mvarp_supplier_lot_nbr = vData
End Property
Public Property Set p_supplier_lot_nbr(ByVal vData As Variant)
    Set mvarp_supplier_lot_nbr = vData
End Property
Public Property Get p_supplier_lot_nbr() As Variant
    If IsObject(mvarp_supplier_lot_nbr) Then
        Set p_supplier_lot_nbr = mvarp_supplier_lot_nbr
    Else
        p_supplier_lot_nbr = mvarp_supplier_lot_nbr
    End If
End Property
