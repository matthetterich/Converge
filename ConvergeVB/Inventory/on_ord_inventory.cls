VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "on_ord_inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvaro_trans_data As Object 'local copy
Private mvaro_dbconnection As Object 'local copy
Private mvaro_error As Object 'local copy
Private mvarp_item_id As Variant 'local copy
Private mvarp_po_id As Variant 'local copy
Private mvarp_po_line_nbr As Variant 'local copy
Private mvarp_assem_ord_id As Variant 'local copy
Private mvarp_on_ord_qty As Variant 'local copy
Public Property Let p_on_ord_qty(ByVal vData As Variant)
    mvarp_on_ord_qty = vData
End Property
Public Property Set p_on_ord_qty(ByVal vData As Variant)
    Set mvarp_on_ord_qty = vData
End Property
Public Property Get p_on_ord_qty() As Variant
    If IsObject(mvarp_on_ord_qty) Then
        Set p_on_ord_qty = mvarp_on_ord_qty
    Else
        p_on_ord_qty = mvarp_on_ord_qty
    End If
End Property
Public Property Let p_assem_ord_id(ByVal vData As Variant)
    mvarp_assem_ord_id = vData
End Property
Public Property Set p_assem_ord_id(ByVal vData As Variant)
    Set mvarp_assem_ord_id = vData
End Property
Public Property Get p_assem_ord_id() As Variant
    If IsObject(mvarp_assem_ord_id) Then
        Set p_assem_ord_id = mvarp_assem_ord_id
    Else
        p_assem_ord_id = mvarp_assem_ord_id
    End If
End Property
Public Property Let p_po_line_nbr(ByVal vData As Variant)
    mvarp_po_line_nbr = vData
End Property
Public Property Set p_po_line_nbr(ByVal vData As Variant)
    Set mvarp_po_line_nbr = vData
End Property
Public Property Get p_po_line_nbr() As Variant
    If IsObject(mvarp_po_line_nbr) Then
        Set p_po_line_nbr = mvarp_po_line_nbr
    Else
        p_po_line_nbr = mvarp_po_line_nbr
    End If
End Property
Public Property Let p_po_id(ByVal vData As Variant)
    mvarp_po_id = vData
End Property
Public Property Set p_po_id(ByVal vData As Variant)
    Set mvarp_po_id = vData
End Property
Public Property Get p_po_id() As Variant
    If IsObject(mvarp_po_id) Then
        Set p_po_id = mvarp_po_id
    Else
        p_po_id = mvarp_po_id
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
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
