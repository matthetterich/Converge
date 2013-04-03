VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"customer"
Attribute VB_Ext_KEY = "Member0" ,"customer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit
Private mCol As Collection
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_cust_grp_id As Variant
Public Property Let p_cust_grp_id(ByVal vData As Variant)
    mvarp_cust_grp_id = vData
End Property
Public Property Set p_cust_grp_id(ByVal vData As Variant)
    Set mvarp_cust_grp_id = vData
End Property
Public Property Get p_cust_grp_id() As Variant
    If IsObject(mvarp_cust_grp_id) Then
        Set p_cust_grp_id = mvarp_cust_grp_id
    Else
        p_cust_grp_id = mvarp_cust_grp_id
    End If
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
Public Function Add(o_customer As Object, Optional sKey As String) As customer
    
    If Len(sKey) = 0 Then
        mCol.Add o_customer
    Else
        mCol.Add o_customer, sKey
    End If

    Set Add = o_customer

End Function

Public Sub getCustGrpCustomers()
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "customers.cls.getCustGrpCustomers"
    Dim objRecordset As adodb.Recordset
    Dim objCustomer As customer
    Dim intUpdatedQty As Integer
    Dim strSql As String
    Dim key As Integer
    
    strSql = "select * from customer "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("cust_grp_id", p_cust_grp_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.clear
        Exit Sub
    End If
    
    key = 0
    Do Until objRecordset.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        Set objCustomer = New customer
        Set objCustomer.o_error = o_error
        Set objCustomer.o_trans_data = o_trans_data
        objCustomer.p_cust_id = objRecordset("cust_id")
        Call objCustomer.GetCustomer
        key = key + 1
        Call Add(objCustomer, CStr(key))
        Set objCustomer = Nothing
        Call objRecordset.MoveNext
    Loop

End Sub

Public Sub getCustomers()
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "customers.cls.getCustomers"
    Dim objRecordset As adodb.Recordset
    Dim objCustomer As customer
    Dim intUpdatedQty As Integer
    Dim strSql As String
    Dim key As Integer
    
    strSql = "select cust_id from customer "
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.clear
        Exit Sub
    End If
    
    key = 0
    Do Until objRecordset.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        Set objCustomer = New customer
        Set objCustomer.o_error = o_error
        Set objCustomer.o_trans_data = o_trans_data
        objCustomer.p_cust_id = objRecordset("cust_id")
        Call objCustomer.GetCustomer
        key = key + 1
        Call Add(objCustomer, CStr(key))
        Set objCustomer = Nothing
        Call objRecordset.MoveNext
    Loop

End Sub
Public Property Get Item(vntIndexKey As Variant) As customer
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
