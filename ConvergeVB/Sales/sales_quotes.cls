VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sales_quotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"sales_quote"
Attribute VB_Ext_KEY = "Member0" ,"sales_quote"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mCol As Collection
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvarp_cust_id As Variant
Public Property Let p_cust_id(ByVal vData As Variant)
    mvarp_cust_id = vData
End Property
Public Property Set p_cust_id(ByVal vData As Variant)
    Set mvarp_cust_id = vData
End Property
Public Property Get p_cust_id() As Variant
    If IsObject(mvarp_cust_id) Then
        Set p_cust_id = mvarp_cust_id
    Else
        p_cust_id = mvarp_cust_id
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
Public Function Add(o_sales_quote, Optional sKey As String) As sales_quote
    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add o_sales_quote
    Else
        mCol.Add o_sales_quote, sKey
    End If
    
    'return the object created
    Set Add = o_sales_quote

End Function
Public Property Get Item(vntIndexKey As Variant) As sales_quote
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
Public Sub getCustSalesQuotes()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "Sales_Quotes.cls.GetCustSalesQuotes"
    Dim objRecordset As adodb.Recordset
    Dim objSales_Quote As sales_quote
    Dim intUpdatedQty As Integer
    Dim strSql As String
    Dim key As Integer
    
    strSql = "SELECT sales_quote_id FROM sales_quote " & _
                " Order by sales_quote_id "
              
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.clear
        Exit Sub
    End If
    
    key = 0
    Do Until objRecordset.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        Set objSales_Quote = New sales_quote
        Set objSales_Quote.o_error = o_error
        Set objSales_Quote.o_trans_data = o_trans_data
        objSales_Quote.p_sales_quote_id = objRecordset("sales_quote_id")
        Call objSales_Quote.GetSalesQuote
        key = key + 1
        Call Add(objSales_Quote, CStr(key))
        Set objSales_Quote = Nothing
        Call objRecordset.MoveNext
    Loop

    Err.clear
    
End Sub
