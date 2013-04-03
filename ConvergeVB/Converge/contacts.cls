VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "contacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Contact"
Attribute VB_Ext_KEY = "Member0" ,"Contact"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit

Private m_colContacts As Collection
Private m_objTrans_data As Object
Private m_objError As Object
Private m_varCust_id As Variant
Private m_varSupplier_id As Variant

Public Property Let cust_id(ByVal vData As Variant)
    m_varCust_id = vData
End Property
Public Property Get cust_id() As Variant
    cust_id = m_varCust_id
End Property

Public Property Let supplier_id(ByVal vData As Variant)
    m_varSupplier_id = vData
End Property
Public Property Get supplier_id() As Variant
    supplier_id = m_varSupplier_id
End Property

Public Property Set error(ByVal vData As error)
    Set m_objError = vData
End Property
Public Property Get error() As Object
    Set error = m_objError
End Property

Public Property Set o_error(ByVal vData As error)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property

Public Property Set Trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get Trans_data() As Object
    Set Trans_data = m_objTrans_data
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property

Public Sub getContacts()
    
    If o_error.type_cd = "E" Or error.type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "contacts.cls.GetContacts"
    Dim objRecordset As adodb.Recordset
    Dim objContact As Contact
    Dim strSql As String
    Dim key As Integer
    
    ' get customer contacts if supplier id is null.  get supplier
    ' contacts if supplier id is not null.
    
    If IsNull(supplier_id) = True Then
        strSql = "select * from contact "
        strSql = strSql & "where "
        strSql = strSql & " " & fWhereVariable("cust_id", cust_id, "N", "=")
    Else
        strSql = "select * from contact "
        strSql = strSql & "where "
        strSql = strSql & " " & fWhereVariable("supplier_id", supplier_id, "N", "=")
    End If
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)

    If o_error.type_cd = "E" Or error.type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If

    key = 0
    Do Until objRecordset.EOF = True Or error.type_cd = "E" Or error.type_cd = "F"
        Set objContact = New Contact
        Set objContact.error = error
        Set objContact.Trans_data = Trans_data
        objContact.contact_id = objRecordset("contact_id")
        objContact.bypass_getcust_fl = "Y"
        Call objContact.GetContact
        key = key + 1
        Call Add(objContact, CStr(key))
        Set objContact = Nothing
        Call objRecordset.MoveNext
    Loop

    Set objRecordset = Nothing
    Call Err.Clear
End Sub
Public Function Add(Contact, Optional sKey As String) As Contact
    
    If Len(sKey) = 0 Then
        m_colContacts.Add Contact
    Else
        m_colContacts.Add Contact, sKey
    End If

    Set Add = Contact

End Function

Public Property Get Item(vntIndexKey As Variant) As Contact
Attribute Item.VB_UserMemId = 0
  Set Item = m_colContacts(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colContacts.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    m_colContacts.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colContacts.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colContacts = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colContacts = Nothing
End Sub
