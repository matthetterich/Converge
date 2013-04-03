VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"item"
Attribute VB_Ext_KEY = "Member0" ,"Item"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private m_colItems As Collection
Private m_objTrans_data As Object 'local copy
Private m_objConnection As Object 'local copy
Private m_objError As Object 'local copy
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property

Public Property Set error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get error() As Object
    Set error = m_objError
End Property


Public Property Set o_dbConnection(ByVal vData As Object)
    Set m_objConnection = vData
End Property
Public Property Get o_dbConnection() As Object
    Set o_dbConnection = m_objConnection
End Property

Public Property Set Connection(ByVal vData As Object)
    Set m_objConnection = vData
End Property
Public Property Get Connection() As Object
    Set Connection = m_objConnection
End Property

Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property

Public Property Set trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get trans_data() As Object
    Set trans_data = m_objTrans_data
End Property


Public Function Add(objItem As Object, Optional sKey As String) As Item
    
    If Len(sKey) = 0 Then
        m_colItems.Add objItem
    Else
        m_colItems.Add objItem, sKey
    End If

    'return the object created
    Set Add = objItem


End Function

Public Sub getItems()
    
    If error.type_cd = "E" Or error.type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim sRoutine_nm As String
    Let sRoutine_nm = "Items.cls.GetItems"
    Dim objRecordset As Object
    Dim objItem As Item
    Dim sSql As String
    Dim key As Integer
    
    Set objRecordset = CreateObject("adodb.recordset")
    
    Let sSql = "select item_id from Item "
    
    Set objRecordset = o_dbConnection.Execute(sSql)
    If Err.Number = 0 Then
    Else
        Let error.type_cd = "F"
        Let error.err_cd = "0300"
        Let error.routine_nm = sRoutine_nm
        Let error.message_id = 0
        Let error.nbr = Err.Number
        Let error.Desc = Err.Description
        Let error.sqlstate_cd = o_dbConnection.Errors(0).SQLState
        Call Err.Clear
        Exit Sub
    End If

    Let key = 0
    Do Until objRecordset.EOF = True Or error.type_cd = "E" Or error.type_cd = "F"
        Set objItem = New Item
        Set objItem.o_error = error
        Set objItem.o_trans_data = trans_data
        Set objItem.o_dbConnection = Connection
        Let objItem.p_item_id = objRecordset("item_id")
        Call objItem.GetItem
        Let key = key + 1
        Call Add(objItem, CStr(key))
        Set objItem = Nothing
        Call objRecordset.MoveNext
    Loop

End Sub
Public Property Get Item(vntIndexKey As Variant) As Item
Attribute Item.VB_UserMemId = 0
  Set Item = m_colItems(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = m_colItems.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    m_colItems.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colItems.[_NewEnum]
End Property
Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub
Private Sub Class_Terminate()
    Set m_colItems = Nothing
End Sub
