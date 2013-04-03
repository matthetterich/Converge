VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "associates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"associate"
Attribute VB_Ext_KEY = "Member0" ,"associate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit
Private mCol As Collection
Private mvaro_error As Variant
Private mvaro_trans_data As Variant
Public Property Let o_trans_data(ByVal vData As Variant)
    mvaro_trans_data = vData
End Property
Public Property Set o_trans_data(ByVal vData As Variant)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Variant
    If IsObject(mvaro_trans_data) Then
        Set o_trans_data = mvaro_trans_data
    Else
        o_trans_data = mvaro_trans_data
    End If
End Property
Public Property Let o_error(ByVal vData As Variant)
    mvaro_error = vData
End Property
Public Property Set o_error(ByVal vData As Variant)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Variant
    If IsObject(mvaro_error) Then
        Set o_error = mvaro_error
    Else
        o_error = mvaro_error
    End If
End Property
Public Sub getAssociates()
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "associates.cls.GetAssociates"
    Dim objRecordset As adodb.Recordset
    Dim objAssociate As associate
    Dim strSql As String
    Dim key As Integer
    
    strSql = "select assoc_id from associate "
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If
    
    key = 0
    Do Until objRecordset.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        Set objAssociate = New associate
        Set objAssociate.o_error = o_error
        Set objAssociate.o_trans_data = o_trans_data
        objAssociate.p_assoc_id = objRecordset("assoc_id")
        Call objAssociate.GetAssociate
        key = key + 1
        Call Add(objAssociate, CStr(key))
        Set objAssociate = Nothing
        Call objRecordset.MoveNext
    Loop

End Sub
Public Function Add(o_associate As Object, Optional sKey As String) As associate
    
    If Len(sKey) = 0 Then
        mCol.Add o_associate
    Else
        mCol.Add o_associate, sKey
    End If

    Set Add = o_associate

End Function
Public Property Get Item(vntIndexKey As Variant) As associate
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
