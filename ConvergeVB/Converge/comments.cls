VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Comment"
Attribute VB_Ext_KEY = "Member0" ,"Comment"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Item"
Option Explicit

Private mCol As Collection
Private mvaro_trans_data As Variant 'local copy
Private mvaro_error As Variant 'local copy
Private mvarp_obj_nm As Variant 'local copy
Private mvarp_obj_id As Variant 'local copy

Public Property Let p_obj_id(ByVal vData As Variant)
    mvarp_obj_id = vData
End Property
Public Property Get p_obj_id() As Variant
        p_obj_id = mvarp_obj_id
End Property


Public Property Let p_obj_nm(ByVal vData As Variant)
    mvarp_obj_nm = vData
End Property
Public Property Get p_obj_nm() As Variant
        p_obj_nm = mvarp_obj_nm
End Property

Public Property Set o_error(ByVal vData As Variant)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Variant
    Set o_error = mvaro_error
End Property


Public Property Set o_trans_data(ByVal vData As Variant)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Variant
    Set o_trans_data = mvaro_trans_data
End Property

Public Sub getComments()
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim strRoutine_nm As String
    strRoutine_nm = "comments.cls.GetComments"
    Dim objRecordset As Object
    Dim objComment As Comment
    Dim strSql As String
    Dim key As Integer
    
    strSql = "select * from comment "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("obj_nm", p_obj_nm, "S", "=")
    strSql = strSql & "and" & fWhereVariable("obj_id", p_obj_id, "N", "=")
    strSql = strSql & " order by entered_ts"
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Err.Clear
        Exit Sub
    End If

    key = 0
    Do Until objRecordset.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        Set objComment = New Comment
        Set objComment.o_error = o_error
        Set objComment.o_trans_data = o_trans_data
        objComment.p_comment_id = objRecordset("comment_id")
        Call objComment.GetComment
        key = key + 1
        Call Add(objComment, CStr(key))
        Set objComment = Nothing
        Call objRecordset.MoveNext
    Loop

    Set objRecordset = Nothing
    Call Err.Clear

End Sub



Public Function Add(o_comment, Optional sKey As String) As Comment
    
    If Len(sKey) = 0 Then
        mCol.Add o_comment
    Else
        mCol.Add o_comment, sKey
    End If


    Set Add = o_comment

End Function

Public Property Get Item(vntIndexKey As Variant) As Comment
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

