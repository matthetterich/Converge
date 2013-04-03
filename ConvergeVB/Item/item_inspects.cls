VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "item_inspects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_trans_data As Object 'local copy
Private mvaro_dbConnection As Object 'local copy
Private mvaro_record_set As Object 'local copy
Private mvarp_item_id As Variant 'local copy
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
Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property
Public Property Set o_dbConnection(ByVal vData As Object)
    Set mvaro_dbConnection = vData
End Property
Public Property Get o_dbConnection() As Object
    Set o_dbConnection = mvaro_dbConnection
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
Public Sub GetItemInspects()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim sRoutine_nm As String
    Let sRoutine_nm = "item_inspect_crit.cls.getItemInspects"
    Dim sSql As String
    
    Set o_record_set = CreateObject("adodb.recordset")
    
    Let sSql = "SELECT * from item_inspection " & _
               " WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
    

    Set o_record_set = o_dbConnection.Execute(sSql)
    If Err.Number = 0 Then
    Else
        Let o_error.p_type_cd = "F"
        Let o_error.p_err_cd = "0300"
        Let o_error.p_routine_nm = sRoutine_nm
        Let o_error.p_message_id = 0
        Let o_error.p_nbr = Err.Number
        Let o_error.p_desc = Err.Description
        Let o_error.p_sqlstate_cd = o_dbConnection.Errors(0).SQLState
        Call Err.Clear
        Exit Sub
    End If

    Call Err.Clear

End Sub



