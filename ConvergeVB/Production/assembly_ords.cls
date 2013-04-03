VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "assembly_ords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object 'local copy
Private mvaro_dbconnection As Object 'local copy
Private mvaro_trans_data As Object 'local copy
Private mvaro_record_set As Object 'local copy
Private mvarp_item_id As Variant 'local copy
Private mvarp_status_cd As Variant 'local copy
Private mvarp_due_dt As Variant 'local copy
Public Property Let p_due_dt(ByVal vData As Variant)
    mvarp_due_dt = vData
End Property
Public Property Set p_due_dt(ByVal vData As Variant)
    Set mvarp_due_dt = vData
End Property
Public Property Get p_due_dt() As Variant
    If IsObject(mvarp_due_dt) Then
        Set p_due_dt = mvarp_due_dt
    Else
        p_due_dt = mvarp_due_dt
    End If
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Set p_status_cd(ByVal vData As Variant)
    Set mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
    If IsObject(mvarp_status_cd) Then
        Set p_status_cd = mvarp_status_cd
    Else
        p_status_cd = mvarp_status_cd
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
Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property
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
Public Sub GetAssemblyOrders()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim sRoutine_nm As String
    Let sRoutine_nm = "assembly_ords.cls.GetAssemblyOrders"
    Dim sSql As String
    
    Set o_record_set = CreateObject("adodb.recordset")
    
    Let sSql = "SELECT * from assembly_order "

    
    If IsNull(p_item_id) Then
        If IsNull(p_status_cd) Or p_status_cd = "" Then
           
        Else
             Let sSql = sSql & " WHERE " & fWhereVariable("status_cd", p_status_cd, "S", "=")
        End If
    Else
        Let sSql = sSql & " WHERE " & fWhereVariable("item_id", p_item_id, "N", "=")
        
        If IsNull(p_status_cd) Or p_status_cd = "" Then
             Let sSql = sSql & " AND " & fWhereVariable("status_cd", p_status_cd, "S", "=")
        
        Else
        
        End If
    End If
    

    Set o_record_set = o_dbconnection.Execute(sSql)
    If Err.Number = 0 Then
    Else
        Let o_error.p_type_cd = "F"
        Let o_error.p_err_cd = "0300"
        Let o_error.p_routine_nm = sRoutine_nm
        Let o_error.p_message_id = 0
        Let o_error.p_nbr = Err.Number
        Let o_error.p_desc = Err.Description
        Let o_error.p_sqlstate_cd = o_dbconnection.Errors(0).SQLState
        Call Err.Clear
        Exit Sub
    End If

    Call Err.Clear

End Sub


