VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_varSeq_nm As Variant
Private m_varSeq_nbr As Variant
Private m_objError As Object
Private m_objTrans_data As Object

Public Property Set o_trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = m_objTrans_data
End Property
Public Property Set o_error(ByVal vData As Object)
    Set m_objError = vData
End Property
Public Property Get o_error() As Object
    Set o_error = m_objError
End Property
Public Property Let p_seq_nbr(ByVal vData As Variant)
    m_varSeq_nbr = vData
End Property
Public Property Get p_seq_nbr() As Variant
        p_seq_nbr = m_varSeq_nbr
End Property
Public Property Let p_seq_nm(ByVal vData As Variant)
    m_varSeq_nm = vData
End Property
Public Property Get p_seq_nm() As Variant
        p_seq_nm = m_varSeq_nm
End Property



Public Property Set Trans_data(ByVal vData As Object)
    Set m_objTrans_data = vData
End Property
Public Property Get Trans_data() As Object
    Set Trans_data = m_objTrans_data
End Property
Public Property Set error(ByVal vData As Object)
    Set o_error = vData
End Property
Public Property Get error() As Object
    Set error = m_objError
End Property
Public Property Let seq_nbr(ByVal vData As Variant)
    m_varSeq_nbr = vData
End Property
Public Property Get seq_nbr() As Variant
        seq_nbr = m_varSeq_nbr
End Property
Public Property Let seq_nm(ByVal vData As Variant)
    m_varSeq_nm = vData
End Property
Public Property Get seq_nm() As Variant
        seq_nm = m_varSeq_nm
End Property




Public Sub getSeqNbr()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSqlSelect As String
    Dim strSqlUpdate As String
    Dim sSuccessful_fl As Boolean
    Dim dblRecords_affected As Double
    Dim iLoop_ct As Double
    Dim objNewConnection As adodb.Connection
    Dim objRecordset As adodb.Recordset
    
    strRoutine_nm = "sequence.cls.GetSeqNbr"
    
    sSuccessful_fl = False
    iLoop_ct = 0
    
    Set objNewConnection = New adodb.Connection
    Call objNewConnection.Open(o_trans_data.ConnectString)
    
    strSqlSelect = "SELECT * FROM sequence "
    strSqlSelect = strSqlSelect & "WHERE "
    strSqlSelect = strSqlSelect & " " & fWhereVariable("seq_nm", seq_nm, "S", "=")

    Do Until sSuccessful_fl = True Or iLoop_ct > 100 Or o_error.type_cd = "E" Or o_error.type_cd = "F"
        
        iLoop_ct = iLoop_ct + 1
        
        Call o_trans_data.OpenRecordset(objRecordset, strSqlSelect, o_error)
        
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            objRecordset.Close
            Set objRecordset = Nothing
            Exit Do
        End If
        
        If objRecordset.EOF = True Then
            o_error.type_cd = "F"
            o_error.err_cd = "0300"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
            Exit Do
        End If
        
        p_seq_nbr = objRecordset("seq_nbr")
        
        Call objRecordset.Close
        Set objRecordset = Nothing
        
        Call objNewConnection.BeginTrans
        
        strSqlUpdate = "UPDATE sequence "
        strSqlUpdate = strSqlUpdate & "set "
        strSqlUpdate = strSqlUpdate & "  seq_nbr = seq_nbr + 1 "
        strSqlUpdate = strSqlUpdate & "where "
        strSqlUpdate = strSqlUpdate & "   " & fWhereVariable("seq_nm", m_varSeq_nm, "S", "=")
        strSqlUpdate = strSqlUpdate & "and" & fWhereVariable("seq_nbr", m_varSeq_nbr, "N", "=")
        
        Call objNewConnection.Execute(strSqlUpdate, dblRecords_affected)
        
        If dblRecords_affected = 1 Then
            Call objNewConnection.CommitTrans
            sSuccessful_fl = True
        Else
            Call objNewConnection.RollbackTrans
        End If
    Loop
        
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
    Else
        If sSuccessful_fl = False Then
            o_error.type_cd = "F"
            o_error.err_cd = "0900"
            o_error.routine_nm = strRoutine_nm
            o_error.message_id = 0
        End If
    End If
    
    If o_error.type_cd = "E" Or o_error.type_cd = "F" Then
        Call objNewConnection.Close
    Else
        Call objNewConnection.Close
    End If
    Set objNewConnection = Nothing

    Exit Sub
error_handler:
    With o_error
      .p_type_cd = "F"
      .p_err_cd = "0100"
      .p_nbr = Err.Number
      .p_desc = Err.Description
      .p_routine_nm = strRoutine_nm
      .p_message_id = 0
    End With
    Err.Clear
End Sub




