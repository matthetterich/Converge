VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_error As Object 'local copy
Private mvaro_trans_data As Object 'local copy
Private mvaro_record_set As Object 'local copy
Private mvaro_dbconnection As Object 'local copy
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_dbconnection 'Set' and 'Get' are used to
'               set up the database connection.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_record_set 'Set' and 'Get' are used to
'               set up the record set.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_trans_data 'Set' and 'Get' are used to
'               set up the transaction data.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   The o_error 'Set' and 'Get' are used to
'               set up the error object.
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:       Gets the list of messages from the database,
'                   calls SELECT function
'
'Arguments:         None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub GetMessages()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim sRoutine_nm As String
    Let sRoutine_nm = "messages.cls.getMessages"
    Dim sSql As String
    
    Set o_record_set = CreateObject("adodb.recordset")
    
    Let sSql = "SELECT * FROM error_message ORDER BY message_id "

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


