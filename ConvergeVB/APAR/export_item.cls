VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "export_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_dbconnection As Object
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As Object
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Transaction Data Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Error Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Database Connection Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Database Connection Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property

Public Sub ExportToQuickBooks()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
    
    Dim sRoutine_nm As String
    Dim strSql As String
    Dim strLine As String
    Dim o_record_set As ADODB.Recordset
    Dim objFileName As FileSystemObject
    Dim tsrLine As TextStream
    Dim strFileName As String
    Dim strItemNbr As String
    Dim strItemDesc As String
    Dim intLength As Integer
    Dim intCount As Integer
    
            
        
    'init
    Set o_record_set = New ADODB.Recordset
    sRoutine_nm = "export_item.cls.ExportToQuickBooks"
    Set objFileName = New FileSystemObject
    
    strSql = " SELECT * from item "

    Call o_trans_data.OpenRecordset(o_record_set, strSql, o_error)
    
    If Err.Number = 0 Then
    Else
        Let o_error.p_type_cd = "F"
        Let o_error.p_err_cd = "0300"
        Let o_error.p_routine_nm = sRoutine_nm
        Let o_error.p_message_id = 0
        Let o_error.p_nbr = Err.Number
        Let o_error.p_desc = Err.Description
        Call Err.Clear
        Exit Sub
    End If
    
    strFileName = "c:\temp\item-" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".iif"
    
    Set tsrLine = objFileName.CreateTextFile(strFileName, True)
    
    tsrLine.WriteLine "!INVITEM,NAME,INVITEMTYPE,DESC,PRICE,COST"
    
    o_record_set.MoveFirst
    
    Do Until o_record_set.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
    
        strItemDesc = ""
        strItemNbr = ""
        intLength = Len(o_record_set("item_desc"))
        
        If InStr(1, o_record_set("item_desc"), ",") > 0 Then
            For intCount = 1 To intLength
                If Mid(o_record_set("item_desc"), intCount, 1) = "," Then
                    strItemDesc = strItemDesc & " "
                Else
                    strItemDesc = strItemDesc & Mid(o_record_set("item_desc"), intCount, 1)
                End If
            Next intCount
        Else
            strItemDesc = o_record_set("item_desc")
        End If
        
        intLength = Len(o_record_set("item_nbr"))
        
        If InStr(1, o_record_set("item_nbr"), ",") > 0 Then
            For intCount = 1 To intLength
                If Mid(o_record_set("item_nbr"), intCount, 1) = "," Then
                    strItemNbr = strItemNbr & " "
                Else
                    strItemNbr = strItemNbr & Mid(o_record_set("item_nbr"), intCount, 1)
                End If
            Next intCount
        Else
            strItemNbr = o_record_set("item_nbr")
        End If
        
        strLine = "INVITEM," & strItemNbr & "," & "INVENTORY" & "," & strItemDesc & _
                  "," & o_record_set("item_price") & "," & o_record_set("item_cost")
        
        tsrLine.WriteLine strLine
        
        If Err.Number = 0 Then
        Else
            Let o_error.p_type_cd = "F"
            Let o_error.p_err_cd = "0300"
            Let o_error.p_routine_nm = sRoutine_nm
            Let o_error.p_message_id = 0
            Let o_error.p_nbr = Err.Number
            Let o_error.p_desc = Err.Description
            Call Err.Clear
            Exit Sub
        End If
            
        o_record_set.MoveNext
    Loop
    
    Set o_record_set = Nothing
    tsrLine.Close
End Sub
