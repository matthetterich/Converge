VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "utilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarp_host_var As Variant
Private mvarp_format_str As Variant
Private mvarp_data_type_cd As Variant
Private mvarp_host_var_nm As Variant
Private mvarp_Operation_cd As Variant
Private mvaro_error As Object
Private mvaro_trans_data As Object

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
Public Property Let p_Operation_cd(ByVal vData As Variant)
    mvarp_Operation_cd = vData
End Property
Public Property Set p_Operation_cd(ByVal vData As Variant)
    Set mvarp_Operation_cd = vData
End Property
Public Property Get p_Operation_cd() As Variant
    p_Operation_cd = mvarp_Operation_cd
End Property
Public Property Let p_host_var_nm(ByVal vData As Variant)
    mvarp_host_var_nm = vData
End Property
Public Property Set p_host_var_nm(ByVal vData As Variant)
    Set mvarp_host_var_nm = vData
End Property
Public Property Get p_host_var_nm() As Variant
    p_host_var_nm = mvarp_host_var_nm
End Property
Public Property Let p_data_type_cd(ByVal vData As Variant)
    mvarp_data_type_cd = vData
End Property
Public Property Set p_data_type_cd(ByVal vData As Variant)
    Set mvarp_data_type_cd = vData
End Property
Public Property Get p_data_type_cd() As Variant
    p_data_type_cd = mvarp_data_type_cd
End Property
Public Property Let p_format_str(ByVal vData As Variant)
    mvarp_format_str = vData
End Property
Public Property Set p_format_str(ByVal vData As Variant)
    Set mvarp_format_str = vData
End Property
Public Property Get p_format_str() As Variant
    p_format_str = mvarp_format_str
End Property
Public Property Let p_host_var(ByVal vData As Variant)
    mvarp_host_var = vData
End Property
Public Property Set p_host_var(ByVal vData As Variant)
    Set mvarp_host_var = vData
End Property
Public Property Get p_host_var() As Variant
    p_host_var = mvarp_host_var
End Property

Public Sub FormatInsertVariable()
            
    'Call FormatAccessInsertVariable
    ''Call FormatOracleInsertVariable
    Call FormatSqlserverInsertVariable
    
End Sub


Public Sub FormatUpdateVariable()
            
    'Call FormatAccessUpdateVariable
    'Call FormatOracleInsertVariable
    Call FormatSqlServerUpdateVariable
    
End Sub


Public Sub FormatWhereVariable()
            
    'Call FormatAccessWhereVariable
    'Call FormatOracleInsertVariable
    Call FormatSqlServerWhereVariable
    
End Sub


Private Sub FormatAccessInsertVariable()

    If fValuePresent(p_host_var) = False Then
        p_format_str = " null "
        Exit Sub
    End If

    If p_host_var = "current_ts" Then
        p_format_str = " now "
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = " " & p_host_var & " "
        Case "S"
           p_format_str = " '" & fFormatTextField(p_host_var) & "' "
        Case "D"
           p_format_str = " datevalue('" & p_host_var & "') "
    End Select
        
End Sub


Private Sub FormatSqlserverInsertVariable()
    If fValuePresent(p_host_var) = False Then
        p_format_str = " null "
        Exit Sub
    End If

    If p_host_var = "current_ts" Then
        p_format_str = " GetDate() "
        
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = p_host_var
        Case "S"
           p_format_str = "'" & fFormatTextField(p_host_var) & "' "
        Case "D"
            p_format_str = " convert(datetime,'" & Format(p_host_var, "yyyymmdd") & "', 112 )"
    End Select
    

End Sub


Private Sub FormatOracleInsertVariable()

End Sub



Private Sub FormatAccessUpdateVariable()

    If IsEmpty(p_host_var) = True Then
        p_format_str = " " & p_host_var_nm & " = " & p_host_var_nm
        Exit Sub
    End If

    If Len(p_host_var) = 0 Then
        p_format_str = " " & p_host_var_nm & " = " & p_host_var_nm
        Exit Sub
    End If

    If IsNull(p_host_var) Then
        p_format_str = " " & p_host_var_nm & p_Operation_cd & " NULL "
        Exit Sub
    End If
        
    If p_host_var = "current_ts" Then
        p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " now "
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " " & p_host_var & " "
        Case "S"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " '" & fFormatTextField(p_host_var) & "' "
        Case "D"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " datevalue('" & p_host_var & "') "
    End Select

End Sub


Private Sub FormatSqlServerUpdateVariable()
    
    If IsEmpty(p_host_var) = True Then
        p_format_str = " " & p_host_var_nm & " = " & p_host_var_nm
        Exit Sub
    End If

    If Len(p_host_var) = 0 Then
        p_format_str = " " & p_host_var_nm & " = " & p_host_var_nm
        Exit Sub
    End If

    If IsNull(p_host_var) Then
        p_format_str = " " & p_host_var_nm & p_Operation_cd & " NULL "
        Exit Sub
    End If
        
    If p_host_var = "current_ts" Then
        p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " getdate() "
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " " & p_host_var & " "
        Case "S"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " '" & fFormatTextField(p_host_var) & "' "
        Case "D"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " convert(datetime,'" & Format(p_host_var, "yyyymmdd") & "',112) "
    End Select

End Sub


Private Sub FormatOracleUpdateVariable()

End Sub



Private Sub FormatAccessWhereVariable()

    If fValuePresent(p_host_var) = False Then
        If p_Operation_cd = "=" Then
            p_format_str = " " & p_host_var_nm & " is null "
        Else
            p_format_str = " " & p_host_var_nm & " is not null "
        End If
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & p_host_var & " "
        Case "S"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " '" & fFormatTextField(p_host_var) & "' "
        Case "D"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " datevalue('" & p_host_var & "') "
    End Select

End Sub


Private Sub FormatSqlServerWhereVariable()
    If fValuePresent(p_host_var) = False Then
        If p_Operation_cd = "=" Then
            p_format_str = " " & p_host_var_nm & " is null "
        Else
            p_format_str = " " & p_host_var_nm & " is not null "
        End If
        Exit Sub
    End If

    Select Case p_data_type_cd
        Case "N"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & p_host_var
        Case "S"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " " & " '" & fFormatTextField(p_host_var) & "' "
        Case "D"
            p_format_str = " " & p_host_var_nm & " " & p_Operation_cd & " convert(datetime,'" & Format(p_host_var, "yyyymmdd") & "',112) "
    End Select

End Sub


Private Sub FormatOracleWhereVariable()

End Sub



Private Function fFormatTextField(sStr)
  
  Dim tStr
  Dim iPos

  tStr = sStr
  iPos = InStr(1, tStr, "'")
  If iPos = 0 Then
     fFormatTextField = tStr
     Exit Function
  End If

  While iPos > 0
    tStr = Left(tStr, iPos) & "'" & Mid(tStr, (iPos + 1))
    iPos = iPos + 2
    iPos = InStr(iPos, tStr, "'")
  Wend

  fFormatTextField = tStr

End Function


