VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_dbconnection As Object
Private mvaro_record_set As Object
Private mvaro_trans_data As Object
Private mvaro_error As Object
'

Public Property Set o_error(ByVal vData As Object)
    Set mvaro_error = vData
End Property
Public Property Get o_error() As Object
    Set o_error = mvaro_error
End Property
Public Property Set o_trans_data(ByVal vData As Object)
    Set mvaro_trans_data = vData
End Property
Public Property Get o_trans_data() As Object
    Set o_trans_data = mvaro_trans_data
End Property
Public Property Set o_record_set(ByVal vData As ADODB.Recordset)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As ADODB.Recordset
    Set o_record_set = mvaro_record_set
End Property
Public Property Set o_dbconnection(ByVal vData As Object)
    Set mvaro_dbconnection = vData
End Property
Public Property Get o_dbconnection() As Object
    Set o_dbconnection = mvaro_dbconnection
End Property

Public Sub PrintAddressReport()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    On Error Resume Next
    Dim sRoutine_nm As String
    Dim objReport As CrystalReport
    'Dim objReport1 As Crystal
    
    Dim objRecordSet As ADODB.Recordset
    Dim strSql As String
    
    sRoutine_nm = "report.cls.PrintAddressReport"
    strSql = "SELECT * from Address "
    Set objReport = CreateObject("CrystalReport")
    
    
    
    Call o_trans_data.OpenRecordset(objRecordSet, strSql, o_error)
    
    objReport.ReportFileName = "c:\program files\converge\report\address.rpt"
    
    objReport.SetTablePrivateData 0, 3, objRecordSet
    
    objReport.Action = 1
    
End Sub
