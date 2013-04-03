Attribute VB_Name = "modGeneral"
Option Explicit
Public Const G_STR_PROJECT_NM = "convergeTestClient"
Private Const M_STR_MODULE_NM = "modGeneral"
Global g_objTransData As Object
Global g_objError As Object
Global g_objConnection As ADODB.Connection

Public Sub main()
    Set g_objError = CreateObject("Converge.Error")
    Set g_objTransData = CreateObject("converge.Trans_data")
    Set g_objConnection = New ADODB.Connection
    Call g_objConnection.Open("Provider=SQLOLEDB; DRIVER=SQL Server; UID=Converge; PWD=volters; WSID=TIMCOSERVER; APP=Converge; DATABASE=Converge; SERVER=(local)")
    Set g_objTransData.o_dbConnection = g_objConnection
    g_objTransData.p_connectstring = g_objConnection.ConnectionString
    Call LoadSystemTables(g_objConnection)
    
    Dim aForm As frmTest
    Set aForm = New frmTest
    aForm.Show
    
End Sub


Function LoadSystemTables(ByRef r_objConnection As ADODB.Connection)
    
    Dim strSql As String
    Dim rstCodeDesc As ADODB.Recordset
    Dim rstStateCode As ADODB.Recordset

    strSql = "SELECT * FROM code_desc ORDER by field_nm, seq_nbr "
    Set rstCodeDesc = CreateObject("ADODB.RecordSet")
    Call rstCodeDesc.Open(strSql, r_objConnection, adOpenStatic)
    g_objTransData.p_CodeDesc = rstCodeDesc.GetRows()

    strSql = "SELECT * FROM code_desc WHERE field_nm = 'state_cd' ORDER by seq_nbr "
    Set rstStateCode = CreateObject("ADODB.RecordSet")
    Call rstStateCode.Open(strSql, r_objConnection, adOpenStatic)
    g_objTransData.p_StateCode = rstStateCode.GetRows()

End Function


