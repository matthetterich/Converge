Attribute VB_Name = "FmtSqlVar"
Option Explicit

Global objUtilities As Object

Public Function fInsertVariable(sHost_var, sData_type_cd) As String
    
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    Let objUtilities.p_host_var = sHost_var
    Let objUtilities.p_data_type_cd = sData_type_cd
    Let objUtilities.p_format_str = ""
    Call objUtilities.FormatInsertVariable
    Let fInsertVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function

Public Function fUpdateVariable(sUpdate_var_nm, sHost_var, sData_type_cd, sOperation_cd) As String
     
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    Let objUtilities.p_host_var = sHost_var
    Let objUtilities.p_data_type_cd = sData_type_cd
    Let objUtilities.p_host_var_nm = sUpdate_var_nm
    Let objUtilities.p_Operation_cd = sOperation_cd
    Let objUtilities.p_format_str = ""
    Call objUtilities.FormatUpdateVariable
    Let fUpdateVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function



Public Function fWhereVariable(sUpdate_var_nm, sHost_var, sData_type_cd, sOperation_cd) As String
     
    Dim objUtilities
    Set objUtilities = CreateObject("converge.utilities")
    Let objUtilities.p_host_var = sHost_var
    Let objUtilities.p_data_type_cd = sData_type_cd
    Let objUtilities.p_host_var_nm = sUpdate_var_nm
    Let objUtilities.p_Operation_cd = sOperation_cd
    Let objUtilities.p_format_str = ""
    Call objUtilities.FormatWhereVariable
    Let fWhereVariable = objUtilities.p_format_str
    Set objUtilities = Nothing

End Function




