VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "export_vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvaro_dbconnection As Object
Private mvaro_supplier_nm As Object
Private mvaro_bill_nm As Object
Private mvaro_bill_addr As Object
Private mvaro_supplier_addr As Object
Private mvaro_error As Object
Private mvaro_trans_data As Object
Private mvaro_record_set As Object
'

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Bill Name Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_bill_nm(ByVal vData As Object)
    Set mvaro_bill_nm = vData
End Property
Public Property Get o_bill_nm() As Object
    Set o_bill_nm = mvaro_bill_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Bill Address Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_bill_addr(ByVal vData As Object)
    Set mvaro_bill_addr = vData
End Property
Public Property Get o_bill_addr() As Object
    Set o_bill_addr = mvaro_bill_addr
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for supplier Name Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_supplier_nm(ByVal vData As Object)
    Set mvaro_supplier_nm = vData
End Property
Public Property Get o_supplier_nm() As Object
    Set o_supplier_nm = mvaro_supplier_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for supplier Address Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_supplier_addr(ByVal vData As Object)
    Set mvaro_supplier_addr = vData
End Property
Public Property Get o_supplier_addr() As Object
    Set o_supplier_addr = mvaro_supplier_addr
End Property

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

'//////////////////////////////////////////////////////////////////////////////
'Description:   Export the Customer Information to QuickBooks
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub ExportToQuickBooks()
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    ' Define
    Dim sRoutine_nm As String
    Dim sSql As String
    Dim strLine As String
    Dim o_record_set As Object
    Dim objFileName As FileSystemObject
    Dim tsrLine As TextStream
    Dim strFileName As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    Set o_record_set = CreateObject("adodb.recordset")
    sRoutine_nm = "export_customer.cls.ExportToQuickBooks"
    Set o_supplier_nm = CreateObject("converge.name")
    Set o_supplier_addr = CreateObject("converge.address")
    Set o_supplier_nm.o_trans_data = o_trans_data
    Set o_supplier_nm.o_error = o_error
'    Set o_supplier_nm.o_dbconnection = o_dbconnection
    Set o_supplier_addr.o_trans_data = o_trans_data
    Set o_supplier_addr.o_error = o_error
'    Set o_supplier_addr.o_dbconnection = o_dbconnection
    Set objFileName = New FileSystemObject
   
    sSql = " SELECT a.*, b.nm, c.* FROM supplier a, name b, address c " & _
           "  WHERE (a.supplier_id = b.obj_id " & _
           "    AND b.nm_type_cd = 'STND')" & _
           "    AND (a.supplier_id = c.obj_id " & _
           "    AND c.addr_type_cd = 'STND')" & _
           "    AND a.export_fl = 'N'"
           
    Call o_trans_data.OpenRecordset(o_record_set, sSql, o_error)
        
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

    strFileName = "vendor" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & ".iif"
    'Set tsrLine = objFileName.CreateTextFile("c:/converge/customer.inf", True)
    Set tsrLine = objFileName.CreateTextFile(strFileName, True)
  
    tsrLine.WriteLine "!HDR" & vbTab & "PROD" & vbTab & "VER" & vbTab & "REL" & vbTab & "IIFVER" & vbTab & "DATE" & vbTab & "TIME" & vbTab & "ACCNTNT" & vbTab & "ACCNTNTSPLITTIME"
    tsrLine.WriteLine "HDR" & vbTab & "QuickBooks Pro for Windows" & vbTab & "Version 8.0D" & vbTab & "Release R1P" & vbTab & "1" & vbTab & "08/16/2000" & vbTab & "966440020" & vbTab & "N" & vbTab & "0"
    tsrLine.WriteLine "!CUSTNAMEDICT" & vbTab & "INDEX" & vbTab & "LABEL" & vbTab & "CUSTOMER" & vbTab & "VENDOR" & vbTab & "EMPLOYEE"
    tsrLine.WriteLine "!ENDCUSTNAMEDICT"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "0" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "1" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "2" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "3" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "4" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "5" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "6" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "7" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "8" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "9" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "10" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "11" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "12" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "13" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "CUSTNAMEDICT" & vbTab & "14" & vbTab & vbTab & "N" & vbTab & "N" & vbTab & "N"
    tsrLine.WriteLine "ENDCUSTNAMEDICT"
    tsrLine.WriteLine "!VEND" & vbTab & "NAME" & vbTab & "REFNUM" & vbTab & "TIMESTAMP" & vbTab & "PRINTAS" & vbTab & _
                      "ADDR1" & vbTab & "ADDR2" & vbTab & "ADDR3" & vbTab & "ADDR4" & vbTab & "ADDR5" & vbTab & _
                      "VTYPE" & vbTab & "CONT1" & vbTab & "CONT2" & vbTab & "PHONE1" & vbTab & "PHONE2" & vbTab & "FAXNUM" & vbTab & "EMAIL" & vbTab & _
                      "NOTE" & vbTab & "TAXID" & vbTab & "LIMIT" & vbTab & "TERMS" & vbTab & "NOTEPAD" & vbTab & "SALUTATION" & vbTab & "COMPANYNAME" & vbTab & _
                      "FIRSTNAME" & vbTab & "MIDINIT" & vbTab & "LASTNAME" & vbTab & "CUSTFLD1" & vbTab & "CUSTFLD2" & vbTab & "CUSTFLD3" & vbTab & "CUSTFLD4" & vbTab & _
                      "CUSTFLD5" & vbTab & "CUSTFLD6" & vbTab & "CUSTFLD7" & vbTab & "CUSTFLD8" & vbTab & "CUSTFLD9" & vbTab & "CUSTFLD10" & vbTab & "CUSTFLD11" & vbTab & _
                      "CUSTFLD12" & vbTab & "CUSTFLD13" & vbTab & "CUSTFLD14" & vbTab & "CUSTFLD15" & vbTab & "1099" & vbTab & "HIDDEN" & vbTab & "DELCOUNT"
    
    Do Until o_record_set.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
        
        Let o_supplier_nm.p_nm_id = Null
        Let o_supplier_nm.p_obj_nm = "Supplier"
        Let o_supplier_nm.p_obj_id = o_record_set("supplier_id")
        Let o_supplier_nm.p_nm_type_cd = "STND"
        Call o_supplier_nm.getDefName
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
                      
        Let o_supplier_addr.p_addr_id = Null
        Let o_supplier_addr.p_obj_nm = "Supplier"
        Let o_supplier_addr.p_obj_id = o_record_set("supplier_id")
        Let o_supplier_addr.p_addr_type_cd = "STND"
        Let o_supplier_addr.p_def_addr_fl = "Y"
        Call o_supplier_addr.getDefAddress
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
        
           
        strLine = "VEND " & vbTab & _
                 Trim(o_record_set("nm")) & vbTab & o_record_set("supplier_id") & vbTab & "966437205" & vbTab & Trim(o_record_set("nm")) & vbTab & _
                 Trim(o_record_set("nm")) & vbTab & Trim(o_supplier_addr.p_Street_addr) & vbTab & Trim(o_supplier_addr.p_Street_addr) & vbTab & "'" & Trim(o_supplier_addr.p_City_nm) & ", " & Trim(o_supplier_addr.p_State_cd) & " " & Trim(o_supplier_addr.p_Zip_cd) & "'" & vbTab & _
                 vbTab & _
                 "VTYPE" & vbTab & "cont1" & vbTab & "cont2" & vbTab & _
                 Trim(o_record_set("phone_nbr")) & vbTab & vbTab & Trim(o_record_set("fax_nbr")) & vbTab & Trim(o_record_set("email_addr")) & vbTab & _
                 "note" & vbTab & "taxid" & vbTab & "limit" & vbTab & "terms" & vbTab & "notepad" & vbTab & "salutation" & vbTab & "comp name" & vbTab & "f name" & vbTab & "initial" & vbTab & "l name" & vbTab & _
                 vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & _
                 "N" & vbTab & "N" & vbTab & "0"

                 

                 
        ' Write Line to file name
        tsrLine.WriteLine strLine
        
        'CUST    GSS 3   966437205   GSS Jason R Paarlberg   7655 W Alex "Middletown, OH 45042"      GSS Jason R Paarlberg   7655 W Alex "Middletown, OH 45042"
        
 '       sSql = "UPDATE customer SET export_fl = 'y'" & _
 '              " WHERE " & fWhereVariable("cust_id", o_record_set("cust_id"), "N", "=")
 '       Call o_trans_data.ExecuteDatabase(sSql, dblUpdateQty, o_error)
        If Err.Number = 0 Then
        Else
            Let o_error.p_type_cd = "F"
            Let o_error.p_err_cd = "0300"
            Let o_error.p_routine_nm = sRoutine_nm
            Let o_error.p_message_id = 0
            Let o_error.p_nbr = Err.Number
            Let o_error.p_desc = Err.Description
            'Let o_error.p_sqlstate_cd = o_dbconnection.Errors(0).SQLState
            Call Err.Clear
            Exit Sub
        End If
            
        o_record_set.MoveNext
    Loop

    tsrLine.Close
    
End Sub


