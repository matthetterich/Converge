VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "export_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_dbconnection As Object
Private mvaro_bill_nm As Object
Private mvaro_bill_addr As Object
Private mvaro_ship_nm As Object
Private mvaro_ship_addr As Object
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
'Description:   Set and Get for Ship Name Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_ship_nm(ByVal vData As Object)
    Set mvaro_ship_nm = vData
End Property
Public Property Get o_ship_nm() As Object
    Set o_ship_nm = mvaro_ship_nm
End Property

'//////////////////////////////////////////////////////////////////////////////
'Description:   Set and Get for Ship Address Object
'
'//////////////////////////////////////////////////////////////////////////////
Public Property Set o_ship_addr(ByVal vData As Object)
    Set mvaro_ship_addr = vData
End Property
Public Property Get o_ship_addr() As Object
    Set o_ship_addr = mvaro_ship_addr
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
    Dim strName As String
    Dim strBillAddr As String
    Dim strShipAddr As String
    Dim lngLength As Long
    Dim dblUpdateQty As Double
    
    ' Initialize
    Set o_record_set = CreateObject("adodb.recordset")
    sRoutine_nm = "export_customer.cls.ExportToQuickBooks"
    Set o_ship_nm = CreateObject("converge.name")
    Set o_ship_addr = CreateObject("converge.address")
    Set o_ship_nm.o_trans_data = o_trans_data
    Set o_ship_nm.o_error = o_error
    Set o_ship_addr.o_trans_data = o_trans_data
    Set o_ship_addr.o_error = o_error
    Set o_bill_nm = CreateObject("converge.name")
    Set o_bill_addr = CreateObject("converge.address")
    Set o_bill_nm.o_trans_data = o_trans_data
    Set o_bill_nm.o_error = o_error
    Set o_bill_addr.o_trans_data = o_trans_data
    Set o_bill_addr.o_error = o_error
    Set objFileName = New FileSystemObject
   
    sSql = " SELECT a.*, b.nm, c.* FROM customer a, name b, address c " & _
           "  WHERE (a.cust_id = b.obj_id " & _
           "    AND b.nm_type_cd = 'STND')" & _
           "    AND (a.cust_id = c.obj_id " & _
           "    AND c.addr_type_cd = 'STND')" & _
           "    AND a.export_fl = 'N'" & _
           "    AND status_cd = 'AC' "
           
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

    strFileName = "c:\temp\customer" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & ".iif"
    'Set tsrLine = objFileName.CreateTextFile("c:/converge/customer.inf", True)
    Set tsrLine = objFileName.CreateTextFile(strFileName, True)
  
    tsrLine.WriteLine "!CUST,NAME,BADDR1,BADDR2,BADDR3,BADDR4,BADDR5," & _
                      "SADDR1,SADDR2,SADDR3,SADDR4,SADDR5,PHONE1,PHONE2," & _
                      "FAXNUM,CONT1,CONT2,CTYPE,TERMS,TAXABLE,LIMIT," & _
                      "RESALENUM,REP,TAXITEM,NOTEPAD,SALUTATION,FIRSTNAME,MIDINIT,LASTNAME," & _
                      "CUSTFLD1,CUSTFLD2,CUSTFLD3,CUSTFLD4,CUSTFLD5,CUSTFLD6,CUSTFLD7," & _
                      "JOBDESC,JOBTYPE,JOBSTATUS,JOBSTART,JOBPROJEND,JOBEND,NOTE"
                      
    Do Until o_record_set.EOF = True Or o_error.p_type_cd = "E" Or o_error.p_type_cd = "F"
    
        Let o_ship_nm.p_nm_id = Null
        Let o_ship_nm.p_obj_nm = "Customer"
        Let o_ship_nm.p_obj_id = o_record_set("cust_id")
        Let o_ship_nm.p_nm_type_cd = "SHIP"
        Call o_ship_nm.getDefName
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
                      
        Let o_ship_addr.p_addr_id = Null
        Let o_ship_addr.p_obj_nm = "Customer"
        Let o_ship_addr.p_obj_id = o_record_set("cust_id")
        Let o_ship_addr.p_addr_type_cd = "SHIP"
        Let o_ship_addr.p_def_addr_fl = "Y"
        Call o_ship_addr.getDefAddress
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
        
        If IsNull(o_ship_addr.p_addr_id) Then
            Let o_ship_addr.p_obj_nm = "Customer"
            Let o_ship_addr.p_obj_id = o_record_set("cust_id")
            Let o_ship_addr.p_addr_type_cd = "STND"
            Let o_ship_addr.p_def_addr_fl = "Y"
            Call o_ship_addr.getDefAddress
            If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
                Exit Sub
            End If
        End If
        
        Let o_bill_nm.p_nm_id = Null
        Let o_bill_nm.p_obj_nm = "Customer"
        Let o_bill_nm.p_obj_id = o_record_set("cust_id")
        Let o_bill_nm.p_nm_type_cd = "BILL"
        Call o_bill_nm.getDefName
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
                      
        Let o_bill_addr.p_addr_id = Null
        Let o_bill_addr.p_obj_nm = "Customer"
        Let o_bill_addr.p_obj_id = o_record_set("cust_id")
        Let o_bill_addr.p_addr_type_cd = "BILL"
        Let o_bill_addr.p_def_addr_fl = "Y"
        Call o_bill_addr.getDefAddress
        If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
            Exit Sub
        End If
        
        If IsNull(o_bill_addr.p_addr_id) Then
            Let o_bill_addr.p_obj_nm = "Customer"
            Let o_bill_addr.p_obj_id = o_record_set("cust_id")
            Let o_bill_addr.p_addr_type_cd = "STND"
            Let o_bill_addr.p_def_addr_fl = "Y"
            Call o_bill_addr.getDefAddress
            If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
                Exit Sub
            End If
        End If
        
        lngLength = InStr(1, Trim(o_record_set("nm")), ",", 1)
        
        If lngLength > 0 Then
            strName = Left(Trim(o_record_set("nm")), lngLength - 1)
        Else
            strName = Trim(o_record_set("nm"))
        End If
            
        lngLength = InStr(1, Trim(o_bill_addr.p_Street_addr), Chr(13), 1)
        
        If lngLength > 0 Then
            strBillAddr = Left(Trim(o_bill_addr.p_Street_addr), lngLength - 1) & " " & Right(Trim(o_bill_addr.p_Street_addr), Len(Trim(o_bill_addr.p_Street_addr)) - lngLength - 1)
        Else
            strBillAddr = Trim(o_bill_addr.p_Street_addr)
        End If
        
        lngLength = InStr(1, Trim(o_ship_addr.p_Street_addr), Chr(13), 1)
        
        If lngLength > 0 Then
            strShipAddr = Left(Trim(o_ship_addr.p_Street_addr), lngLength - 1) & " " & Right(Trim(o_ship_addr.p_Street_addr), Len(Trim(o_ship_addr.p_Street_addr)) - lngLength - 1)
        Else
            strShipAddr = Trim(o_ship_addr.p_Street_addr)
        End If
        
        
                    
        strLine = "CUST," & _
                 Trim(o_record_set("cust_nbr")) & "," & _
                 strName & "," & strBillAddr & "," & Trim(o_bill_addr.p_City_nm) & "," & Trim(o_bill_addr.p_State_cd) & "," & Trim(o_bill_addr.p_Zip_cd) & "," & _
                 "" & _
                 strName & "," & strShipAddr & "," & Trim(o_ship_addr.p_City_nm) & "," & Trim(o_ship_addr.p_State_cd) & "," & Trim(o_ship_addr.p_Zip_cd) & "," & _
                 "" & _
                 Trim(o_record_set("phone_nbr")) & "," & "," & Trim(o_record_set("fax_nbr"))
                 
        ' Write Line to file name
        tsrLine.WriteLine strLine
        
        o_trans_data.p_update_fl = True
        sSql = "UPDATE customer SET " & _
               " " & fUpdateVariable("export_fl", "Y", "S", "=") & _
               " WHERE " & fWhereVariable("cust_id", o_record_set("cust_id"), "N", "=")
        
        
        Call o_trans_data.ExecuteDatabase(sSql, dblUpdateQty, o_error)
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
