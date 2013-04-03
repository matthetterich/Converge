VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "phy_inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvaro_error As Object
Private mvaro_record_set As Object
Private mvaro_trans_data As Object
Private mvarp_opened_dt As Variant
Private mvarp_closed_dt As Variant
Private mvarp_status_desc As Variant
Private mvarp_status_cd As Variant
Private mvarp_phy_inv_id As Variant
Private mvarp_start_entered_dt As Variant
Private mvarp_end_entered_dt As Variant
Private mvarp_start_closed_dt As Variant
Private mvarp_end_closed_dt As Variant
'

Public Property Let p_start_entered_dt(ByVal vData As Variant)
    mvarp_start_entered_dt = vData
End Property
Public Property Get p_start_entered_dt() As Variant
        p_start_entered_dt = mvarp_start_entered_dt
End Property

Public Property Let p_end_entered_dt(ByVal vData As Variant)
    mvarp_end_entered_dt = vData
End Property
Public Property Get p_end_entered_dt() As Variant
        p_end_entered_dt = mvarp_end_entered_dt
End Property

Public Property Let p_start_closed_dt(ByVal vData As Variant)
    mvarp_start_closed_dt = vData
End Property
Public Property Get p_start_closed_dt() As Variant
        p_start_closed_dt = mvarp_start_closed_dt
End Property

Public Property Let p_end_closed_dt(ByVal vData As Variant)
    mvarp_end_closed_dt = vData
End Property
Public Property Get p_end_closed_dt() As Variant
        p_end_closed_dt = mvarp_end_closed_dt
End Property

Public Property Let p_phy_inv_id(ByVal vData As Variant)
    mvarp_phy_inv_id = vData
End Property
Public Property Get p_phy_inv_id() As Variant
        p_phy_inv_id = mvarp_phy_inv_id
End Property
Public Property Let p_status_cd(ByVal vData As Variant)
    mvarp_status_cd = vData
End Property
Public Property Get p_status_cd() As Variant
        p_status_cd = mvarp_status_cd
End Property
Public Property Let p_status_desc(ByVal vData As Variant)
    mvarp_status_desc = vData
End Property
Public Property Get p_status_desc() As Variant
        p_status_desc = mvarp_status_desc
End Property
Public Property Let p_closed_dt(ByVal vData As Variant)
    mvarp_closed_dt = vData
End Property
Public Property Get p_closed_dt() As Variant
        p_closed_dt = mvarp_closed_dt
End Property
Public Property Let p_opened_dt(ByVal vData As Variant)
    mvarp_opened_dt = vData
End Property
Public Property Get p_opened_dt() As Variant
        p_opened_dt = mvarp_opened_dt
End Property
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

Public Property Set o_record_set(ByVal vData As Object)
    Set mvaro_record_set = vData
End Property
Public Property Get o_record_set() As Object
    Set o_record_set = mvaro_record_set
End Property
'

Public Sub GetPhyInv()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
'init
    strRoutine_nm = "phy_inv.cls.GetPhyInv"

    strSql = "SELECT *      "
    strSql = strSql & "FROM phy_inv_order "
    strSql = strSql & "WHERE         "
    strSql = strSql & "   " & fWhereVariable("phy_inv_ord_id", p_phy_inv_id, "N", "=")
    
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        objRecordset.Close
        Set objRecordset = Nothing
        Exit Sub
    End If
    
    If objRecordset.EOF = True Then
        p_phy_inv_id = Null
    Else
        p_phy_inv_id = objRecordset("phy_inv_ord_id")
        p_status_cd = objRecordset("status_cd")
        p_status_desc = fGetDesc("phy_inv_status_cd", p_status_cd, o_error, o_trans_data)
        p_opened_dt = objRecordset("entered_dt")
        p_closed_dt = objRecordset("closed_dt")
    End If
     
    objRecordset.Close
    Set objRecordset = Nothing
    
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


Public Sub Add()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "phy_inv.cls.Add"
    
    Call validate
    Call IssueInsert
    
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



Public Sub Change()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "phy_inv.cls.Change"
    
    Call validate
    Call IssueUpdate
    
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


Public Sub Delete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "phy_inv.cls.Delete"
    
    Call IssueDelete
    
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




Private Sub validate()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    strRoutine_nm = "phy_inv.cls.Validate"

'possible change to p_item_id
    If Len(p_status_cd) > 0 Then
    Else
        o_error.type_cd = "F"
        o_error.err_cd = "0600"
        o_error.message_id = 0
        o_error.routine_nm = strRoutine_nm
        o_error.nbr = Err.Number
        o_error.Desc = Err.Description
    End If
    
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



Private Sub IssueInsert()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    strRoutine_nm = "phy_inv.cls.IssueInsert"
    
    p_phy_inv_id = fGetId("phy_inv_id", o_error, o_trans_data)
    
    strSql = "INSERT INTO phy_inv_order ( " & _
                "  phy_inv_ord_id " & _
                ", status_cd " & _
                ", entered_dt " & _
                ", closed_dt " & _
                ", updated_assoc_id" & _
                ", updated_ts" & _
                ", entered_assoc_id" & _
                ", entered_ts" & _
                " ) " & _
                "VALUES (" & _
                " " & fInsertVariable(p_phy_inv_id, "N") & _
                "," & fInsertVariable(p_status_cd, "S") & _
                "," & fInsertVariable(p_opened_dt, "S") & _
                "," & fInsertVariable(p_closed_dt, "S") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                "," & fInsertVariable(o_trans_data.assoc_id, "N") & _
                "," & fInsertVariable("current_ts", "D") & _
                ")"
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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



Private Sub IssueUpdate()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
'init
    strRoutine_nm = "phy_inv.cls.IssueUpdate"
    
    strSql = "UPDATE phy_inv_order SET " & _
               " " & fUpdateVariable("status_cd", p_status_cd, "S", "=") & _
               " " & fUpdateVariable("closed_dt", p_closed_dt, "S", "=") & _
               "," & fUpdateVariable("updated_assoc_id", o_trans_data.assoc_id, "N", "=") & _
               "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
               " WHERE" & _
               " " & fWhereVariable("phy_inv_ord_id", p_phy_inv_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
    
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


Private Sub IssueDelete()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim dblUpdateQty As Double
    
'init
    strRoutine_nm = "message.cls.IssueDelete"
    
    strSql = "DELETE FROM phy_inv_order "
    strSql = strSql & "where "
    strSql = strSql & " " & fWhereVariable("phy_inv_ord_id", p_phy_inv_id, "N", "=")
    
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)

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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to change the purchase order status
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Public Sub ChangeStatus()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    
    ' Initialize
    strRoutine_nm = "phy_inv.cls.changeStatus"

    Select Case p_status_cd
        Case "CL"
            Call CloseOrder
        Case Else
    End Select
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If
 
    Call GetPhyInv
    
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



'//////////////////////////////////////////////////////////////////////////////
'Description:   Function to change the purchase order status to closed.
'  Arguments:   None
'
'//////////////////////////////////////////////////////////////////////////////
Private Sub CloseOrder()

    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim objPurchaseOrderItem As Object
    Dim dblUpdatedQty As Double
    Dim strSql As String
    Dim dblUpdateQty As Double
    
    ' Initialize
    strRoutine_nm = "phy_inv.cls.closeOrder"
       
     strSql = "UPDATE phy_inv_order " & _
            "  SET " & fUpdateVariable("status_cd", "CL", "S", "=") & _
                  "," & fUpdateVariable("closed_dt", "current_ts", "D", "=") & _
                  "," & fUpdateVariable("updated_assoc_id", o_trans_data.p_assoc_id, "N", "=") & _
                  "," & fUpdateVariable("updated_ts", "current_ts", "D", "=") & _
            " WHERE " & fWhereVariable("phy_inv_ord_id", p_phy_inv_id, "N", "=")
           
                 
    Call o_trans_data.ExecuteDatabase(strSql, dblUpdateQty, o_error)
        
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

Public Sub GetPhyInvs()
    
    On Error GoTo error_handler
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
       Exit Sub
    End If

    Dim strRoutine_nm As String
    Dim strSql As String
    Dim objRecordset As adodb.Recordset
    
'init
    strRoutine_nm = "phy_inv.cls.GetPhyInvs"

    strSql = " SELECT * " & _
             " FROM phy_inv_order " & _
             " WHERE phy_inv_ord_id is not null "
             
    If fValuePresent(p_status_cd) = True Then
        strSql = strSql & " AND " & fWhereVariable("status_cd", p_status_cd, "S", "=")
    End If
             
    If fValuePresent(p_start_entered_dt) = True Then
       If fValuePresent(p_end_entered_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("entered_dt", p_start_entered_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("entered_dt", p_end_entered_dt, "D", "<=") & ")"
       Else
            strSql = strSql & "  AND " & fWhereVariable("entered_dt", p_start_entered_dt, "D", ">=")
       End If
    End If
    
    If fValuePresent(p_start_closed_dt) = True Then
       If fValuePresent(p_end_closed_dt) = True Then
            strSql = strSql & "  AND(" & fWhereVariable("closed_dt", p_start_closed_dt, "D", ">=")
            strSql = strSql & " AND " & fWhereVariable("closed_dt", p_end_closed_dt, "D", "<=") & ")"
       Else
            strSql = strSql & "  AND " & fWhereVariable("closed_dt", p_start_closed_dt, "D", ">=")
       End If
    End If
            
    Call o_trans_data.OpenRecordset(objRecordset, strSql, o_error)
    
    If o_error.p_type_cd = "E" Or o_error.p_type_cd = "F" Then
        
        Set objRecordset = Nothing
        Exit Sub
    Else
        Set o_record_set = objRecordset
        Set objRecordset = Nothing
    End If
    
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
