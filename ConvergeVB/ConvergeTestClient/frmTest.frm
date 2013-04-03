VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim objTransaction As Object
    Set objTransaction = CreateObject("converge.transaction")
    Set objTransaction.o_trans_data = g_objTransData
    
    Set objTransaction.o_trans_object = CreateObject("converge_purchase.supplier")
    objTransaction.o_trans_object.p_supplier_nbr = "BH"
    objTransaction.o_trans_object.p_supplier_id = Null

' set transaction properties
    g_objError.Clear
    Set objTransaction.o_error = g_objError
    objTransaction.o_trans_data.p_origin_type_cd = "OL"
    objTransaction.o_trans_data.p_origin_nm = G_STR_PROJECT_NM
    objTransaction.o_trans_data.p_ConnectString = g_objConnection.ConnectionString
    objTransaction.o_trans_data.p_assoc_id = 2007
    objTransaction.o_trans_data.p_Function_cd = "GetSupplier"
    objTransaction.o_trans_data.p_origin_type_cd = "OL"
    objTransaction.o_trans_data.p_origin_nm = G_STR_PROJECT_NM
    objTransaction.o_trans_data.p_ConnectString = g_objConnection.ConnectionString
    objTransaction.o_trans_data.p_assoc_id = 2007
        objTransaction.ProcessTrans

End Sub


Private Sub Command2_Click()
    Dim objTransaction As Object
    Set objTransaction = CreateObject("converge.transaction")
    Set objTransaction.o_trans_data = g_objTransData
    
    Set objTransaction.o_trans_object = CreateObject("converge_sales.sales_order")
    objTransaction.o_trans_object.p_sales_ord_id = 21949
    objTransaction.o_trans_object.p_updated_assoc_id = 2007

' set transaction properties
    g_objError.Clear
    Set objTransaction.o_error = g_objError
    objTransaction.o_trans_data.p_origin_type_cd = "OL"
    objTransaction.o_trans_data.p_origin_nm = G_STR_PROJECT_NM
    objTransaction.o_trans_data.p_ConnectString = g_objConnection.ConnectionString
    objTransaction.o_trans_data.p_assoc_id = 2007
    Dim action
    action = "Plan Order"
    Select Case action
        Case "Open Order"
            objTransaction.o_trans_data.p_Function_cd = "ChangeSalesOrderStatus"
            objTransaction.o_trans_object.p_status_cd = "OP"
        Case "Hold Order"
            objTransaction.o_trans_data.p_Function_cd = "ChangeSalesOrderStatus"
            objTransaction.o_trans_object.p_status_cd = "OH"
        Case "Close Order"
            objTransaction.o_trans_data.p_Function_cd = "ChangeSalesOrderStatus"
            objTransaction.o_trans_object.p_status_cd = "CL"
        Case "Cancel Order"
            objTransaction.o_trans_data.p_Function_cd = "ChangeSalesOrderStatus"
            objTransaction.o_trans_object.p_status_cd = "CA"
        Case "Plan Order"
            objTransaction.o_trans_data.p_Function_cd = "AddDropShipPo"
        Case "Add Sub Sales Ord"
            objTransaction.o_trans_data.p_Function_cd = "AddSubSalesOrder"
    End Select

' process transaction
    objTransaction.o_trans_data.update_fl = True
    Call objTransaction.ProcessTrans
    
'if no error occured, commit unit of work
    If objTransaction.o_error.type_cd = "E" Or objTransaction.o_error.type_cd = "F" Then
        Set objTransaction.o_trans_object = Nothing
    Else
        Set objTransaction.o_trans_object = Nothing
    End If

End Sub
