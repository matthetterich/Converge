VERSION 5.00
Begin VB.Form frmDWHExtract 
   Caption         =   "Data Warehouse Extract"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin VB.Label Label1 
         Caption         =   "The Data Warehouse is currently being Generated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmDWHExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private M_STR_CONVERGE_CONNECT_STRING As String
Private M_STR_DWH_CONNECT_STRING As String
'Private Const M_STR_LOAD_FILE_NM As String = "\\TIMCOTEST\TEMP\CONVERGE_DWH_LOAD.TXT"
Private Const M_STR_LOAD_FILE_NM As String = "\\tr-sql\dwhextract\CONVERGE_DWH_LOAD.TXT"
Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    
    Dim objSalesDetail As Object _
      , objInventoryDetail As Object _
      , objPurchasingDetail As Object _
      , objError As Object _
      , strText As String _
      , conDwh As ADODB.Connection _
      , conConverge As ADODB.Connection
    
    Dim convergeConnnectString As String
    convergeConnnectString = "" & _
            "Provider=SQLOLEDB; " & _
            "DRIVER=SQL Server; " & _
            "UID=Converge; " & _
            "PWD=volters; " & _
            "WSID=TR-SQL; " & _
            "SERVER=TR-SQL; " & _
            "DATABASE=converge; " & _
            "APP=Microsoft Data Access Components; " & _
            "Description=Converge Env Var "

    Dim datawarehouseConnectString As String
    datawarehouseConnectString = "" & _
            "Provider=SQLOLEDB; " & _
            "DRIVER=SQL Server; " & _
            "UID=Converge; " & _
            "PWD=volters; " & _
            "WSID=TR-SQL; " & _
            "SERVER=TR-SQL; " & _
            "DATABASE=datawarehouse; " & _
            "APP=Microsoft Data Access Components; " & _
            "Description=Converge Env Var "

    M_STR_CONVERGE_CONNECT_STRING = convergeConnnectString
    M_STR_DWH_CONNECT_STRING = datawarehouseConnectString
    Call Me.Show
    Me.cmdExit.Visible = False
    Me.Refresh
    Me.MousePointer = vbHourglass
    Set objError = CreateObject("converge.error")
    Set objSalesDetail = CreateObject("converge_dwh.clsSalesDetail")
    Set objInventoryDetail = CreateObject("converge_dwh.clsInventoryDetail")
    Set objPurchasingDetail = CreateObject("converge_dwh.clsPurchasingDetail")
    Set conDwh = New ADODB.Connection
    Set conConverge = New ADODB.Connection
    
    Call conDwh.Open(M_STR_DWH_CONNECT_STRING)
    Call conConverge.Open(M_STR_CONVERGE_CONNECT_STRING)
    
    Call objSalesDetail.extract(objError, conDwh, conConverge, M_STR_LOAD_FILE_NM)
    Call objInventoryDetail.extract(objError, conDwh, conConverge, M_STR_LOAD_FILE_NM)
    Call objPurchasingDetail.extract(objError, conDwh, conConverge, M_STR_LOAD_FILE_NM)
        
    If objError.p_type_cd = "E" Or objError.p_type_cd = "F" Then
        MsgBox objError.p_desc
        Me.Label1.Caption = "The Data Warehouse Generator Failed, please contact Systems Support"
    End If
    
    If objError.p_type_cd = "E" Or objError.p_type_cd = "F" Then
        Me.cmdExit.Visible = True
        Set objSalesDetail = Nothing
        Set objError = Nothing
        Call conDwh.Close
        Call conConverge.Close
        Set conDwh = Nothing
        Set conConverge = Nothing
    Else
'        Call sendEmail("Yesterday's Sales Results", getEmailTx(conConverge))
        Set objSalesDetail = Nothing
        Set objError = Nothing
        Call conDwh.Close
        Call conConverge.Close
        Set conDwh = Nothing
        Set conConverge = Nothing
        End
    End If
    Me.MousePointer = vbDefault

End Sub

'Private Sub sendEmail( _
'              ByVal v_strSubject As String _
'            , ByVal v_strText As String _
'            )

'    Dim objAspEmail As ASPEMAILLib.MailSender
'    Set objAspEmail = New ASPEMAILLib.MailSender
'    With objAspEmail
'        Call .AddAddress("MattHetterich@hotmail.com", "Matt Hetterich")
'        Call .AddAddress("MattHetterich@cinci.rr.com", "Matt Hetterich")
'        .From = "MattHetterich@Cinci.rr.com"
'        .Subject = "Converge Datawarehouse Load Confirmation"
'        .IsHTML = True
'        .Body = v_strText
'        .Username = "MattHetterich@Cinci.rr.com"
'        .Password = "mh021000"
'        .Host = "smtp-server.cinci.rr.com"
'        .Port = 25
'        Call .Send
'    End With
'
'    With objAspEmail
'        Call .AddAddress("johnkuz@timcorubber.com", "John Kuzmick")
'        Call .AddAddress("BKuzmick@timcorubber.com", "Bob Kuzmick")
'        Call .AddAddress("jkuzmick@timcorubber.com", "Jim Kuzmick")
'        Call .AddAddress("MattHetterich@cinci.rr.com", "Matt Hetterich")
'        Call .AddAddress("mhetterich@timcorubber.com", "Matt Hetterich")
''        Call .AddCC("jpaarlberg@timcorubber.com", "Jason Paarlberg")
'        .From = "MHetterich@TimcoRubber.com"
'        .Subject = "Converge Datawarehouse Load Confirmation"
'        .IsHTML = True
'        .Body = v_strText
'        .Username = "MHetterich@TimcoRubber.com"
'        .Password = "timco1"
'        .Host = "smtp.registeredsite.com"
'        .Port = 25
'         Call .Send
'    End With
'End Sub

Private Function getEmailTx(ByRef r_conConverge As ADODB.Connection) As String
    
    Dim objRecordset As ADODB.Recordset _
      , objInvRs As ADODB.Recordset _
      , strSqlTx As String _
      , strEmailTx As String _
      , dteSummaryDt As Date _
      , dblMarkUpPct As Double _
      , dblGrossMarginPct As Double _
      , dblGrossMarginAm As Double _
      , dblSalesPrice As Double _
      , dblSalesCost As Double
      

    dteSummaryDt = CDate(FormatDateTime(Now, vbShortDate))
    dteSummaryDt = dteSummaryDt - 1

    strSqlTx = "select sum(sales_price) as sales_price " & _
               " ,sum(sales_cost) as sales_cost " & _
               " ,sum(sales_profit_am) as sales_profit_am " & _
               " from dwh_sales_detail " & _
               " where sales_dt = '" & dteSummaryDt & "'"
    
    Set objRecordset = r_conConverge.Execute(strSqlTx)
    
    strSqlTx = "select sum(sales_price) as sales_price " & _
               " ,sum(sales_cost) as sales_cost " & _
               " ,sum(sales_profit_am) as sales_profit_am " & _
               " from dwh_sales_detail " & _
               " where sales_dt = '" & dteSummaryDt & "'"
    
    strSqlTx = _
            " SELECT inv_loc_id, SUM(inv_adj_qty * item_price) AS inv_value_am " _
            & " From dwh_inventory_detail " _
            & " GROUP BY inv_loc_id " _
            & " ORDER BY inv_loc_id "
            
    Set objInvRs = r_conConverge.Execute(strSqlTx)
    
    If objRecordset.EOF = True Then
        dblSalesPrice = 0
        dblSalesCost = 0
        dblGrossMarginAm = 0
        dblMarkUpPct = 0
        dblGrossMarginPct = 0
    Else
        If IsNull(objRecordset("sales_price")) = True Then
            dblSalesPrice = 0
        Else
            dblSalesPrice = objRecordset("sales_price")
        End If
        
        If IsNull(objRecordset("sales_cost")) = True Then
            dblSalesCost = 0
        Else
            dblSalesCost = objRecordset("sales_cost")
        End If
        
        If IsNull(objRecordset("sales_profit_am")) = True Then
            dblGrossMarginAm = 0
        Else
            dblGrossMarginAm = objRecordset("sales_profit_am")
        End If
        
        If dblSalesCost = 0 Then
            dblMarkUpPct = 0
        Else
            dblMarkUpPct = (dblSalesPrice - dblSalesCost) / dblSalesCost
        End If
        
        If dblSalesPrice = 0 Then
            dblGrossMarginPct = 0
        Else
            dblGrossMarginPct = (dblSalesPrice - dblSalesCost) / dblSalesPrice
        End If
    End If
    
    strEmailTx = _
        "<html><head>" & _
        "<meta name='author' content='Matt Hetterich'><meta name='description' content='Customer Maintanance'><title>Daily Sales Summary Email</title></head>" & _
        "<body bgcolor='#CCCC99' text='#000000'>" & _
        "<table width='100%'>" & _
        "<tr><td width='100%' align='left'>" & _
        "<font face='Arial' size='2'><b>The Sales and Inventory Datawarehouses were loaded successfully on " & FormatDateTime(Now, vbShortDate) & ". <BR>Here is yesterday's daily sales summary and close of business inventory value.</b></font>" & _
        "</td></tr>" & _
        "</table>" & _
        "<br>" & _
        " "
'    strEmailTx = strEmailTx & _
'        "<table border='0' width='400' cellpadding='0' cellspacing='0'>" & _
'        "<tr><td align='left' bgcolor='000080' valign='top' colspan='2'>" & _
'        "   &nbsp;<font face='Arial' size='2' color='#ffffff'><b>Daily Sales Summary for: " &(dteSummaryDt, "dddd") & ", " & dteSummaryDt & "</b></font>" & _
'        "</td></tr>" & _
'        " "
    strEmailTx = strEmailTx & _
        "<tr><td width='0' align='Left'>" & _
        "    <font face='Arial' size='2' color='#0000CC'><b>Total Sales Amount<b></font>" & _
        "</td><td align='right'>" & _
        " " & FormatCurrency(dblSalesPrice, 2) & _
        "</td></tr>" & _
        " "
    strEmailTx = strEmailTx & _
        "<tr><td width='0' align='Left'>" & _
        "    <font face='Arial' size='2' color='#0000CC'><b>Total Cost Amount<b></font>" & _
        "</td><td align='right'>" & _
        " " & FormatCurrency(dblSalesCost, 2) & _
        "</td></tr>" & _
        ""
    strEmailTx = strEmailTx & _
        "<tr><td width='0' align='Left'>" & _
        "    <font face='Arial' size='2' color='#0000CC'><b>Gross Margin Amount<b></font>" & _
        "</td><td align='right'>" & _
        " " & FormatCurrency(dblGrossMarginAm, 2) & _
        "</td></tr>" & _
        " "
    strEmailTx = strEmailTx & _
        "<tr><td width='0' align='Left'>" & _
        "    <font face='Arial' size='2' color='#0000CC'><b>Gross Margin Percent<b></font>" & _
        "</td><td align='right'>" & _
        " " & FormatPercent(dblGrossMarginPct, 2) & _
        "</td></tr>" & _
        " "
    strEmailTx = strEmailTx & _
        "<tr><td width='0' align='Left'>" & _
        "    <font face='Arial' size='2' color='#0000CC'><b>Mark Up Percent<b></font>" & _
        "</td><td align='right'>" & _
        " " & FormatPercent(dblMarkUpPct, 2) & _
        "</td></tr>" & _
        ""
    strEmailTx = strEmailTx & _
        "<tr><td valign='top' align='center' colspan=2>" & _
        "    <hr>" & _
        "</td></tr>" & _
        "" & _
        "</table>" & _
        "<br> "
    strEmailTx = strEmailTx & _
        "<table border='0' width='400' cellpadding='0' cellspacing='0'>" & _
        "<tr><td align='left' bgcolor='000080' valign='top' colspan='2'>" & _
        "   &nbsp;<font face='Arial' size='2' color='#ffffff'><b>Close of Business Inventory Value: " & Format(dteSummaryDt, "dddd") & ", " & dteSummaryDt & "</b></font>" & _
        "</td></tr>" & _
        " "
    Do While objInvRs.EOF = False
        strEmailTx = strEmailTx & _
            "<tr><td width='0' align='Left'>" & _
            "    <font face='Arial' size='2' color='#0000CC'><b>" & UCase(objInvRs("inv_loc_id")) & "<b></font>" & _
            "</td><td align='right'>" & _
            " " & FormatCurrency(objInvRs("inv_value_am"), 2) & _
            "</td></tr>" & _
            " "
        Call objInvRs.MoveNext
    Loop
    strEmailTx = strEmailTx & _
        "<tr><td valign='top' align='center' colspan=2>" & _
        "    <hr>" & _
        "</td></tr>" & _
        "" & _
        "</table>" & _
        "<br> "
    strEmailTx = strEmailTx & _
        "" & _
        "</body>" & _
        "</html>"

    getEmailTx = strEmailTx
    Call objRecordset.Close
    Call objInvRs.Close
End Function
