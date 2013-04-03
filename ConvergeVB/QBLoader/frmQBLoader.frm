VERSION 5.00
Begin VB.Form frmQBLoader 
   Caption         =   "Quickbooks Batch Loader"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Text            =   "d:/Company Shared Folders/Quickbooks/TimcoRubberProducts1.qbw"
      Top             =   3600
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set QB DB Path"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Debug"
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Load Now"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   288
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   3240
      Top             =   120
   End
   Begin VB.Label Label6 
      Caption         =   "Prev wake up TS"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Last run TS"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Run Count"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Wake Up Ct"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Load Count"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Error Count"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmQBLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_lngCounter As Long
Dim m_objLoader As clsLoader
Dim m_objTransData As Object
Dim m_objError As Object


Private Sub cmdRun_Click()
    
    Dim loadCt As Long
    Dim ErrorCt As Long
    Me.Text4.Text = CLng(Me.Text4.Text) + 1
    Me.Text7.Text = Me.Text5.Text
    Me.Text5.Text = Now
    m_lngCounter = 121
    loadCt = CLng(Text2.Text)
    ErrorCt = 0
    Set m_objError = CreateObject("Converge.Error")
    If m_objLoader Is Nothing Then
        Set m_objLoader = New clsLoader
    End If
    Call m_objLoader.ProcessTrans(m_objError, m_objTransData, ErrorCt, loadCt)
    If Time < CDate("6:00 AM") And Time > CDate("6:00 PM") Then
        Set m_objLoader = Nothing
    End If
    Text1.Text = ErrorCt
    Text2.Text = loadCt
End Sub

Private Sub Command1_Click()
    Dim x As Variant
    
    
    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    'Call SaveSetting("CONVERGE", "QUICKBOOKS", "DBNAME", "C:/projects/TimcoQB/TimcoRubberProducts1.qbw")
    'Call SaveSetting("CONVERGE", "QUICKBOOKS", "DBNAME", "d:/Company Shared Folders/Quickbooks/TimcoRubberProducts1.qbw")
    Call SaveSetting("CONVERGE", "QUICKBOOKS", "DBNAME", Text9.Text)
    Text9.Text = GetSetting("CONVERGE", "QUICKBOOKS", "DBNAME", "")
End Sub

Private Sub Form_Load()
    
    Dim objConnection As ADODB.Connection
    Set m_objError = CreateObject("Converge.Error")
    Set m_objTransData = CreateObject("converge.Trans_data")
    Set objConnection = New ADODB.Connection
    m_lngCounter = 121
    Timer1.Enabled = True
    Call objConnection.Open("Provider=SQLOLEDB; DRIVER=SQL Server; UID=Converge; PWD=volters; WSID=TIMCOSERVER; APP=Converge; DATABASE=Converge; SERVER=(local)")
    Set m_objTransData.o_dbConnection = objConnection
    m_objTransData.p_ConnectString = objConnection.ConnectionString
    m_objTransData.p_update_fl = True
    Call LoadSystemTables(objConnection)
    Text9.Text = GetSetting("CONVERGE", "QUICKBOOKS", "DBNAME", "")
    Text1.Text = 0
    Text2.Text = 0
End Sub

Function LoadSystemTables(ByRef r_objConnection As ADODB.Connection)
    
    Dim strSql As String
    Dim rstCodeDesc As ADODB.Recordset
    Dim rstStateCode As ADODB.Recordset

    strSql = "SELECT * FROM code_desc ORDER by field_nm, seq_nbr "
    Set rstCodeDesc = CreateObject("ADODB.RecordSet")
    Call rstCodeDesc.Open(strSql, r_objConnection, adOpenStatic)
    m_objTransData.p_CodeDesc = rstCodeDesc.GetRows()

    strSql = "SELECT * FROM code_desc WHERE field_nm = 'state_cd' ORDER by seq_nbr "
    Set rstStateCode = CreateObject("ADODB.RecordSet")
    Call rstStateCode.Open(strSql, r_objConnection, adOpenStatic)
    m_objTransData.p_StateCode = rstStateCode.GetRows()

End Function

Private Sub Timer1_Timer()

    Dim rstData As ADODB.Recordset
    Dim lngRowCt As Long
    Dim strSqlTx As String
    Me.Text3.Text = CLng(Me.Text3.Text) + 1
    m_lngCounter = m_lngCounter + 1
    Me.Text8.Text = Me.Text6.Text
    Me.Text6.Text = Now
    Set m_objError = CreateObject("Converge.Error")
    If Time > CDate("6:00 AM") And Time < CDate("09:00 PM") Then
        strSqlTx = "select count(*) loadCt from quickbooks_trans  where err_msg is null "
        Call m_objTransData.OpenRecordset(rstData, strSqlTx, m_objError)
        
        If rstData.EOF = False Then
            lngRowCt = rstData("loadCt")
        End If
        Call rstData.Close
        Set rstData = Nothing
        
        If m_lngCounter >= 120 Or lngRowCt > 0 Then
            If m_objLoader Is Nothing Then
                Set m_objLoader = New clsLoader
            End If
            m_lngCounter = 0
            Dim loadCt As Long
            Dim ErrorCt As Long
            loadCt = CLng(Text2.Text)
            ErrorCt = 0
            Set m_objError = CreateObject("Converge.Error")
            Call m_objLoader.ProcessTrans(m_objError, m_objTransData, ErrorCt, loadCt)
            Set m_objError = CreateObject("Converge.Error")
            Me.Text1.Text = ErrorCt
            Me.Text2.Text = loadCt
            Me.Text4.Text = CLng(Me.Text4.Text) + 1
            Me.Text7.Text = Me.Text5.Text
            Me.Text5.Text = Now
        End If
    Else
        Set m_objLoader = Nothing
    End If

End Sub
