VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private m_varErr_cd As Variant
Private m_varErr_desc As Variant
Private m_varMessage_id As Variant
Private m_varErr_nbr As Variant
Private m_varRoutine_nm As Variant
Private m_varType_cd As Variant
Private m_varSqlstate_cd As Variant
Private m_varSqlcode_cd As Variant
Private m_varSql_tx As Variant

Public Property Let p_type_cd(ByVal vData As Variant)
    m_varType_cd = vData
End Property

Public Property Get p_type_cd() As Variant
    p_type_cd = m_varType_cd
End Property

Public Property Let type_cd(ByVal vData As Variant)
    m_varType_cd = vData
End Property

Public Property Get type_cd() As Variant
    type_cd = m_varType_cd
End Property

Public Property Let p_routine_nm(ByVal vData As Variant)
    m_varRoutine_nm = vData
End Property

Public Property Get p_routine_nm() As Variant
    p_routine_nm = m_varRoutine_nm
End Property

Public Property Let routine_nm(ByVal vData As Variant)
    m_varRoutine_nm = vData
End Property

Public Property Get routine_nm() As Variant
    routine_nm = m_varRoutine_nm
End Property

Public Property Let p_nbr(ByVal vData As Variant)
    m_varErr_nbr = vData
End Property

Public Property Get p_nbr() As Variant
    p_nbr = m_varErr_nbr
End Property

Public Property Let err_nbr(ByVal vData As Variant)
    m_varErr_nbr = vData
End Property

Public Property Get err_nbr() As Variant
    err_nbr = m_varErr_nbr
End Property

Public Property Let p_message_id(ByVal vData As Variant)
    m_varMessage_id = vData
End Property

Public Property Get p_message_id() As Variant
    p_message_id = m_varMessage_id
End Property

Public Property Let message_id(ByVal vData As Variant)
    m_varMessage_id = vData
End Property

Public Property Get message_id() As Variant
    message_id = m_varMessage_id
End Property

Public Property Let p_desc(ByVal vData As Variant)
    m_varErr_desc = vData
End Property

Public Property Get p_desc() As Variant
    p_desc = m_varErr_desc
End Property

Public Property Let err_desc(ByVal vData As Variant)
    m_varErr_desc = vData
End Property

Public Property Get err_desc() As Variant
    err_desc = m_varErr_desc
End Property

Public Property Let p_err_cd(ByVal vData As Variant)
    m_varErr_cd = vData
End Property

Public Property Get p_err_cd() As Variant
    p_err_cd = m_varErr_cd
End Property

Public Property Let err_cd(ByVal vData As Variant)
    m_varErr_cd = vData
End Property

Public Property Get err_cd() As Variant
    err_cd = m_varErr_cd
End Property

Public Property Let p_sqlstate_cd(ByVal vData As Variant)
    m_varSqlstate_cd = vData
End Property

Public Property Get p_sqlstate_cd() As Variant
    p_sqlstate_cd = m_varSqlstate_cd
End Property

Public Property Let sqlstate_cd(ByVal vData As Variant)
    m_varSqlstate_cd = vData
End Property

Public Property Get sqlstate_cd() As Variant
    sqlstate_cd = m_varSqlstate_cd
End Property

Public Property Let p_sqlcode_cd(ByVal vData As Variant)
    m_varSqlcode_cd = vData
End Property

Public Property Get p_sqlcode_cd() As Variant
    p_sqlcode_cd = m_varSqlcode_cd
End Property

Public Property Let sqlcode_cd(ByVal vData As Variant)
    m_varSqlcode_cd = vData
End Property

Public Property Get sqlcode_cd() As Variant
    sqlcode_cd = m_varSqlcode_cd
End Property



Public Property Let p_sql_tx(ByVal vData As Variant)
    m_varSql_tx = vData
End Property

Public Property Get p_sql_tx() As Variant
    p_sql_tx = m_varSql_tx
End Property

Public Property Let sql_tx(ByVal vData As Variant)
    m_varSql_tx = vData
End Property

Public Property Get sql_tx() As Variant
    sql_tx = m_varSql_tx
End Property


Public Sub Clear()

    err_cd = Empty
    err_desc = Empty
    message_id = Empty
    err_nbr = Empty
    routine_nm = Empty
    type_cd = Empty
    sqlstate_cd = Empty
    sqlcode_cd = Empty
    sql_tx = Empty
    
End Sub

