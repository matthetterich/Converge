VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_colProps As Collection

Public Sub putProp(ByVal v_strName As String, ByVal v_varValue As Variant)
    Call m_colProps.Add(v_varValue, UCase(v_strName))
End Sub

Public Function getProp(ByVal v_strName As String)
    getProp = m_colProps.Item(UCase(v_strName))
End Function

Private Sub Class_Initialize()
    Set m_colProps = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colProps = Nothing
End Sub
