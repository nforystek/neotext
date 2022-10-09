VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Private m_Col As Collection

Public Function Add(ByVal id As Long) As MaxService.clsOperation
    Dim opr As New MaxService.clsOperation 'New MaxService.clsOperation
    'Set opr = CreateObject("MaxService.MaxService.clsOperation")
        
    opr.id = id

    m_Col.Add opr, "O" & id

    Set Add = opr
End Function

Public Function Count() As Long
    Count = m_Col.Count
End Function

Public Function Item(ByVal Index As Variant) As MaxService.clsOperation
    Set Item = m_Col.Item(Index)
End Function


Public Sub Clear()
    Do While m_Col.Count > 0
        m_Col.Remove 1
    Loop
End Sub

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set m_Col = Nothing
End Sub
