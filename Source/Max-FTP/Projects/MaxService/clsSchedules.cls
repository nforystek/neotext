VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private m_Col As Collection

Public Function Add(ByVal id As Long, ByVal Name As String) As MaxService.clsSchedule
    Dim sch As New MaxService.clsSchedule
    
    sch.id = id
    sch.Name = Name
    
    m_Col.Add sch, "S" & id
    Set Add = sch
End Function

Public Sub Remove(ByVal Index As Variant)
    m_Col.Remove Index
End Sub

Public Function Count() As Long
    Count = m_Col.Count
End Function

Public Function Item(ByVal Index As Variant) As MaxService.clsSchedule
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
