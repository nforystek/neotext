VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private m_Col As Collection

Public Property Get VBCollection() As Collection
    Set VBCollection = m_Col
End Property
Public Property Set VBCollection(ByRef newVal As Collection)
    Set m_Col = newVal
End Property

Public Sub Add(ByVal Data As String, Optional ByVal Key As Variant)
    If Not IsMissing(Key) Then
        m_Col.Add Data, Key
    Else
        m_Col.Add Data
    End If
End Sub

Public Sub Remove(ByVal Key As Variant)
    m_Col.Remove Key
End Sub

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Property Get Item(ByVal Key As Variant) As String
    Item = m_Col.Item(Key)
End Property

Public Function NewEnum() As IUnknown
    Set NewEnum = m_Col.[_NewEnum]
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



Attribute 