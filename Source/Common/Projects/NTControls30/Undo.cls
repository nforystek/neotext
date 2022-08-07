VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Undo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPriorTextData As Strands
Private pAfterTextData As Strands
Private pPriorSelRange As RangeType
Private pAfterSelRange As RangeType

Public Property Get PriorTextData() As Strands
    Set PriorTextData = pPriorTextData
End Property
Public Property Set PriorTextData(ByRef RHS As Strands)
    Set pPriorTextData = RHS
End Property

Public Property Get AfterTextData() As Strands
    Set AfterTextData = pAfterTextData
End Property
Public Property Set AfterTextData(ByRef RHS As Strands)
    Set pAfterTextData = RHS
End Property

Public Property Get PriorSelRange() As RangeType
    PriorSelRange = pPriorSelRange
End Property
Public Property Let PriorSelRange(ByRef RHS As RangeType)
    pPriorSelRange = RHS
End Property

Public Property Get AfterSelRange() As RangeType
    AfterSelRange = pAfterSelRange
End Property
Public Property Let AfterSelRange(ByRef RHS As RangeType)
    pAfterSelRange = RHS
End Property

Private Sub Class_Initialize()
    Set pPriorTextData = New Strands
    Set pAfterTextData = New Strands
End Sub

Private Sub Class_Terminate()
    Set pPriorTextData = Nothing
    Set pAfterTextData = Nothing
End Sub
