VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormHWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Compare Binary

Private propName As String
Private propID As Long
Private prophWnd As Long
Public Property Get hWnd() As Long
    hWnd = prophWnd
End Property
Public Property Let hWnd(ByVal RHS As Long)
    prophWnd = RHS
End Property
Public Property Get ProcessName() As String
    ProcessName = propName
End Property
Public Property Let ProcessName(ByVal RHS As String)
    propName = RHS
End Property
Public Property Get ProcessID() As Long
    ProcessID = propID
End Property
Public Property Let ProcessID(ByVal RHS As Long)
    propID = RHS
End Property

Public Property Get Visible() As Boolean

End Property

Public Property Let Visible(ByVal newVal As Boolean)

End Property

Public Sub Hide()

End Sub

