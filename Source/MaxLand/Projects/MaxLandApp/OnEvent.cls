VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OnEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EventBehavior
    Rapid = 0
    Press = 1
    Locks = 2
End Enum


Private pBehavior As EventBehavior
Private pAppliesTo As Variant
Private pRunScript As String
Private pStartLine As String
Private pEventFlags As Integer


Friend Property Get EventFlags() As Integer
    EventFlags = pEventFlags
End Property
Friend Property Let EventFlags(ByVal RHS As Integer)
    pEventFlags = RHS
End Property
Friend Property Get StartLine() As String
    StartLine = pStartLine
End Property
Friend Property Let StartLine(ByVal RHS As String)
    pStartLine = RHS
End Property



Public Property Get Behavior() As EventBehavior
    Behavior = pBehavior
End Property
Public Property Let Behavior(ByVal RHS As EventBehavior)
    pBehavior = RHS
End Property
Public Property Get AppliesTo() As Variant
    AppliesTo = pAppliesTo
End Property
Public Property Let AppliesTo(ByVal RHS As Variant)
    pAppliesTo = RHS
End Property
Public Property Get RunScript() As String
    RunScript = pRunScript
End Property
Public Property Let RunScript(ByVal RHS As String)
    pRunScript = RHS
End Property

Private Sub Class_Initialize()
    pBehavior = 1
End Sub
