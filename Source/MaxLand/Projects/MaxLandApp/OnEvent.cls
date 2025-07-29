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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pAppliesTo As String
Private pRunMethod As String
Private pStartLine As String
Private pRunFlag As Boolean


Friend Property Get RunFlag() As Boolean
    RunFlag = pRunFlag
End Property
Friend Property Let RunFlag(ByVal RHS As Boolean)
    pRunFlag = RHS
End Property


Public Property Get AppliesTo() As String
    AppliesTo = pAppliesTo
End Property
Public Property Let AppliesTo(ByVal RHS As String)
    pAppliesTo = RHS
End Property
Public Property Get RunMethod() As String
    RunMethod = pRunMethod
End Property
Public Property Let RunMethod(ByVal RHS As String)
    pRunMethod = RHS
End Property

Friend Property Get StartLine() As String
    StartLine = pStartLine
End Property
Friend Property Let StartLine(ByVal RHS As String)
    pStartLine = RHS
End Property


