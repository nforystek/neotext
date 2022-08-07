VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private hProcess As Long
Private hThread As Long
Private dwProcessId As Long
Private dwThreadId As Long

Private pPath As String

Friend Property Get Process() As Long
    Process = hProcess
End Property
Friend Property Let Process(ByVal newValue As Long)
    hProcess = newValue
End Property
Friend Property Get Thread() As Long
    Thread = hThread
End Property
Friend Property Let Thread(ByVal newValue As Long)
    hThread = newValue
End Property
Friend Property Get ProcessId() As Long
    ProcessId = dwProcessId
End Property
Friend Property Let ProcessId(ByVal newValue As Long)
    dwProcessId = newValue
End Property
Friend Property Get ThreadId() As Long
    ThreadId = dwThreadId
End Property
Friend Property Let ThreadId(ByVal newValue As Long)
    dwThreadId = newValue
End Property

Friend Property Get Path() As String
    Path = pPath
End Property
Friend Property Let Path(ByVal newValue As String)
    pPath = newValue
End Property


