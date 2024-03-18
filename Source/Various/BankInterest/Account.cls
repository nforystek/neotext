VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pUpTime As String

Private pAddress As Long

Public Property Get Number()
    Number = pAddress
End Property

Private Sub Class_Initialize()
    Dim Funds As Currency
    pAddress = LocalAlloc(LMEM_NODISCARD, LenB(Funds))
    LocalLock pAddress
End Sub

Private Sub Class_Terminate()
    LocalUnlock pAddress
    LocalFree pAddress
End Sub

