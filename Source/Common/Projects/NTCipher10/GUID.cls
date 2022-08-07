VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUID"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Static Function Generate() As String
    Generate = modGuid.GUID
End Function

Static Function Validate(ByVal CheckGuid) As Boolean
    Validate = modGuid.IsGuid(CheckGuid)
End Function

Private Static Sub Class_Initialize()
    GUID
End Sub

Private Static Sub Class_Terminate()
    GUID
End Sub
