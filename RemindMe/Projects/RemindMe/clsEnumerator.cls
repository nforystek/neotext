VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public EnumValues As Collection

Public EnumeratorName As String

Public Sub AddEnumValue(ByVal EnumName As String, ByVal EnumValue As String)
    Dim newEnum As New clsEnumValue
    newEnum.EnumName = EnumName
    newEnum.EnumValue = EnumValue
    EnumValues.Add newEnum, EnumName
    Set newEnum = Nothing
End Sub
Public Sub ClearEnumValues()
    Do Until EnumValues.Count = 0
        EnumValues.Remove 1
    Loop
End Sub

Private Sub Class_Initialize()
    Set EnumValues = New Collection
End Sub
Private Sub Class_Terminate()
    ClearEnumValues
    Set EnumValues = Nothing
End Sub

