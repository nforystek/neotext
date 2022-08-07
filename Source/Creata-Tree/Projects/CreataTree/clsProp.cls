VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private nKey As String
Private nValue As Variant
Private nCast As String

Friend Property Get Key(Optional ByVal ImageKey As Boolean = False) As String
    If ImageKey Then
        Key = IIf((nValue = vbNullString), "none", SafeKey(nValue))
    Else
        Key = nKey
    End If
End Property

Friend Property Get Value() As Variant
    Select Case Cast
        Case "b"
            Value = CBool(nValue)
        Case "s", "d", "i"
            Value = CStr(nValue)
        Case Else
            Value = nValue
    End Select
End Property
Friend Property Let Value(ByVal NewValue As Variant)
    Select Case nCast
        Case "d", "b"
            Select Case LCase(Trim(NewValue))
                Case "true", "yes"
                    nValue = True
                    nCast = "b"
                Case "false", "no"
                    nValue = False
                    nCast = "b"
                Case "default"
                    nValue = "Default"
                    nCast = "d"
            End Select
        Case "s", "i"
            nValue = CStr(NewValue)
        Case Else
            nValue = NewValue
    End Select
End Property

Public Property Get XMLText() As String
    Select Case nCast
        Case "i"
            XMLText = "<img" & FormalWord(nKey) & ">" & URLEncode(CStr(nValue)) & "</img" & FormalWord(nKey) & ">"
        Case "b"
            XMLText = "<" & FormalWord(nKey) & ">" & IIf(nValue, "True", "False") & "</" & FormalWord(nKey) & ">"
        Case Else
            XMLText = "<" & FormalWord(nKey) & ">" & URLEncode(CStr(nValue)) & "</" & FormalWord(nKey) & ">"
    End Select
End Property

Friend Property Get Cast() As String
    Cast = nCast
End Property

Friend Function SetProperty(ByVal cKey As String, ByVal cValue As Variant)
    If Left(cKey, 3) = "img" Then
        nCast = "i"
        nKey = Mid(cKey, 4)
        nValue = CStr(cValue)
    Else
        nKey = cKey
        Select Case LCase(Trim(cValue))
            Case "false", "no"
                nCast = "b"
                nValue = False
            Case "true", "yes"
                nCast = "b"
                nValue = True
            Case "default"
                nCast = "d"
                nValue = CStr(cValue)
            Case Else
                nCast = "s"
                nValue = CStr(cValue)
        End Select
    End If
End Function

