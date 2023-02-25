VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

'    Visual = 0 'no mouse conduct
'    Hidden = 1 'hidden mouse upon mouse over with focus
'    Trapping = 2 'hidden plus the mouse is trappable/untrappable with esc

Private binds(0 To 255) As String
Public Controller As Long 'ControllerModes
Private pSerialize As Boolean
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get ToString() As String

    Dim i As Long
    For i = 0 To 255
        If binds(i) <> "" Then
            ToString = ToString & Include.URLEncode("Bindings.Item(""" & GetBindingText(i) & """) = Include.URLDecode(""" & Include.URLEncode(binds(i)) & """)") & vbCrLf
        End If
    Next
    ToString = ToString & Include.URLEncode("Bindings.Controller=" & Controller) & vbCrLf
    
    ToString = "  <Bindings>" & vbCrLf & ToString & vbCrLf & "</Bindings>" & vbCrLf
    
End Property


Public Property Get Count() As Long
    Count = 256
End Property

Public Property Get Item(ByVal Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    If TypeName(Index) = "String" Then
        Index = CInt(GetBindingIndex(Index))
    End If
    If IsNumeric(Index) Then
        If (Index > -1) Then
            Item = binds(Index)
        ElseIf (Index = -2) Then
            Item = Controller
        End If
    End If
End Property

Public Property Let Item(ByVal Index As Variant, ByVal RHS As Variant)
    If TypeName(Index) = "String" Then
        Index = CInt(GetBindingIndex(Index))
    End If
    If IsNumeric(Index) Then
        If (Index > -1) Then
            binds(Index) = RHS
        ElseIf (Index = -2) Then
            Controller = RHS
        End If
    End If
End Property
