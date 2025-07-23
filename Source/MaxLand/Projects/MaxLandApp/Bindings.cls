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

Public Enum MouseControl
    Visual = 0 'no mouse conduct
    Hidden = 1 'hidden mouse upon mouse over with focus
    Trapping = 2 'hidden plus the mouse is trappable/untrappable when clicked/esc button is pressed
End Enum

Private binds(0 To 255) As String
Private pMouseInput As MouseControl
Private pSerialize As Boolean
Public Property Get MouseInput() As MouseControl
    MouseInput = pMouseInput
End Property
Public Property Let MouseInput(ByVal RHS As MouseControl)
    pMouseInput = RHS
End Property
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Function ToString() As String
Attribute ToString.VB_MemberFlags = "40"

    Dim I As Long
    For I = 0 To 255
        If binds(I) <> "" Then
            ToString = ToString & Include.URLEncode("Bindings.Item(""" & GetBindingText(I) & """) = Include.URLDecode(""" & Include.URLEncode(binds(I)) & """)") & vbCrLf
        End If
    Next
    ToString = ToString & Include.URLEncode("Bindings.MouseInput=" & MouseInput) & vbCrLf
    
    ToString = "  <Bindings>" & vbCrLf & ToString & vbCrLf & "</Bindings>" & vbCrLf
    
End Function


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
            Item = MouseInput
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
            MouseInput = RHS
        End If
    End If
End Property


Private Sub Class_Initialize()
    Stats_Bindings_Count = Stats_Bindings_Count + 1
    
End Sub

Private Sub Class_Terminate()
    Stats_Bindings_Count = Stats_Bindings_Count - 1
End Sub
