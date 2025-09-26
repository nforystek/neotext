VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
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
    Icon = 0 'normal 2D mouse pointer no matter where it occurs, if the console is down, it acts as this mode
    Hide = 1 'hidden mouse upon mouse over with focus in window mode, can move mouse outsdie window showing the pointer
    Trap = 2 'hidden plus the mouse is trapped, meaning in window mode the mouse can not be moved outside the window
                'hidden and Trap are generally the same in fullscreen mode
End Enum

Private binds(0 To 255) As OnEvent
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

    Dim i As Long
    For i = 0 To 255
        If Not binds(i) Is Nothing Then
            If binds(i).RunScript <> "" Then
                ToString = ToString & Include.URLEncode("Set Bindings.Item(""" & GetBindingText(i) & """) = Include.NewObject(""OnEvent"")") & vbCrLf
                ToString = ToString & Include.URLEncode("Bindings.Item(""" & GetBindingText(i) & """).RunScript = Include.URLDecode(""" & Include.URLEncode(binds(i).RunScript) & """)") & vbCrLf
                ToString = ToString & Include.URLEncode("Bindings.Item(""" & GetBindingText(i) & """).AppliesTo = " & binds(i).AppliesTo) & vbCrLf
                ToString = ToString & Include.URLEncode("Bindings.Item(""" & GetBindingText(i) & """).Behavior = " & binds(i).Behavior) & vbCrLf
            End If
        End If
    Next
    ToString = ToString & Include.URLEncode("Bindings.MouseInput=" & MouseInput) & vbCrLf
    
    ToString = "  <Bindings>" & vbCrLf & ToString & vbCrLf & "</Bindings>" & vbCrLf
    
End Function


Public Property Get Count() As Long
    Count = 256
End Property

Public Property Get Item(ByVal Index As Variant) As OnEvent
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    If TypeName(Index) = "String" Then
        Index = CInt(GetBindingIndex(Index))
    End If
    If IsNumeric(Index) Then
        If (Index > -1) Then
            Set Item = binds(Index)
        End If
    End If
End Property

Public Property Set Item(ByVal Index As Variant, ByRef RHS As OnEvent)
    If TypeName(Index) = "String" Then
        Index = CInt(GetBindingIndex(Index))
    End If
    If IsNumeric(Index) Then
        If (Index > -1) Then
            Set binds(Index) = RHS
        End If
    End If
End Property

Private Sub Class_Initialize()
    Stats_Bindings_Count = Stats_Bindings_Count + 1
End Sub

Private Sub Class_Terminate()
    Stats_Bindings_Count = Stats_Bindings_Count - 1
End Sub
