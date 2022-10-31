VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Swap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private rtol As Collection
Private ppause As Byte

Public Property Get Action() As Byte
    Action = ppause
End Property
Public Property Let Action(ByVal RHS As Byte)
    ppause = RHS
End Property
Public Property Get RightToLeft() As Collection
    Set RightToLeft = rtol
End Property

Private Sub Class_Initialize()
    Set rtol = New Collection
End Sub

Private Sub Class_Terminate()
    Do Until rtol.Count = 0
        rtol.Remove 1
    Loop
    Set rtol = Nothing
End Sub

Public Function ToString() As String
    If ppause = 1 Then
        ToString = "Pause"
    ElseIf ppause = 2 Then
        ToString = "Break"
    ElseIf ppause = 3 Then
        ToString = "Stop"
    Else
        Dim itm As Variant
        ToString = "Swap"
        For Each itm In rtol
            ToString = ToString & " " & itm & ","
        Next
        If Right(ToString, 1) = "," Then ToString = Left(ToString, Len(ToString) - 1)
    End If
End Function
