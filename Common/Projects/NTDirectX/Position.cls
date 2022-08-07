VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private pMacroscopic As Macroscopic

Friend Property Get Macroscopic() As Macroscopic
    Set Macroscopic = pMacroscopic
End Property
Friend Property Set Macroscopic(ByRef RHS As Macroscopic)
    Set pMacroscopic = RHS
End Property

Public Sub Serialize(ByVal Filename As String)
    Dim state As Boolean
    If Not (pMacroscopic Is Nothing) Then
        state = pMacroscopic.RenderState
        If state Then pMacroscopic.PauseRendering
    End If


    If state Then pMacroscopic.ResumeRendering
End Sub
Public Sub Deserialize(ByVal Filename As String)
    Dim state As Boolean
    If Not (pMacroscopic Is Nothing) Then
        state = pMacroscopic.RenderState
        If state Then pMacroscopic.PauseRendering
    End If


    If state Then pMacroscopic.ResumeRendering
End Sub

'Private Sub Class_Terminate()
'    If Not (pMacroscopic Is Nothing) Then
'        pMacroscopic.PauseRendering
'        Set pMacroscopic = Nothing
'    End If
'End Sub

