VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pForm As MotionTypes

Private pKey As String
Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
'    If Motions.Exists(pKey) And pKey <> "" Then Motions.Remove pKey
'    pKey = RHS
'    If RHS <> "" And (Not Motions.Exists(RHS)) Then Motions.Add Me, RHS
    
'    If RHS = "" And Motions.Exists(pKey) Then
'        Motions.Remove pKey
'        pKey = ""
'    Else
        If pKey <> "" And (Not Motions.Exists(RHS)) Then
           ' If All.Exists(pKey) Then All.Remove pKey
            If Motions.Exists(pKey) Then Motions.Remove pKey
           ' All.Add Me, RHS
            Motions.Add Me, RHS
            pKey = RHS
        ElseIf pKey = "" Then
            pKey = RHS
        End If
'    End If
End Property

Public Property Get Form() As Long
    Form = pForm
End Property
Public Property Let Form(ByRef RHS As Long)
    pForm = RHS
End Property

Private Sub Class_Terminate()
    If pKey <> "" Then
        If Motions.Exists(pKey) Then Motions.Remove pKey
      '  If All.Exists(pKey) Then All.Remove pKey
    End If
End Sub
