VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Push"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private xColor As Color
Private xIdentity As Long
Private xRawvalue As String
Private xEndvalue As String

Private sMushes As Mushes
Private sPushes As Pushes
Private sRushes As Rushes

Public Property Get Basename() As String
    Basename = "IObject"
End Property

Public Property Get Typename() As String
    Typename = "Push"
End Property

Public Property Get Color() As Color
    Set Color = xColor
End Property

Public Property Set Color(ByRef newVal As Color)
    Set xColor = newVal
End Property

Public Property Get Identity() As Long
    Identity = xIdentity
End Property

Public Property Let Identity(ByVal newVal As Long)
    xIdentity = newVal
End Property

Public Static Property Get Rawvalue() As String
    Rawvalue = xRawvalue
End Property

Public Property Let Rawvalue(ByVal newVal As String)
    xRawvalue = newVal
End Property

Public Static Property Get Endvalue() As String
    Endvalue = xEndvalue
End Property

Public Property Let Endvalue(ByVal newVal As String)
    xEndvalue = newVal
End Property

Public Property Get Mushes() As Mushes
    Set Mushes = sMushes
End Property
Public Property Get Pushes() As Pushes
    Set Pushes = sPushes
End Property
Public Property Get Rushes() As Rushes
    Set Rushes = sRushes
End Property

Private Sub Class_Initialize()

    xIdentity = GetNewIdentity
    Set sPushes = New Pushes
    Set sMushes = New Mushes
    Set sRushes = New Rushes
End Sub

Private Sub Class_Terminate()
    Set sPushes = Nothing
    Set sMushes = Nothing
    Set sRushes = Nothing
    Set xColor = Nothing
End Sub
