VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private xSweet As Sweet
Private xIdentity As Long
Private xRawvalue As String
Private xEndvalue As String

Public Property Get Basename() As String
    Basename = "IObject"
End Property

Public Property Get Typename() As String
    Typename = "Mush"
End Property

Public Property Get Sweet() As Sweet
    Set Sweet = xSweet
End Property

Public Property Set Sweet(ByRef newVal As Sweet)
    Set xSweet = newVal
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
    If xEndvalue = "" Then
        Endvalue = xRawvalue
    Else
        Endvalue = xEndvalue
    End If
End Property

Public Property Let Endvalue(ByVal newVal As String)
    Endvalue = newVal
End Property

Private Sub Class_Initialize()

    xIdentity = GetNewIdentity
End Sub

Private Sub Class_Terminate()
    Set xSweet = Nothing
End Sub

