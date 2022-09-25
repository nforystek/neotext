VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Color"
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
Private xNumerics As Long
Private xRed As Long
Private xGreen As Long
Private xBlue As Long
Private xIndex As Long

Public Property Get Basename() As String
    Basename = "IObject"
End Property

Public Property Get Typename() As String
    Typename = "Color"
End Property

Public Property Get Index() As Long
    Index = xIndex
End Property
Public Property Let Index(ByVal newVal As Long)
    xIndex = newVal
End Property

Public Property Get Numerics() As Long
    Numerics = xNumerics
End Property
Public Property Let Numerics(ByVal newVal As Long)
    xNumerics = newVal
End Property

Public Property Get Red() As Long
    Red = xRed
End Property
Public Property Let Red(ByVal newVal As Long)
    xRed = newVal
End Property
Public Property Get Green() As Long
    Green = xGreen
End Property
Public Property Let Green(ByVal newVal As Long)
   xGreen = newVal
End Property
Public Property Get Blue() As Long
    Blue = xBlue
End Property
Public Property Let Blue(ByVal newVal As Long)
    xBlue = newVal
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

Public Property Get Identity() As Long
    Identity = xIdentity
End Property

Public Property Let Identity(ByVal newVal As Long)
    xIdentity = newVal
End Property



