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

Private pAlpha As Integer
Private pRed As Integer
Private pGreen As Integer
Private pBlue As Integer
Public Event Changed()

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property

Public Property Get RGB() As Long
    RGB = VBA.RGB(pRed, pGreen, pBlue)
End Property
Public Property Get Alpha() As Integer
    Alpha = pAlpha
End Property
Public Property Let Alpha(ByVal RHS As Integer)
    pAlpha = RHS
    RaiseEvent Changed
End Property
 
Public Property Get Red() As Integer
    Red = pRed
End Property
Public Property Let Red(ByVal RHS As Integer)
    pRed = RHS
    RaiseEvent Changed
End Property
 
Public Property Get Green() As Integer
    Green = pGreen
End Property
Public Property Let Green(ByVal RHS As Integer)
    pGreen = RHS
    RaiseEvent Changed
End Property
 
Public Property Get Blue() As Integer
    Blue = pBlue
End Property
Public Property Let Blue(ByVal RHS As Integer)
    pBlue = RHS
    RaiseEvent Changed
End Property


Public Property Let ToString(ByVal RHS As String)
    If InStr(RHS, "[") > 0 Then RHS = RemoveQuotedArg(RHS, "[", "]")
    pAlpha = CSng(RemoveNextArg(RHS, ","))
    pRed = CSng(RemoveNextArg(RHS, ","))
    pGreen = CSng(RemoveNextArg(RHS, ","))
    pBlue = CSng(RemoveNextArg(RHS, ","))
    RaiseEvent Changed
End Property
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "600"
    ToString = "[ " & pAlpha & " , " & pRed & ", " & pGreen & ", " & pBlue & "]"
End Property
