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

Private pAlpha As Single
Private pRed As Single
Private pGreen As Single
Private pBlue As Single
Public Event Changed()

Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property
Public Property Get RGBA() As Long

    RGBA = D3DColorRGBA(red * 255, green * 255, blue * 255, Alpha)

End Property
Public Property Get ARGB() As Long

    ARGB = D3DColorARGB(Alpha, red * 255, green * 255, blue * 255)

End Property
Public Property Let RGB(ByVal RHS As Long)
    Dim r As Long
    Dim g As Long
    Dim b As Long
    
    ConvertColor CLng(RHS), r, g, b
    pRed = CSng(r) / 255
    pGreen = CSng(g) / 255
    pBlue = CSng(b) / 255
    
'    pRed = rgbRed(RHS) / 255
'    pGreen = rgbGreen(RHS) / 255
'    pBlue = rgbBlue(RHS) / 255
End Property
Public Property Get RGB() As Long
    If pRed > 1 Or pGreen > 1 Or pBlue > 1 Then
        RGB = VBA.RGB(pRed, pGreen, pBlue)
    Else
        RGB = VBA.RGB(pRed * 255, pGreen * 255, pBlue * 255)
    End If
End Property

Public Property Get Alpha() As Single
    Alpha = pAlpha
End Property
Public Property Let Alpha(ByVal RHS As Single)
    pAlpha = RHS
    RaiseEvent Changed
End Property
 
Public Property Get red() As Single
    red = pRed
End Property
Public Property Let red(ByVal RHS As Single)
    pRed = RHS
    RaiseEvent Changed
End Property
 
Public Property Get green() As Single
    green = pGreen
End Property
Public Property Let green(ByVal RHS As Single)
    pGreen = RHS
    RaiseEvent Changed
End Property
 
Public Property Get blue() As Single
    blue = pBlue
End Property
Public Property Let blue(ByVal RHS As Single)
    pBlue = RHS
    RaiseEvent Changed
End Property


Public Property Let ToString(ByVal RHS As String)
    RHS = Trim(RHS)
    If Left(RHS, 1) = "[" And Right(RHS, 1) = "]" Then
        RHS = Mid(RHS, 2, Len(RHS) - 2)
        pRed = ParseNumerical(RHS)
        If RHS <> "" Then
            pGreen = ParseNumerical(RHS)
            If RHS <> "" Then
                pBlue = ParseNumerical(RHS)
                If RHS <> "" Then
                    pAlpha = ParseNumerical(RHS)
                End If
            End If
        End If
    End If

    RaiseEvent Changed
End Property
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
Attribute ToString.VB_MemberFlags = "640"
    ToString = "[ " & pRed & " , " & pGreen & ", " & pBlue & ", " & pAlpha & "]"
End Property

Private Sub Class_Initialize()
    Stats_Color_Count = Stats_Color_Count + 1
    Alpha = 1
End Sub

Private Sub Class_Terminate()
    Stats_Color_Count = Stats_Color_Count - 1
End Sub
